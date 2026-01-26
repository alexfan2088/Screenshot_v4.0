import AVFoundation
import CoreMedia
import CoreVideo
import Foundation
import ScreenCaptureKit

struct RecorderConfig {
    let outputMp4: URL
    let outputWav: URL
    let fps: Int
    let width: Int
    let height: Int
    let left: Int
    let top: Int
    let displayId: UInt32
    let captureCurrentSpaceOnly: Bool
    let includeVideo: Bool
    let includeAudio: Bool
    let audioMode: String
}

final class RecorderService: NSObject, SCStreamOutput, SCStreamDelegate {
    private let config: RecorderConfig
    private var stream: SCStream?
    private var assetWriter: AVAssetWriter?
    private var videoInput: AVAssetWriterInput?
    private var audioInput: AVAssetWriterInput?
    private var adaptor: AVAssetWriterInputPixelBufferAdaptor?
    private var audioFile: AVAudioFile?
    private var sessionStarted = false
    private var startTime: CMTime?

    init(config: RecorderConfig) {
        self.config = config
        super.init()
    }

    func start() async throws {
        if config.audioMode != "native" {
            throw NSError(domain: "RecorderHelper", code: 2, userInfo: [NSLocalizedDescriptionKey: "audioMode only supports native for now"])
        }

        let content = try await SCShareableContent.excludingDesktopWindows(false, onScreenWindowsOnly: true)
        let display = content.displays.first { $0.displayID == config.displayId } ?? content.displays.first
        guard let display else {
            throw NSError(domain: "RecorderHelper", code: 1, userInfo: [NSLocalizedDescriptionKey: "No display available"])
        }

        let filter: SCContentFilter
        if config.captureCurrentSpaceOnly {
            filter = SCContentFilter(display: display, excludingWindows: [])
        } else {
            let windows = content.windows.filter { $0.isOnScreen }
            if windows.isEmpty {
                filter = SCContentFilter(display: display, excludingWindows: [])
            } else {
                filter = SCContentFilter(display: display, including: windows)
                if #available(macOS 14.2, *) {
                    filter.includeMenuBar = true
                }
            }
        }
        let configuration = SCStreamConfiguration()
        configuration.width = config.width
        configuration.height = config.height
        configuration.minimumFrameInterval = CMTime(value: 1, timescale: CMTimeScale(config.fps))
        configuration.capturesAudio = config.includeAudio
        configuration.sampleRate = 48000
        configuration.channelCount = 2
        if config.width > 0 && config.height > 0 && (config.left != 0 || config.top != 0) {
            configuration.sourceRect = CGRect(x: config.left, y: config.top, width: config.width, height: config.height)
        }

        let stream = SCStream(filter: filter, configuration: configuration, delegate: self)
        try stream.addStreamOutput(self, type: .screen, sampleHandlerQueue: DispatchQueue(label: "recorder.video"))
        try stream.addStreamOutput(self, type: .audio, sampleHandlerQueue: DispatchQueue(label: "recorder.audio"))

        try prepareAssetWriter()

        self.stream = stream
        try await stream.startCapture()
    }

    func stop() async throws {
        try await stream?.stopCapture()

        if let writer = assetWriter {
            videoInput?.markAsFinished()
            audioInput?.markAsFinished()

            await withCheckedContinuation { continuation in
                writer.finishWriting {
                    continuation.resume()
                }
            }
        }

        audioFile = nil
        stream = nil
    }

    private func prepareAssetWriter() throws {
        if FileManager.default.fileExists(atPath: config.outputMp4.path) {
            try FileManager.default.removeItem(at: config.outputMp4)
        }
        if FileManager.default.fileExists(atPath: config.outputWav.path) {
            try FileManager.default.removeItem(at: config.outputWav)
        }

        let writer = try AVAssetWriter(outputURL: config.outputMp4, fileType: .mp4)

        if config.includeVideo {
            let videoSettings: [String: Any] = [
                AVVideoCodecKey: AVVideoCodecType.h264,
                AVVideoWidthKey: config.width,
                AVVideoHeightKey: config.height,
                AVVideoCompressionPropertiesKey: [
                    AVVideoAverageBitRateKey: 6_000_000,
                    AVVideoProfileLevelKey: AVVideoProfileLevelH264HighAutoLevel
                ]
            ]
            let videoInput = AVAssetWriterInput(mediaType: .video, outputSettings: videoSettings)
            videoInput.expectsMediaDataInRealTime = true
            let adaptor = AVAssetWriterInputPixelBufferAdaptor(assetWriterInput: videoInput, sourcePixelBufferAttributes: [
                kCVPixelBufferPixelFormatTypeKey as String: kCVPixelFormatType_32BGRA,
                kCVPixelBufferWidthKey as String: config.width,
                kCVPixelBufferHeightKey as String: config.height
            ])

            if writer.canAdd(videoInput) {
                writer.add(videoInput)
            }

            self.videoInput = videoInput
            self.adaptor = adaptor
        }

        if config.includeAudio {
            let audioSettings: [String: Any] = [
                AVFormatIDKey: kAudioFormatMPEG4AAC,
                AVNumberOfChannelsKey: 2,
                AVSampleRateKey: 48000,
                AVEncoderBitRateKey: 192_000
            ]
            let audioInput = AVAssetWriterInput(mediaType: .audio, outputSettings: audioSettings)
            audioInput.expectsMediaDataInRealTime = true
            if writer.canAdd(audioInput) {
                writer.add(audioInput)
            }
            self.audioInput = audioInput
        }

        self.assetWriter = writer
    }

    func stream(_ stream: SCStream, didStopWithError error: Error) {
        NSLog("Stream stopped with error: \(error)")
    }

    func stream(_ stream: SCStream, didOutputSampleBuffer sampleBuffer: CMSampleBuffer, of type: SCStreamOutputType) {
        guard CMSampleBufferDataIsReady(sampleBuffer) else { return }

        if startTime == nil {
            startTime = CMSampleBufferGetPresentationTimeStamp(sampleBuffer)
            assetWriter?.startWriting()
            if let startTime {
                assetWriter?.startSession(atSourceTime: startTime)
            }
            sessionStarted = true
        }

        guard sessionStarted else { return }

        switch type {
        case .screen:
            guard config.includeVideo, let videoInput, let adaptor else { return }
            guard videoInput.isReadyForMoreMediaData else { return }
            if let pixelBuffer = CMSampleBufferGetImageBuffer(sampleBuffer) {
                let pts = CMSampleBufferGetPresentationTimeStamp(sampleBuffer)
                adaptor.append(pixelBuffer, withPresentationTime: pts)
            }
        case .audio:
            guard config.includeAudio, let audioInput else { return }
            if audioInput.isReadyForMoreMediaData {
                audioInput.append(sampleBuffer)
            }
            writeWav(sampleBuffer: sampleBuffer)
        case .microphone:
            break
        @unknown default:
            break
        }
    }

    private func writeWav(sampleBuffer: CMSampleBuffer) {
        guard let wavFile = audioFile ?? createWavFile(from: sampleBuffer) else { return }
        audioFile = wavFile
        guard let pcmBuffer = pcmBufferFromSampleBuffer(sampleBuffer) else { return }
        do {
            try wavFile.write(from: pcmBuffer)
        } catch {
            NSLog("Failed to write wav: \(error)")
        }
    }

    private func createWavFile(from sampleBuffer: CMSampleBuffer) -> AVAudioFile? {
        guard let formatDesc = CMSampleBufferGetFormatDescription(sampleBuffer),
              let asbdPtr = CMAudioFormatDescriptionGetStreamBasicDescription(formatDesc) else { return nil }

        var asbd = asbdPtr.pointee
        guard let format = AVAudioFormat(streamDescription: &asbd) else { return nil }

        do {
            return try AVAudioFile(forWriting: config.outputWav, settings: format.settings, commonFormat: format.commonFormat, interleaved: format.isInterleaved)
        } catch {
            NSLog("Failed to create wav file: \(error)")
            return nil
        }
    }

    private func pcmBufferFromSampleBuffer(_ sampleBuffer: CMSampleBuffer) -> AVAudioPCMBuffer? {
        guard let formatDesc = CMSampleBufferGetFormatDescription(sampleBuffer),
              let asbdPtr = CMAudioFormatDescriptionGetStreamBasicDescription(formatDesc) else { return nil }

        var asbd = asbdPtr.pointee
        guard let format = AVAudioFormat(streamDescription: &asbd) else { return nil }

        let frameCount = Int32(CMSampleBufferGetNumSamples(sampleBuffer))
        guard let pcmBuffer = AVAudioPCMBuffer(pcmFormat: format, frameCapacity: AVAudioFrameCount(frameCount)) else { return nil }
        pcmBuffer.frameLength = AVAudioFrameCount(frameCount)

        let status = CMSampleBufferCopyPCMDataIntoAudioBufferList(
            sampleBuffer,
            at: 0,
            frameCount: frameCount,
            into: pcmBuffer.mutableAudioBufferList
        )

        guard status == noErr else { return nil }

        return pcmBuffer
    }
}

@main
struct RecorderHelper {
    static func main() async {
        do {
            let config = try parseConfig()
            let service = RecorderService(config: config)
            try await service.start()

            let stopSignal = DispatchSemaphore(value: 0)
            let stdinQueue = DispatchQueue(label: "recorder.stdin")
            stdinQueue.async {
                while let line = readLine() {
                    if line.lowercased().contains("stop") {
                        stopSignal.signal()
                        break
                    }
                }
            }

            signal(SIGINT, SIG_IGN)
            let sigSource = DispatchSource.makeSignalSource(signal: SIGINT, queue: DispatchQueue.global())
            sigSource.setEventHandler {
                stopSignal.signal()
            }
            sigSource.resume()

            await waitSemaphore(stopSignal)
            try await service.stop()
        } catch {
            fputs("RecorderHelper error: \(error)\n", stderr)
            exit(1)
        }
    }

    private static func parseConfig() throws -> RecorderConfig {
        var outputMp4: URL?
        var outputWav: URL?
        var fps = 30
        var width = 1920
        var height = 1080
        var left = 0
        var top = 0
        var includeVideo = true
        var includeAudio = true
        var audioMode = "native"
        var displayId: UInt32 = 0
        var captureCurrentSpaceOnly = false

        var iterator = CommandLine.arguments.dropFirst().makeIterator()
        while let arg = iterator.next() {
            switch arg {
            case "--output":
                if let value = iterator.next() { outputMp4 = URL(fileURLWithPath: value) }
            case "--wav":
                if let value = iterator.next() { outputWav = URL(fileURLWithPath: value) }
            case "--fps":
                if let value = iterator.next(), let parsed = Int(value) { fps = parsed }
            case "--width":
                if let value = iterator.next(), let parsed = Int(value) { width = parsed }
            case "--height":
                if let value = iterator.next(), let parsed = Int(value) { height = parsed }
            case "--left":
                if let value = iterator.next(), let parsed = Int(value) { left = parsed }
            case "--top":
                if let value = iterator.next(), let parsed = Int(value) { top = parsed }
            case "--display-id":
                if let value = iterator.next(), let parsed = UInt32(value) { displayId = parsed }
            case "--current-space-only":
                captureCurrentSpaceOnly = true
            case "--no-video":
                includeVideo = false
            case "--no-audio":
                includeAudio = false
            case "--audio-mode":
                if let value = iterator.next() { audioMode = value.lowercased() }
            default:
                continue
            }
        }

        guard let outputMp4, let outputWav else {
            throw NSError(domain: "RecorderHelper", code: 3, userInfo: [NSLocalizedDescriptionKey: "Missing --output or --wav"])
        }

        return RecorderConfig(
            outputMp4: outputMp4,
            outputWav: outputWav,
            fps: fps,
            width: width,
            height: height,
            left: left,
            top: top,
            displayId: displayId,
            captureCurrentSpaceOnly: captureCurrentSpaceOnly,
            includeVideo: includeVideo,
            includeAudio: includeAudio,
            audioMode: audioMode
        )
    }

    private static func waitSemaphore(_ semaphore: DispatchSemaphore) async {
        await withCheckedContinuation { continuation in
            DispatchQueue.global().async {
                semaphore.wait()
                continuation.resume()
            }
        }
    }
}
