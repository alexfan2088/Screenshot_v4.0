// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "RecorderHelper",
    platforms: [
        .macOS(.v13)
    ],
    products: [
        .executable(name: "RecorderHelper", targets: ["RecorderHelper"])
    ],
    targets: [
        .executableTarget(
            name: "RecorderHelper",
            path: "Sources/RecorderHelper"
        )
    ]
)
