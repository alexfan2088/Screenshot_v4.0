# Git和GitHub同步指南

## 第一步：安装Git
1. 访问 https://git-scm.com/download/win
2. 下载并安装Git for Windows
3. 安装时选择默认选项
4. 安装完成后重启命令行

## 第二步：配置Git用户信息
打开命令行，运行：
```bash
git config --global user.name "您的姓名"
git config --global user.email "您的邮箱@example.com"
```

## 第三步：在GitHub创建仓库
1. 访问 https://github.com
2. 登录账号
3. 点击右上角 "+" → "New repository"
4. 仓库名：Screenshot_v3.0
5. 描述：WPF截图和音频录制工具
6. 选择Public或Private
7. 不要勾选任何选项
8. 点击"Create repository"

## 第四步：在项目目录执行Git命令
在 C:\dev\Screenshot_v3.0 目录下执行：

```bash
# 初始化Git仓库
git init

# 添加所有文件到暂存区
git add .

# 提交文件
git commit -m "初始提交：WPF截图和音频录制工具"

# 添加远程仓库（替换为您的GitHub仓库地址）
git remote add origin https://github.com/您的用户名/Screenshot_v3.0.git

# 推送到GitHub
git push -u origin main
```

## 第五步：验证上传成功
1. 刷新GitHub页面
2. 应该能看到所有文件
3. 包括：MainWindow.xaml, AudioRecorder.cs 等

## 后续更新代码的步骤
```bash
# 添加修改的文件
git add .

# 提交修改
git commit -m "描述您的修改"

# 推送到GitHub
git push
```


