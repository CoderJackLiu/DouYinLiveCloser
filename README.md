# 直播伴侣自动关播工具 V1.0

## 项目简介

`直播伴侣自动关播工具` 是一款自动化工具，旨在帮助主播在指定时间自动关闭直播。通过模拟鼠标点击，自动完成关播和确认操作。工具支持用户自定义鼠标点击的偏移量，以确保操作精确无误。此外，还提供倒计时功能，可以在用户选择的时间自动执行关播任务。

## 功能特性

- **自动窗口激活**：工具能够自动激活 `直播伴侣` 窗口，并进行鼠标点击操作。
- **鼠标点击模拟**：模拟鼠标左键点击，执行关播操作并确认。
- **自定义偏移设置**：支持用户设置关播按钮和确认按钮的水平和垂直偏移，确保鼠标点击准确。
- **倒计时功能**：允许用户选择具体的时间来执行关播，或者快速倒计时 1 秒以立即执行任务。
- **配置保存**：偏移设置将保存到配置文件中，确保下次启动时能够自动载入。

## 环境要求

- Windows 系统
- 需要管理员权限运行
- Python 3.6 及以上版本
- 依赖库（已打包在工具内，无需用户单独安装）

## 安装步骤

1. 解压下载的压缩包文件到任意目录。
2. 找到 `直播伴侣自动关播工具.exe`，双击运行程序。
3. 程序启动时将自动检查是否具备管理员权限，如果没有，将提示以管理员身份重新启动。

## 使用说明

### 1. 设置倒计时

- 打开软件后，选择你希望执行关播操作的日期和时间。
- 点击“启动倒计时”按钮，程序将开始倒计时，并在时间到达后自动执行关播任务。
- 也可以点击“倒计时 1 秒”按钮，快速执行关播操作。

### 2. 调整鼠标偏移

- 在界面中的“关播按钮水平偏移”和“垂直偏移”中输入合适的偏移值，以确保鼠标点击准确。
- 偏移值会自动保存，下次启动时会自动加载之前的设置。

### 3. 自动关播流程

- 当倒计时结束后，程序会：
  1. 激活 `直播伴侣` 窗口。
  2. 将鼠标移动到窗口底部中间的“关播”按钮，应用偏移后点击。
  3. 移动鼠标到确认窗口的中间，点击确认按钮，完成关播。

## 文件说明

- `直播伴侣自动关播工具.exe`：主程序文件，用户只需运行此文件。
- `offset_config.json`：保存鼠标偏移值的配置文件，用户无需手动修改，程序会自动读取和保存。

## 注意事项

- **管理员权限**：该工具需要管理员权限运行。如果你以非管理员身份启动，程序将提示重新启动。
- **偏移配置**：偏移值会自动保存在 `offset_config.json` 文件中。请确保该文件有读写权限，以便保存设置。
- **窗口标题**：该工具默认寻找标题为 `直播伴侣` 的窗口。如果你的直播工具窗口标题不同，可能需要手动修改配置。

## 常见问题

1. **鼠标点击不准确怎么办？**
   - 你可以通过调整偏移值来微调鼠标点击的位置。
2. **工具无法激活窗口？**
   - 请确保 `直播伴侣` 窗口是打开的，并且窗口标题是 `直播伴侣`。
3. **管理员权限问题？**
   - 运行程序时确保以管理员身份运行。如果没有管理员权限，程序将无法模拟鼠标操作。

## 更新日志

### V1.0

- 初始版本发布，支持倒计时自动关播、鼠标点击偏移设置及保存功能。
