# 微信消息通知工具

## 项目简介

这是一个基于Python开发的微信消息通知工具，可以在后台监控微信消息，当有新消息到达时，通过Windows系统通知提醒用户，即使微信窗口被最小化或被其他窗口覆盖。

## 功能特点

- 自动检测微信窗口并监控新消息
- 通过Windows系统通知显示新消息提醒
- 显示消息发送者和新消息数量
- 后台运行，低资源占用
- 自动清理过期的会话记录

## 技术实现

该工具主要通过以下技术实现消息监控：

1. **Windows API钩子**：使用Windows事件钩子监听微信窗口的变化
2. **UI自动化**：使用uiautomation库扫描微信界面查找未读消息
3. **Windows通知**：使用winotify库发送系统通知

## 安装依赖

在使用前，请确保安装以下依赖库：

```bash
pip install pywin32 psutil winotify uiautomation
```

## 使用方法

1. 确保微信已经登录并运行
2. 运行微信通知工具：

```bash
python d:\Code\微信notice\微信通知.py
```

3. 工具将在后台运行，当有新消息时会自动弹出系统通知
4. 按Ctrl+C可以停止程序

## 工作原理

1. 程序启动后，首先查找微信主窗口
2. 安装Windows事件钩子，监听窗口变化事件
3. 定期扫描微信界面，查找包含未读消息的会话
4. 当检测到新消息时，发送Windows系统通知
5. 自动清理5分钟内未更新的会话记录，减少内存占用

## 日志记录

程序运行过程中的日志会记录在`wechat_notifier.log`文件中，可用于排查问题。日志级别默认设置为DEBUG，记录详细的运行信息。

## 注意事项

- 请确保微信已经登录并正常运行
- 程序使用端口12345检测是否已有实例在运行，避免重复启动
- 如遇到问题，可查看日志文件进行排查
- 程序需要在有图形界面的环境中运行

## 系统要求

- Windows操作系统
- Python 3.6+
- 微信PC版
