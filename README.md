# Windows微信消息弹窗通知工具

![image](https://github.com/user-attachments/assets/b328aeb1-ceb4-4366-ade6-0acff4ddd4fe)


[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Windows微信消息弹窗通知，通过系统通知提醒新消息，支持消息过滤和防骚扰模式。

## 功能特性 ✨

- 实时监控微信聊天列表的新消息（被遮挡的会话无法进行通知）
- 智能过滤系统消息和时间戳
- 防骚扰模式（5秒内超过3条消息自动聚合）
- 支持多种消息类型：
  - 文本消息
  - 语音消息提示
  - 转账提醒
  - 红包通知
- 多线程监控架构
- 完整的日志记录系统

## 运行要求 🖥️

- Windows 10/11
- WeChat 3.9+ 桌面版
- Python 3.7+

## 快速开始 🚀

### 安装依赖
```bash
pip install uiautomation==2.0.15 winotify==0.3.0 psutil==5.9.0 pywin32==306
```
### 配置说明
1. 修改微信图标路径（可选）

# 微信通知.py 第35行
icon=r"C:\path\to\wechat.png"

## 常见问题 ❓
### 通知不显示？
✅ 检查：

1. 系统通知权限是否开启
2. 微信窗口是否保持前台打开
3. 图标路径是否正确

# 启用调试模式
set LOG_LEVEL=DEBUG && python 微信通知.py

## 贡献指南 🤝
欢迎提交PR！建议改进方向：

- 支持自定义通知模板
- 添加多语言支持
- 实现消息历史存储
