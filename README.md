# Windows微信消息弹窗通知工具

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)

实时监控微信聊天消息并通过系统通知提醒的工具，支持智能过滤和防骚扰模式。

## 功能特性 ✨

- **实时监控**：自动检测微信窗口消息变化
- **智能过滤**：排除时间戳和系统干扰信息
- **频率控制**：5秒内超过3条消息自动聚合
- **多消息支持**：
  - 文本消息
  - 语音提示
  - 转账提醒
  - 红包通知
- **日志系统**：完整记录运行状态和错误信息

<div align="center">
  <img src="https://github.com/user-attachments/assets/b328aeb1-ceb4-4366-ade6-0acff4ddd4fe" width="600" alt="界面示意图">
</div>

## 运行要求 🖥️

- Windows 10/11 系统
- 微信桌面版 3.9+
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
