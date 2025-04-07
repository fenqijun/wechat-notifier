# Windows微信消息弹窗通知工具

<div align="center">
  <img src="https://github.com/user-attachments/assets/b328aeb1-ceb4-4366-ade6-0acff4ddd4fe" width="600" alt="界面示意图">
  <br>
  
  [![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
  
  > 🔔 **重要提醒**  
  > 建议置顶聊天不超过5个，否则可能遮挡会话导致无法通知
</div>

---

## 📥 版本下载
```diff
+ 最新稳定版：微信消息提醒-1.0-win64.zip 
! 已弃用版本：~~微信消息通知.exe~~（不再维护）
```

## 功能 ✨

- **实时监控**：自动检测微信窗口消息变化
- **智能过滤**：排除时间戳和系统干扰信息
- **频率控制**：5秒内超过3条消息自动聚合
- **多消息支持**：
  - 文本消息
  - 语音提示
  - 转账提醒
  - 红包通知
- **日志系统**：完整记录运行状态和错误信息
- **开机自启**： 默认开机自动启动

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
```
icon=r"C:\path\to\wechat.png"
```


### 核心依赖
| 包名称         | 版本     | 用途描述                  |
|----------------|----------|-------------------------|
| uiautomation   | 2.0.15   | Windows UI自动化控制       |
| winotify       | 0.3.0    | 系统通知中心集成           |
| psutil         | 5.9.0    | 进程监控与管理             |
| pywin32        | 306      | Windows API集成          |

### 开发依赖
| 包名称         | 版本     | 用途描述                  |
|----------------|----------|-------------------------|
| pythoncom      | 内置      | COM组件支持              |
| ctypes         | 内置      | C语言接口调用            |

## 常见问题 ❓
### 通知不显示？
✅ 检查：

1. 系统通知权限是否开启
2. 微信窗口是否保持前台打开
3. 图标路径是否正确
