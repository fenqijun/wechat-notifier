import os
import time
import re
import threading
import win32gui
import win32process
import psutil
import ctypes
import logging
from datetime import datetime
import pythoncom  # 添加这行
import uiautomation as auto
from winotify import Notification, audio

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 微信窗口信息
WECHAT_WINDOW_NAME = "微信"
SCAN_INTERVAL = 0.2  # 扫描间隔（秒）

# 通知声音设置
NOTIFICATION_SOUND = audio.Mail  # 添加这行

# 正则表达式匹配消息内容
MESSAGE_REGEX = r"(.+?)(?:(\d+)条新消息)?(?:已置顶)?$"  # 匹配联系人名称和新消息数量
TIMESTAMP_REGEX = r"\d{1,2}/\d{1,2}/\d{1,2}|\d{1,2}:\d{2}"

# 添加全局变量
notified_messages = {}  # 改用字典存储消息历史，格式: {contact_name: (count, content)}

def find_wechat_window():
    """查找微信主窗口"""
    wechat_window = auto.WindowControl(searchDepth=1, Name=WECHAT_WINDOW_NAME)
    if not wechat_window.Exists():
        logging.error("微信窗口未找到")
        return None
    
    # 设置 UIAutomation 配置，允许处理隐藏控件
    auto.SetGlobalSearchTimeout(0.5)  # 设置搜索超时
    auto.uiautomation.SEARCH_INTERVAL = 0.2  # 设置搜索间隔
    auto.uiautomation.MAX_SEARCH_DEPTH = 10  # 设置最大搜索深度
    
    logging.info("找到微信窗口")
    return wechat_window

def monitor_wechat_messages():
    """监控微信消息"""
    pythoncom.CoInitialize()
    
    chat_list = find_wechat_window()
    if not chat_list:
        return

    logging.info("开始监控微信消息...")
    try:
        while True:
            try:
                # 递归扫描所有控件，包括隐藏的
                def scan_controls(control):
                    try:
                        if control.Name and "条新消息" in control.Name:
                            message = extract_message_content(control)
                            if message:
                                send_notification("微信新消息", message)
                                logging.info(f"发现新消息: {message}")
                        
                        # 获取所有子控件，包括隐藏的
                        children = control.GetChildren()
                        for child in children:
                            scan_controls(child)
                    except Exception as e:
                        logging.debug(f"扫描控件失败: {e}")
                
                scan_controls(chat_list)
                        
            except Exception as e:
                logging.error(f"监控失败: {e}")
            time.sleep(SCAN_INTERVAL)
    finally:
        pythoncom.CoUninitialize()

def send_notification(title, message):
    """发送Windows系统通知"""
    try:
        # 分割消息内容
        message_parts = message.split('\n')
        title_text = message_parts[0]
        content_text = message_parts[1] if len(message_parts) > 1 else ""
        
        toast = Notification(
            app_id="微信",
            title=title_text,
            msg=content_text,
            icon=r"C:\Users\Administrator\Desktop\微信.png",
            duration="short"
        )
        
        # 添加声音
        toast.set_audio(NOTIFICATION_SOUND, loop=False)
        toast.show()
        logging.info(f"通知已发送: {message}")
    except Exception as e:
        logging.error(f"发送通知失败: {e}")

def extract_message_content(control):
    """从控件中提取消息内容"""
    try:
        if not control.Name:
            return None
            
        # 检查是否包含"条新消息"
        if "条新消息" not in control.Name:
            return None
            
        # 提取联系人名称和消息数量（处理"已置顶"的情况）
        match = re.match(r"(.+?)(?:已置顶)?(\d+)条新消息$", control.Name)
        if not match:
            return None
            
        contact_name = match.group(1).strip()
        message_count = match.group(2)
        
        # 获取最新消息内容
        latest_message = ""
        max_top = 0
        
        def find_text_controls(ctrl):
            nonlocal latest_message, max_top
            # 递归查找所有子控件
            for child in ctrl.GetChildren():
                # 检查是否是消息内容控件
                if (child.ControlType == 50020 and child.Name and 
                    child.Name != contact_name and 
                    not re.match(r"\d{1,2}:\d{2}", child.Name) and  # 排除时间
                    not child.Name.isdigit()):  # 排除消息数量
                    latest_message = child.Name
                    break  # 找到消息内容后直接退出
                
                # 继续递归查找
                find_text_controls(child)
        
        # 开始递归查找文本控件
        find_text_controls(control)
        
        # 检查消息是否有变化
        if contact_name in notified_messages:
            old_count, old_content = notified_messages[contact_name]
            if old_count == message_count and old_content == latest_message:
                return None
        
        # 更新消息历史
        notified_messages[contact_name] = (message_count, latest_message)
            
        # 构建通知消息
        notification = contact_name
        if int(message_count) > 1:  # 只在消息数量大于1时显示数量
            notification += f" ({message_count})"
        if latest_message:
            notification += f"\n{latest_message}"
        
        return notification
        
    except Exception as e:
        logging.error(f"提取消息内容失败: {e}")
    return None

def main():
    """主函数"""
    # 检查微信是否运行
    wechat_process = None
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == "WeChat.exe":
            wechat_process = proc
            break
    
    if not wechat_process:
        # 发送通知提醒用户启动微信
        toast = Notification(
            app_id="微信监控",
            title="微信未运行",
            msg="请先启动微信客户端，程序将在微信启动后自动运行",
            icon=r"C:\Users\Administrator\Desktop\微信.png",
            duration="short"
        )
        toast.set_audio(NOTIFICATION_SOUND, loop=False)
        toast.show()
        logging.error("微信未运行")
        
        # 持续检测微信进程
        while not wechat_process:
            time.sleep(2)  # 每2秒检查一次
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] == "WeChat.exe":
                    wechat_process = proc
                    break
    
    # 等待并检测微信窗口
    wechat_window = find_wechat_window()
    if not wechat_window:
        toast = Notification(
            app_id="微信",
            title="等待微信窗口",
            msg="请确保微信窗口已打开，程序将在检测到窗口后自动运行",
            icon=r"C:\Users\Administrator\Desktop\微信.png",
            duration="short"
        )
        toast.set_audio(NOTIFICATION_SOUND, loop=False)
        toast.show()
        
        # 持续检测微信窗口
        while not wechat_window:
            time.sleep(2)  # 每2秒检查一次
            wechat_window = find_wechat_window()
    
    # 启动监控线程
    monitor_thread = threading.Thread(target=monitor_wechat_messages, daemon=True)
    monitor_thread.start()
    
    # 主线程保持运行
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("程序已停止")

if __name__ == "__main__":
    # 设置更详细的日志级别
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('wechat_monitor.log', encoding='utf-8')
        ]
    )
    main()
