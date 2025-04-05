import os
import time
import sys
import win32gui
import win32con
import win32api
import win32process
import psutil
import re
import threading
import ctypes
from ctypes import wintypes
from win32gui import FindWindow, FindWindowEx, GetWindowText, GetClassName
from win32process import GetWindowThreadProcessId
from winotify import Notification, audio
import logging
from datetime import datetime
import uiautomation as auto

# 添加诊断信息
try:
    import win32gui
    print(f"win32gui 已成功导入，版本信息: {win32gui}")
    logging.info(f"win32gui 已成功导入")
except ImportError as e:
    print(f"导入 win32gui 失败: {e}")
    logging.error(f"导入 win32gui 失败: {e}")
    # 显示 Python 路径信息
    print(f"Python 路径: {sys.executable}")
    print(f"sys.path: {sys.path}")

# 配置日志
logging.basicConfig(
    level=logging.DEBUG,  # 临时改为DEBUG级别以便调试
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='wechat_notifier.log',
    filemode='a'
)

# Windows API常量
WM_COPYDATA = 0x004A
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNTHREAD = 0x0001
WINEVENT_SKIPOWNPROCESS = 0x0002
EVENT_OBJECT_NAMECHANGE = 0x800C
EVENT_OBJECT_CREATE = 0x8000
EVENT_SYSTEM_FOREGROUND = 0x0003

# 定义回调函数类型
WinEventProcType = ctypes.WINFUNCTYPE(
    None, 
    wintypes.HANDLE, 
    wintypes.DWORD, 
    wintypes.HWND, 
    wintypes.LONG, 
    wintypes.LONG, 
    wintypes.DWORD, 
    wintypes.DWORD
)

class WeChatMonitor:
    def __init__(self):
        self.last_messages = {}  # 存储上次检测到的消息 {会话名: (消息数, 时间戳)}
        self.notification_queue = []  # 通知队列，用于消息叠加
        self.wechat_hwnd = None
        self.wechat_process_id = None
        self.hook_thread = None
        self.running = False
        self.hook_installed = False
        self.hook_handle = None
        self.event_proc = None
        self.user32 = ctypes.windll.user32
        self.ole32 = ctypes.windll.ole32
        
        # 初始化COM
        self.ole32.CoInitialize(0)
    
    def find_wechat_window(self):
        """查找微信主窗口"""
        self.wechat_hwnd = FindWindow("WeChatMainWndForPC", None)
        if not self.wechat_hwnd:
            # 尝试查找托盘中的微信
            def enum_windows_callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    logging.debug(f"窗口: {window_text}, 类名: {class_name}")
                    # 更宽松的匹配条件
                    if "微信" in window_text or "WeChat" in window_text:
                        results.append(hwnd)
                return True
            
            results = []
            win32gui.EnumWindows(enum_windows_callback, results)
            
            if results:
                self.wechat_hwnd = results[0]
                logging.info(f"找到可能的微信窗口: {win32gui.GetWindowText(self.wechat_hwnd)}")
            else:
                # 尝试通过进程名查找
                for proc in psutil.process_iter(['pid', 'name']):
                    if proc.info['name'] and 'WeChat' in proc.info['name']:
                        logging.info(f"找到微信进程: {proc.info['name']}, PID: {proc.info['pid']}")
                        self.wechat_process_id = proc.info['pid']
                        # 尝试查找该进程的主窗口
                        def enum_process_windows(hwnd, process_id):
                            tid, pid = GetWindowThreadProcessId(hwnd)
                            if pid == process_id and win32gui.IsWindowVisible(hwnd):
                                results.append(hwnd)
                            return True
                        
                        results = []
                        win32gui.EnumWindows(lambda hwnd, param: enum_process_windows(hwnd, self.wechat_process_id), None)
                        
                        if results:
                            self.wechat_hwnd = results[0]
                            logging.info(f"通过进程ID找到微信窗口: {win32gui.GetWindowText(self.wechat_hwnd)}")
                            break
                
                if not self.wechat_hwnd:
                    logging.warning("未找到微信窗口，请确保微信已启动")
                    return False
        
        # 获取微信进程ID
        if not self.wechat_process_id:
            _, self.wechat_process_id = GetWindowThreadProcessId(self.wechat_hwnd)
        
        logging.info(f"找到微信窗口，句柄: {self.wechat_hwnd}, 进程ID: {self.wechat_process_id}")
        return True
    
    def win_event_callback(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        """Windows事件回调函数"""
        try:
            # 检查是否是微信窗口或其子窗口
            if hwnd:
                _, process_id = GetWindowThreadProcessId(hwnd)
                if process_id == self.wechat_process_id:
                    # 获取窗口文本和类名
                    window_text = GetWindowText(hwnd)
                    class_name = GetClassName(hwnd)
                    
                    # 检查是否是聊天列表项
                    if class_name in ["ListItem", "Static"] and window_text:
                        # 检查是否包含未读消息计数 [数字]
                        match = re.search(r'(.*?)\s*\[(\d+)\]', window_text)
                        if match:
                            chat_name = match.group(1).strip()
                            message_count = int(match.group(2))
                            self.process_new_message(chat_name, message_count)
                        
                        # 记录窗口变化
                        logging.debug(f"窗口变化: {window_text} (类: {class_name})")
        except Exception as e:
            logging.error(f"事件回调中发生错误: {e}")
    
    def install_hook(self):
        """安装Windows事件钩子"""
        if self.hook_installed:
            return True
        
        # 创建回调函数
        self.event_proc = WinEventProcType(self.win_event_callback)
        
        # 安装多个钩子以捕获不同类型的事件
        hooks = []
        
        # 监听名称变化事件
        hook1 = self.user32.SetWinEventHook(
            EVENT_OBJECT_NAMECHANGE,
            EVENT_OBJECT_NAMECHANGE,
            0,
            self.event_proc,
            self.wechat_process_id,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNTHREAD
        )
        if hook1:
            hooks.append(hook1)
            
        # 监听对象创建事件
        hook2 = self.user32.SetWinEventHook(
            EVENT_OBJECT_CREATE,
            EVENT_OBJECT_CREATE,
            0,
            self.event_proc,
            self.wechat_process_id,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNTHREAD
        )
        if hook2:
            hooks.append(hook2)
        
        # 监听前台窗口变化事件
        hook3 = self.user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            0,
            self.event_proc,
            0,  # 监听所有进程
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNTHREAD
        )
        if hook3:
            hooks.append(hook3)
        
        if hooks:
            self.hook_handle = hooks  # 保存所有钩子句柄
            self.hook_installed = True
            logging.info(f"成功安装Windows事件钩子，共{len(hooks)}个")
            return True
        else:
            logging.error("安装Windows事件钩子失败")
            return False
    
    def uninstall_hook(self):
        """卸载Windows事件钩子"""
        if self.hook_installed and self.hook_handle:
            success = True
            for hook in self.hook_handle:
                result = self.user32.UnhookWinEvent(hook)
                if not result:
                    success = False
                    logging.error(f"卸载钩子 {hook} 失败")
            
            if success:
                self.hook_installed = False
                self.hook_handle = None
                logging.info("成功卸载所有Windows事件钩子")
                return True
            else:
                logging.error("部分Windows事件钩子卸载失败")
                return False
        return True
    
    def process_new_message(self, chat_name, message_count):
        """处理新消息"""
        current_time = time.time()
        
        # 检查是否是新消息
        if chat_name not in self.last_messages or message_count > self.last_messages[chat_name][0]:
            # 计算新增消息数
            prev_count = 0 if chat_name not in self.last_messages else self.last_messages[chat_name][0]
            new_messages = message_count - prev_count
            
            if new_messages > 0:
                # 发送通知
                # 创建并发送Windows通知
                notification = Notification(
                    app_id="微信",
                    title=f"来自 {chat_name} 的新消息",
                    msg=f"收到 {new_messages} 条新消息",
                    duration="short"
                )
                notification.set_audio(audio.Default, loop=False)
                notification.show()
                logging.info(f"检测到新消息: {chat_name} - {new_messages}条")
            
            # 更新记录
            self.last_messages[chat_name] = (message_count, current_time)
        else:
            # 更新时间戳
            self.last_messages[chat_name] = (message_count, current_time)
    
    def scan_wechat_ui(self):
        """使用UI自动化扫描微信界面查找未读消息"""
        try:
            # 添加调试信息
            logging.debug("开始扫描微信UI...")
            
            # 尝试查找微信主窗口
            wechat_window = None
            
            # 方法1: 通过类名查找
            try:
                wechat_window = auto.WindowControl(searchDepth=1, ClassName="WeChatMainWndForPC")
                if wechat_window.Exists(1):
                    logging.debug("通过类名找到微信窗口")
            except Exception as e:
                logging.debug(f"通过类名查找微信窗口失败: {e}")
            
            # 如果找到了微信窗口，尝试查找聊天列表
            if wechat_window and wechat_window.Exists(1):
                logging.debug(f"微信窗口信息: 名称={wechat_window.Name}, 类名={wechat_window.ClassName}")
                
                # 根据提供的元素结构，直接查找会话列表
                try:
                    # 查找名为"会话"的List控件
                    chat_list = wechat_window.ListControl(Name="会话")
                    if chat_list.Exists(1):
                        logging.debug("找到会话列表控件")
                        
                        # 获取所有可见列表项
                        list_items = chat_list.GetChildren()
                        logging.debug(f"会话列表项数量: {len(list_items)}")
                        
                        # 先处理可见项
                        self._process_list_items(list_items)
                        
                        # 尝试滚动列表查找更多项
                        self._scroll_and_scan(chat_list)
                        
                    else:
                        # 如果没有找到会话列表，尝试查找所有ListItem控件
                        logging.debug("未找到会话列表，尝试查找所有ListItem控件")
                        list_items = wechat_window.GetChildren()
                        self._process_list_items(list_items)
                except Exception as e:
                    logging.debug(f"查找会话列表时出错: {e}")
                    
                    # 备用方法：直接查找包含"条新消息"的ListItem
                    try:
                        # 使用模糊匹配查找所有可能包含新消息的列表项
                        for item in wechat_window.GetChildren():
                            if item.ControlType == auto.ControlType.ListItemControl:
                                try:
                                    # 处理每个列表项
                                    item_name = item.Name
                                    if "条新消息" in item_name:
                                        match = re.search(r'(.+?)(\d+)条新消息', item_name)
                                        if match:
                                            chat_name = match.group(1).strip()
                                            message_count = int(match.group(2))
                                            logging.info(f"备用方法找到未读消息: {chat_name} - {message_count}条")
                                            self.process_new_message(chat_name, message_count)
                                except:
                                    pass
                    except Exception as e:
                        logging.debug(f"备用方法查找失败: {e}")
        
        except Exception as e:
            logging.error(f"扫描微信UI时出错: {e}")
            import traceback
            logging.error(traceback.format_exc())
    
    def _process_list_items(self, list_items):
        """处理列表项"""
        for item in list_items:
            try:
                item_name = item.Name
                logging.debug(f"列表项名称: {item_name}")
                
                # 检查是否包含"条新消息"
                if "条新消息" in item_name:
                    # 提取联系人名称和消息数量
                    # 格式: "联系人名称X条新消息"
                    match = re.search(r'(.+?)(\d+)条新消息', item_name)
                    if match:
                        chat_name = match.group(1).strip()
                        message_count = int(match.group(2))
                        logging.info(f"找到未读消息: {chat_name} - {message_count}条")
                        self.process_new_message(chat_name, message_count)
                    else:
                        # 尝试查找子控件获取更多信息
                        panes = item.GetChildren()
                        contact_name = None
                        message_count = None
                        
                        for pane in panes:
                            try:
                                # 尝试查找联系人名称
                                static_text = pane.TextControl(searchDepth=2)
                                if static_text.Exists(1):
                                    if not contact_name:
                                        contact_name = static_text.Name
                                    elif static_text.Name.isdigit():
                                        message_count = int(static_text.Name)
                            except:
                                pass
                        
                        if contact_name and message_count:
                            logging.info(f"通过子控件找到未读消息: {contact_name} - {message_count}条")
                            self.process_new_message(contact_name, message_count)
            except Exception as e:
                logging.debug(f"处理列表项时出错: {e}")
    
    def _scroll_and_scan(self, chat_list):
        """滚动列表并扫描更多项"""
        try:
            # 获取列表的位置和大小
            rect = chat_list.BoundingRectangle
            
            # 尝试滚动列表
            scroll_attempts = 3  # 滚动尝试次数
            for i in range(scroll_attempts):
                # 模拟滚动操作
                auto.WheelDown(wheelTimes=3, x=rect.xcenter(), y=rect.ycenter())
                time.sleep(0.2)  # 等待滚动完成
                
                # 获取滚动后的列表项
                new_items = chat_list.GetChildren()
                logging.debug(f"滚动后列表项数量: {len(new_items)}")
                
                # 处理新的列表项
                self._process_list_items(new_items)
                
            # 滚动回顶部
            for i in range(scroll_attempts):
                auto.WheelUp(wheelTimes=3, x=rect.xcenter(), y=rect.ycenter())
                time.sleep(0.1)
                
        except Exception as e:
            logging.debug(f"滚动列表时出错: {e}")

    def message_loop(self):
        """消息循环，处理Windows消息"""
        msg = wintypes.MSG()
        last_scan_time = 0
        scan_interval = 0.5  # 减少扫描间隔到0.5秒，提高消息检测的及时性
        
        logging.info("开始消息循环...")
        print("开始监听微信消息...")
        
        while self.running:
            # 处理Windows消息
            if self.user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0, 1):
                self.user32.TranslateMessage(ctypes.byref(msg))
                self.user32.DispatchMessageW(ctypes.byref(msg))
            
            # 定期扫描微信UI
            current_time = time.time()
            if current_time - last_scan_time >= scan_interval:
                logging.debug(f"执行定期扫描，间隔: {current_time - last_scan_time:.1f}秒")
                self.scan_wechat_ui()
                last_scan_time = current_time
                
                # 清理过期的会话记录
                self.cleanup_old_records()
            
            # 极短休眠，减少CPU占用但保持高响应速度
            time.sleep(0.01)  # 将休眠时间从0.05秒减少到0.01秒

    def cleanup_old_records(self):
        """清理超过5分钟未更新的会话记录"""
        current_time = time.time()
        outdated = []
        for chat_name, (count, timestamp) in self.last_messages.items():
            if current_time - timestamp > 300:  # 5分钟 = 300秒
                outdated.append(chat_name)
        
        for chat_name in outdated:
            del self.last_messages[chat_name]
    
    def run(self):
        """运行监控"""
        print("微信消息监控已启动...")
        logging.info("微信消息监控已启动")
        
        try:
            # 查找微信窗口
            retry_count = 0
            while not self.find_wechat_window() and retry_count < 3:
                print(f"未找到微信窗口，{5}秒后重试...")
                logging.warning(f"未找到微信窗口，{5}秒后重试...")
                time.sleep(5)
                retry_count += 1
            
            if not self.wechat_hwnd:
                print("多次尝试后仍未找到微信窗口，请确保微信已启动")
                logging.error("多次尝试后仍未找到微信窗口，请确保微信已启动")
                return
            
            # 安装钩子
            if not self.install_hook():
                print("安装事件钩子失败，将使用轮询方式监控")
                logging.warning("安装事件钩子失败，将使用轮询方式监控")
            
            # 设置运行标志
            self.running = True
            
            # 启动消息循环
            self.message_loop()
            
        except KeyboardInterrupt:
            print("监控已停止")
            logging.info("监控已停止")
        except Exception as e:
            print(f"发生错误: {e}")
            logging.error(f"发生错误: {e}")
        finally:
            # 清理资源
            self.running = False
            self.uninstall_hook()
            self.ole32.CoUninitialize()

if __name__ == "__main__":
    try:
        print("微信消息监控程序启动中...")
        print("按Ctrl+C可停止程序")
        
        # 检查是否已有实例在运行
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            # 尝试绑定一个本地端口，如果绑定失败说明已有实例在运行
            sock.bind(('127.0.0.1', 12345))
            sock.listen(5)
            
            # 启动监控
            monitor = WeChatMonitor()
            monitor.run()
        except socket.error:
            print("检测到程序已在运行！")
            logging.warning("检测到程序已在运行")
        finally:
            sock.close()
    except Exception as e:
        print(f"程序启动失败: {e}")
        logging.error(f"程序启动失败: {e}")
        import traceback
        logging.error(traceback.format_exc())
        
        # 等待用户按键退出
        input("按Enter键退出...")
