import socket
import cv2
import numpy as np
import os
import re
import threading
import signal
import json
import pyautogui
import time
import keyboard
import ctypes
import sys
import mss
import traceback
import select
from tkinter import Tk, Label


# ================= 管理员权限获取部分 =================
def is_admin():
    """检查当前进程是否以管理员权限运行"""
    try:
        # ctypes.windll.shell32 访问Windows Shell32库
        # IsUserAnAdmin() 是该库中的函数，返回非零值表示管理员
        return ctypes.windll.shell32.IsUserAnAdmin() # 非零为True
    except:
        return False # 异常时默认非管理员


def request_admin_privileges():
    """请求提升管理员权限(重新启动程序并获取管理员身份)"""
    if not is_admin(): # 如果当前非管理员
        print("正在请求管理员权限...")
        # ShellExecuteW 是Windows API，用于启动新进程
        # 参数说明：
        # hwnd: 父窗口句柄 None表示无
        # lpVerb: 操作动词,"runas"表示以管理员身份运行
        # lpFile: 要运行的程序路径(sys.executable是Python解释器路径)
        # lpParameters: 命令行参数(__file__是当前脚本路径)
        # lpDirectory: 工作目录(None表示使用默认)
        # nShowCmd: 窗口显示方式(1表示正常显示)
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, __file__, None, 1
        )
        sys.exit() # 退出当前非管理员进程,等待新进程启动


request_admin_privileges() # 程序入口处强制检查权限
# ================= 管理员权限获取部分 =================


# ================= ipv6地址获取部分 =================
def get_ipv6_address():
    """获取本机有效IPv6地址(优先全局单播地址,非链路本地地址)"""
    try:
        # os.popen 执行系统命令,返回命令输出流
        # "ipconfig /all" 获取所有网络接口配置信息
        output = os.popen("ipconfig /all").read()  # 读取命令输出文本
        # 正则表达式1: 匹配标准8段IPv6地址,非fe80::开头
        # (?!fe80::) 负向先行断言: 排除以fe80::开头的链路本地地址
        # ([0-9a-f]{1,4}:){7}[0-9a-f]{1,4} 匹配8个16进制段,每段1-4字符,用:分隔
        # re.I 标志使匹配不区分大小写
        result = re.findall(r"((?!fe80::)([0-9a-f]{1,4}:){7}[0-9a-f]{1,4})", output, re.I)

        if not result:# 未找到标准格式,尝试匹配压缩格式(允许省略部分段)
            # ([0-9a-f]{1,4}(:[0-9a-f]{1,4}){1,6}) 匹配2-7个段(压缩格式)
            result = re.findall(r"((?!fe80::)[0-9a-f]{1,4}(:[0-9a-f]{1,4}){1,6})", output, re.I)
            if not result: # 仍未找到,抛出异常
                raise ValueError("未找到有效的IPv6地址")

        # result[0][0]：正则匹配结果中,第一个全匹配项(组0)的第一个捕获组
        return result[0][0] # 返回第一个有效IPv6地址
    except Exception as e:
        print(f"获取IPv6地址失败: {e}")
        try: # 回退到IPv4地址
            return socket.gethostbyname(socket.gethostname())  # 获取主机IPv4地址
        except:
            return "0.0.0.0" # 兜底返回基础地址
# ================= ipv6地址获取部分 =================


# ================= 动态画质设置部分 =================
QUALITY_CONFIG = {
    # 键为帧率区间元组(左闭右开区间),值为(宽度,高度,JPEG质量,描述)
    (0, 5): (1280, 720, 10, "垃圾帧率: 1280×720 质量10"), # 当帧率<5时使用720P低画质
    (5, 10): (1280, 720, 20, "垃圾帧率: 1280×720 质量20"),
    (10, 15): (1280, 720, 30, "垃圾帧率: 1280×720 质量30"),
    (15, 20): (1280, 720, 40, "一般帧率: 1280×720 质量40"),
    (20, 25): (1280, 720, 50, "一般帧率: 1280×720 质量50"),
    (25, 30): (1280, 720, 60, "一般帧率: 1280×720 质量60"),
    (30, 35): (1920, 1080, 40, "良好帧率: 1920×1080 质量40"),
    (35, 40): (1920, 1080, 50, "良好帧率: 1920×1080 质量50"),
    (40, 45): (1920, 1080, 60, "良好帧率: 1920×1080 质量60"),
    (45, 50): (1920, 1080, 70, "优秀帧率: 1920×1080 质量70"),
    (50, 55): (1920, 1080, 80, "优秀帧率: 1920×1080 质量80"),
    (55, 60): (1920, 1080, 90, "优秀帧率: 1920×1080 质量90"),
    (60, float('inf')): (1920, 1080, 100, "最高帧率: 1920×1080 质量100") # 帧率≥60时使用1080P最高画质
}


class VideoQualityManager:
    def __init__(self):
        self.current_config = QUALITY_CONFIG[(25, 30)] # 默认使用25-30fps档位平衡性能
        self.last_adjust_time = time.time() # 记录上次调整时间,用于冷却机制
        self.adjust_interval = 1 # 调整间隔秒,避免每秒调整多次

    def get_config(self, current_fps):
        """根据当前帧率查找对应的画质配置"""
        for fps_range, config in QUALITY_CONFIG.items():
            if fps_range[0] <= current_fps < fps_range[1]: # 左闭右开区间判断
                return config # 返回匹配的配置
        return self.current_config # 未匹配时返回当前配置(防止配置丢失)

    def adjust_quality(self, current_fps):
        """带冷却机制的画质调整函数"""
        current_time = time.time()
        if current_time - self.last_adjust_time > self.adjust_interval: # 冷却时间已过
            new_config = self.get_config(current_fps) # 获取新配置
            if new_config != self.current_config: # 配置有变化时更新
                self.current_config = new_config # 更新当前配置
                # 打印提示信息，包含新配置描述和当前帧率
                print(f"画质调整: {new_config[3]} (当前帧率: {current_fps:.1f} FPS)")
            self.last_adjust_time = current_time # 记录调整时间
        return self.current_config # 返回当前配置
# ================= 动态画质设置部分 =================


# ================= 屏幕捕捉部分 =================
def capture_screen():
    """捕获主显示器屏幕(使用mss库)"""
    with mss.mss() as sct: # 使用mss上下文管理器
        monitor = sct.monitors[1] # sct.monitors 包含所有显示器信息,monitors[1]是主显示器,monitors[0]是全屏幕
        sct_img = sct.grab(monitor) # sct.grab(monitor) 捕获指定显示器,返回mss.image.Image对象
        img = np.array(sct_img) # np.array 将截图转为numpy数组,形状为(height, width, channels)
        return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB) # cv2.COLOR_RGBA2RGB 将RGBA颜色空间转换为RGB,mss默认RGBA,OpenCV支持RGB
# ================= 屏幕捕捉部分 =================


# ================= 视频流处理线程 =================
def handle_video_client(client_socket, client_address):
    """处理视频流客户端的独立线程函数"""
    try:
        print(f"开始处理客户端 {client_address} 的视频请求")
        MAX_FPS = 60 # 目标最大帧率,限制发送速度
        frame_interval = 1.0 / MAX_FPS # 每帧间隔时间
        last_frame_time = time.time() # 记录上次发送帧的时间戳

        last_second = int(time.time()) # 性能统计的时间戳
        capture_count = 0 # 每秒捕获的帧数统计
        process_count = 0 # 每秒处理的帧数统计
        process_time_sum = 0.0 # 处理耗时总和
        send_count = 0 # 发送帧数统计

        with mss.mss() as sct: # 每个线程独立创建mss实例
            monitor = sct.monitors[1]
            original_width = monitor["width"] # 原始分辨率宽度
            original_height = monitor["height"] # 原始分辨率高度
            quality_manager = VideoQualityManager() # 初始化画质管理器

            while True:
                now = time.time()
                elapsed = now - last_frame_time # 距离上次发送的时间差
                if elapsed < frame_interval: # 控制帧率:若时间差不足一帧间隔,等待剩余时间
                    time.sleep(frame_interval - elapsed)
                    continue  # 跳过本次循环,不处理帧
                last_frame_time = now # 更新上次发送时间

                current_second = int(now) # 每秒统计一次性能数据
                if current_second > last_second:
                    print(f"\n[统计 {last_second}s-{current_second - 1}s] "
                          f"截取帧数: {capture_count} "
                          f"处理帧数: {process_count} "
                          f"处理耗时: {process_time_sum:.1f}ms "
                          f"发送帧数: {send_count} ")
                    # 重置统计变量
                    capture_count = 0
                    process_count = 0
                    process_time_sum = 0.0
                    send_count = 0
                    last_second = current_second

                # 1. 屏幕捕获阶段
                capture_start = time.time()
                frame = capture_screen() # 调用截图函数
                capture_time = (time.time() - capture_start) * 1000 # 转换为毫秒
                capture_count += 1 # 统计捕获次数

                # 2. 画质调整与分辨率缩放
                # 使用send_count近似当前帧率(每秒发送帧数)
                width, height, quality, _ = quality_manager.adjust_quality(send_count)
                resize_start = time.time()
                # 若需要调整分辨率(与原始分辨率不同时)
                if (width, height) != (original_width, original_height):
                    # cv2.resize 调整图像大小,插值方法默认(双线性插值)
                    frame = cv2.resize(frame, (width, height))
                resize_time = (time.time() - resize_start) * 1000 # 计算缩放耗时

                # 3. JPEG编码阶段
                encode_start = time.time()
                # cv2.imencode 编码图像为JPEG格式
                # 参数:'.jpg' 格式，[cv2.IMWRITE_JPEG_QUALITY, quality] 编码质量
                _, img_encoded = cv2.imencode('.jpg', frame, [int(cv2.IMWRITE_JPEG_QUALITY), quality])
                encode_time = (time.time() - encode_start) * 1000 # 计算编码耗时
                process_count += 1 # 统计处理次数
                process_time_sum += resize_time + encode_time # 累计处理耗时

                # 4. 网络发送阶段
                data = img_encoded.tobytes() # 将编码后的图像转为字节流
                size = len(data) # 获取数据大小
                try:# 先发送4字节大端序big-endian的尺寸信息
                    client_socket.sendall(size.to_bytes(4, byteorder='big'))
                    client_socket.sendall(data) # 发送图像数据
                    send_count += 1 # 统计发送次数
                except (ConnectionResetError, ConnectionAbortedError):
                    # 客户端主动断开连接时捕获异常
                    print(f"客户端 {client_address} 主动断开连接")
                    break # 跳出循环，关闭连接

                # 实时状态输出
                print(f"[实时] 捕获: {capture_time:.1f}ms "
                      f"处理: {resize_time + encode_time:.1f}ms "
                      f"队列延迟: {elapsed * 1000:.1f}ms")

    except Exception as e: # 捕获线程内所有异常
        print(f"处理客户端 {client_address} 时出错: {e}")
        traceback.print_exc() # 打印详细异常栈,包含代码行号
    finally:
        client_socket.close() # 确保关闭客户端连接,释放资源
        print(f"客户端 {client_address} 视频连接已关闭")
# ================= 视频流处理线程 =================


# ================= 鼠标控制处理线程 =================
def handle_mouse_client(client_socket, client_address):
    """处理鼠标控制客户端的独立线程函数"""
    try:
        print(f"开始处理客户端 {client_address} 的鼠标控制请求")
        screen_width, screen_height = pyautogui.size() # pyautogui.size() 获取当前屏幕分辨率,宽度,高度
        pyautogui.PAUSE = 0.0 # 关闭pyautogui的操作延迟
        pyautogui.FAILSAFE = True # 启用安全机制:鼠标移到左上角时停止操作
        current_x, current_y = 0, 0 # 当前鼠标绝对坐标,初始为0,0
        is_mouse_down = False  # 鼠标左键按下状态,默认未按下

        while True:
            data = client_socket.recv(1024) # 接收最多1024字节数据
            if not data:  # 客户端断开连接时,recv返回空字节
                break
            # 将接收到的字节流解码为UTF-8字符串,并按换行符分割多条指令
            messages = data.decode('utf-8').split('\n')

            for message in messages:
                if message.strip(): # 跳过空消息
                    try:
                        mouse_event = json.loads(message) # 解析JSON指令
                        # 客户端发送的x/y是0-1之间的相对坐标,转换为绝对坐标
                        abs_x = int(mouse_event["x"] * screen_width)
                        abs_y = int(mouse_event["y"] * screen_height)

                        if mouse_event["type"] == "move":  # 鼠标移动事件
                            # 若坐标有变化,执行平滑移动 duration=0.05秒
                            if abs_x != current_x or abs_y != current_y:
                                pyautogui.moveTo(abs_x, abs_y, duration=0.05)
                                current_x, current_y = abs_x, abs_y # 更新当前坐标
                            # 处理鼠标按下状态,与客户端同步
                            if mouse_event.get("is_down", False) != is_mouse_down:
                                is_mouse_down = mouse_event["is_down"]
                                if is_mouse_down:
                                    pyautogui.mouseDown(button='left') # 按下左键
                                else:
                                    pyautogui.mouseUp(button='left') # 释放左键
                        elif mouse_event["type"] == "left_click":
                            pyautogui.click(button='left')
                        elif mouse_event["type"] == "right_click":
                            pyautogui.click(button='right')
                        elif mouse_event["type"] == "left_double_click":
                            pyautogui.click(button='left', clicks=2, interval=0.25)
                        elif mouse_event["type"] == "wheel":
                            direction = mouse_event["direction"]
                            scroll_delta = 100 if direction == "up" else -100
                            pyautogui.scroll(scroll_delta)
                            print(f"执行滚轮操作: {direction}")
                        elif mouse_event["type"] == "hwheel":
                            direction = mouse_event["direction"]
                            pyautogui.hscroll(100 if direction == "right" else -100)
                            print(f"执行水平滚轮操作: {direction}")

                        print(f"执行鼠标操作: {mouse_event['type']} 在坐标 ({abs_x}, {abs_y})")
                    except json.JSONDecodeError: # 处理无效JSON数据
                        print("收到无效的JSON数据")
                    except Exception as e: # 捕获其他异常
                        print(f"处理鼠标事件时出错: {e}")

    except Exception as e:
        print(f"处理客户端 {client_address} 鼠标控制时出错: {e}")
    finally:
        client_socket.close()
        print(f"客户端 {client_address} 鼠标控制连接已关闭")
# ================= 鼠标控制处理线程 =================


# ================= 键盘控制处理线程 =================
def handle_keyboard_client(client_socket, client_address):
    """处理键盘控制客户端的独立线程函数"""
    pressed_keys = {} # 存储按下的按键及其按下时间(用于重复按键处理)
    repeat_interval = 0.1 # 按键重复间隔秒,即按住不放时每0.1秒重复一次
    # 特殊按键映射表:将客户端发送的按键名称转换为keyboard库识别的名称
    special_keys = {
        'space': ' ', 'enter': 'enter', 'backspace': 'backspace',
        'delete': 'delete', 'tab': 'tab', 'escape': 'esc',
        'up': 'up', 'down': 'down', 'left': 'left', 'right': 'right',
        'shift': 'shift', 'ctrl': 'ctrl', 'alt': 'alt', 'caps_lock': 'caps_lock',
        'f1': 'f1', 'f2': 'f2', 'f3': 'f3', 'f4': 'f4',
        'f5': 'f5', 'f6': 'f6', 'f7': 'f7', 'f8': 'f8',
        'f9': 'f9', 'f10': 'f10', 'f11': 'f11', 'f12': 'f12'
    }

    try:
        print(f"开始处理客户端 {client_address} 的键盘控制请求")

        while True:
            data = client_socket.recv(1024) # 接收键盘指令
            if not data:
                break
            messages = data.decode('utf-8').split('\n') # 分割多条指令

            for message in messages:
                if message.strip():
                    try:
                        key_event = json.loads(message) # 解析JSON指令

                        if key_event.get("type") == "focus_lost": # 窗口失去焦点事件
                            # 释放所有已按下的按键
                            for key in list(pressed_keys.keys()):
                                keyboard.release(key)
                            pressed_keys.clear() # 清空按键状态
                            continue

                        key_name = key_event["name"] # 按键名称(如'enter','a')
                        event_type = key_event["type"] # 事件类型('key_down'或'key_up')
                        # 查找特殊按键映射,若无则使用原始名称
                        key_to_press = special_keys.get(key_name, key_name)

                        if event_type == "key_down": # 按键按下事件
                            if key_to_press not in pressed_keys: # 避免重复按下
                                keyboard.press(key_to_press) # 模拟按键按下
                                pressed_keys[key_to_press] = time.time() # 记录按下时间
                        elif event_type == "key_up": # 按键释放事件
                            if key_to_press in pressed_keys: # 避免释放未按下的按键
                                keyboard.release(key_to_press) # 模拟按键释放
                                del pressed_keys[key_to_press] # 从字典中移除

                    except Exception as e:
                        print(f"处理键盘事件时出错: {e}")

            # 处理按键重复逻辑(针对按住不放的按键)
            current_time = time.time()
            for key in list(pressed_keys.keys()): # 使用list()避免字典修改异常
                if current_time - pressed_keys[key] >= repeat_interval: # 达到重复间隔
                    keyboard.press(key) # 重复按下按键
                    pressed_keys[key] = current_time # 更新时间戳

    except Exception as e: # 捕获线程内异常
        print(f"处理客户端 {client_address} 键盘控制时出错: {e}")
    finally:
        # 确保释放所有残留按键
        for key in list(pressed_keys.keys()):
            keyboard.release(key)
        client_socket.close() # 关闭连接
        print(f"客户端 {client_address} 键盘控制连接已关闭")
# ================= 键盘控制处理线程 =================


# ================= gui界面 =================
def create_gui(stop_event):
    """创建Tkinter GUI界面"""
    root = Tk()
    root.title("F_RC") # 设置窗口名称
    root.geometry("500x120") # 设置窗口大小
    root.iconbitmap('exe.ico')
    #root.resizable(False, False) # 禁止调整窗口大小

    # 计算窗口居中位置
    screen_width = root.winfo_screenwidth() # 获取屏幕宽度,像素
    screen_height = root.winfo_screenheight() # 获取屏幕高度
    x = (screen_width - 500) // 2 # 水平居中坐标
    y = (screen_height - 120) // 2 # 垂直居中坐标
    root.geometry(f"500x120+{x}+{y}") # 设置窗口位置

    # 创建标签组件,显示程序名称
    Label(root, text=get_ipv6_address(), font=('黑体', 14, 'bold')).pack(pady=10)
    #Label(root, text="点击窗口关闭按钮退出程序", fg="red").pack(pady=5)

    def on_close():
        """窗口关闭按钮的回调函数"""
        print("GUI窗口关闭，程序将退出")
        stop_event.set() # 设置全局停止事件(通知其他线程退出)
        root.destroy() # 销毁窗口对象,释放资源

    root.protocol("WM_DELETE_WINDOW", on_close) # 绑定窗口关闭事件(点击标题栏关闭按钮时触发)
    root.mainloop() # 启动GUI主循环
# ================= gui界面 =================


# ================= 主函数 =================
def main(stop_event):
    """主服务函数，负责创建Socket并监听连接"""
    server_ip = get_ipv6_address() # 获取本机IP地址
    video_port = 8585
    mouse_port = 8586
    keyboard_port = 8587

    print(f"""
    =====================================
    IPv6远程控制服务端启动
    地址: [{server_ip}]
    视频端口: {video_port}
    鼠标端口: {mouse_port}
    键盘端口: {keyboard_port}
    =====================================
    """)

    # 创建视频流Socket
    video_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    video_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    video_socket.bind((server_ip, video_port, 0, 0))
    video_socket.listen(1)
    print("    视频服务器已启动,等待连接...")

    # 同理创建鼠标和键盘控制Socket
    mouse_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    mouse_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    mouse_socket.bind((server_ip, mouse_port, 0, 0))
    mouse_socket.listen(1)
    print("    鼠标控制服务器已启动,等待连接...")

    keyboard_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    keyboard_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    keyboard_socket.bind((server_ip, keyboard_port, 0, 0))
    keyboard_socket.listen(1)
    print("    键盘控制服务器已启动,等待连接...")
    print("\n")

    try:
        while not stop_event.is_set(): # 主循环,直到停止事件被设置
            # select.select 实现多路IO复用,监听三个Socket的可读事件
            # 参数:[读套接字列表], [写套接字列表], [错误套接字列表], 超时时间1秒
            readable, _, _ = select.select([video_socket, mouse_socket, keyboard_socket], [], [], 1.0)
            for sock in readable: # 遍历所有可读的Socket
                if sock is video_socket: # 视频客户端连接事件
                    client_socket, addr = video_socket.accept() # 接受连接
                    print(f"视频客户端已连接: {addr}")
                    # 创建守护线程处理客户端(daemon=True：主线程退出时强制终止子线程)
                    threading.Thread(
                        target=handle_video_client,
                        args=(client_socket, addr),
                        daemon=True
                    ).start()
                elif sock is mouse_socket:
                    client_socket, addr = mouse_socket.accept()
                    print(f"鼠标控制客户端已连接: {addr}")
                    threading.Thread(
                        target=handle_mouse_client,
                        args=(client_socket, addr),
                        daemon=True
                    ).start()
                elif sock is keyboard_socket:
                    client_socket, addr = keyboard_socket.accept()
                    print(f"键盘控制客户端已连接: {addr}")
                    threading.Thread(
                        target=handle_keyboard_client,
                        args=(client_socket, addr),
                        daemon=True
                    ).start()

    except Exception as e:
        print(f"主循环异常: {e}")
        traceback.print_exc()
    finally:
        video_socket.close()
        mouse_socket.close()
        keyboard_socket.close()
        print("所有服务器已关闭")
# ================= 主函数 =================



if __name__ == "__main__":
    # 屏蔽Ctrl+C信号
    signal.signal(signal.SIGINT, signal.SIG_IGN) # 忽略SIGINT信号

    stop_event = threading.Event() # 创建线程间通信的事件对象,用于通知关闭程序

    # 启动GUI线程
    gui_thread = threading.Thread(target=create_gui, args=(stop_event,), daemon=True)
    gui_thread.start()

    # 启动服务器线程
    server_thread = threading.Thread(target=main, args=(stop_event,), daemon=True)
    server_thread.start()

    # 主线程循环等待停止事件
    try:
        while not stop_event.is_set():
            time.sleep(0.1)
        print("收到退出信号，程序即将退出")
    finally:
        stop_event.set() # 确保设置停止标志
        server_thread.join(timeout=2.0) # 等待服务器线程最多2秒清理资源
        print("程序已完全退出")