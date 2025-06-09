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

# ================= 管理员权限适配部分 =================
def is_admin():
    """检查当前程序是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()  # 调用Windows API检查管理员权限
    except:
        return False  # 异常时默认非管理员权限


def request_admin_privileges():
    """请求提升程序至管理员权限"""
    if not is_admin():  # 若当前非管理员权限
        print("正在请求管理员权限...")
        # 使用Windows API重新以管理员身份启动当前脚本
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,  # 当前Python解释器路径
            __file__,  # 脚本文件路径
            None,
            1
        )
        sys.exit()  # 退出当前非管理员进程


# 立即检查并请求管理员权限
request_admin_privileges()


# ================= 管理员权限适配部分 =================


def get_public_ip():
    """获取公网IPv4地址（优先UDP协议，失败则解析ipconfig）"""
    try:
        # 使用UDP协议获取公网IP（不会实际发送数据）
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception as e:
        print(f"UDP获取IP失败: {e}，尝试解析ipconfig")
        try:
            # 解析ipconfig输出获取IPv4地址
            output = os.popen("ipconfig /all").read()
            pattern = re.compile(r"IPv4 Address.*?(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})", re.I)
            matches = pattern.findall(output)
            for ip in matches:
                if not ip.startswith(('127.', '192.168.', '10.', '172.16.')):
                    return ip
            # 未找到公网IP时返回第一个有效地址
            return matches[0] if matches else "127.0.0.1"
        except Exception as e:
            print(f"解析ipconfig失败: {e}")
            return "127.0.0.1"


# 定义画质配置字典（帧率范围对应分辨率、画质质量、描述）
QUALITY_CONFIG = {
    (0, 5): (1280, 720, 10, "垃圾帧率: 1280×720 质量10"),
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
    (60, float('inf')): (1920, 1080, 100, "最高帧率: 1920×1080 质量100")
}


class VideoQualityManager:
    """视频质量管理器（根据帧率动态调整分辨率和画质）"""

    def __init__(self):
        self.current_config = QUALITY_CONFIG[(25, 30)]
        self.last_adjust_time = time.time()
        self.adjust_interval = 1

    def get_config(self, current_fps):
        for fps_range, config in QUALITY_CONFIG.items():
            if fps_range[0] <= current_fps < fps_range[1]:
                return config
        return self.current_config

    def adjust_quality(self, current_fps):
        current_time = time.time()
        if current_time - self.last_adjust_time > self.adjust_interval:
            new_config = self.get_config(current_fps)
            if new_config != self.current_config:
                self.current_config = new_config
                print(f"画质调整: {new_config[3]} (当前帧率: {current_fps:.1f} FPS)")
            self.last_adjust_time = current_time
        return self.current_config


def capture_screen():
    """捕获屏幕画面（使用mss库，返回BGR格式的numpy数组）"""
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        sct_img = sct.grab(monitor)
        img = np.array(sct_img)
        return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)


def handle_video_client(client_socket, client_address):
    try:
        print(f"开始处理客户端 {client_address} 的视频请求")
        MAX_FPS = 60
        frame_interval = 1.0 / MAX_FPS
        last_frame_time = time.time()

        last_second = int(time.time())
        capture_count = 0
        process_count = 0
        process_time_sum = 0.0
        send_count = 0

        with mss.mss() as sct:
            monitor = sct.monitors[1]
            original_width = monitor["width"]
            original_height = monitor["height"]
            quality_manager = VideoQualityManager()

            while True:
                now = time.time()
                elapsed = now - last_frame_time
                if elapsed < frame_interval:
                    time.sleep(frame_interval - elapsed)
                    continue
                last_frame_time = now

                current_second = int(now)
                if current_second > last_second:
                    print(f"\n[统计 {last_second}s-{current_second - 1}s] "
                          f"截取帧数: {capture_count} "
                          f"处理帧数: {process_count} "
                          f"处理耗时: {process_time_sum:.1f}ms "
                          f"发送帧数: {send_count} ")
                    capture_count = 0
                    process_count = 0
                    process_time_sum = 0.0
                    send_count = 0
                    last_second = current_second

                capture_start = time.time()
                frame = capture_screen()
                capture_time = (time.time() - capture_start) * 1000
                capture_count += 1

                width, height, quality, _ = quality_manager.adjust_quality(send_count)
                resize_start = time.time()
                if (width, height) != (original_width, original_height):
                    frame = cv2.resize(frame, (width, height))
                resize_time = (time.time() - resize_start) * 1000

                encode_start = time.time()
                _, img_encoded = cv2.imencode('.jpg', frame, [int(cv2.IMWRITE_JPEG_QUALITY), quality])
                encode_time = (time.time() - encode_start) * 1000
                process_count += 1
                process_time_sum += resize_time + encode_time

                data = img_encoded.tobytes()
                size = len(data)
                try:
                    client_socket.sendall(size.to_bytes(4, byteorder='big'))
                    client_socket.sendall(data)
                    send_count += 1
                except (ConnectionResetError, ConnectionAbortedError):
                    print(f"客户端 {client_address} 主动断开连接")
                    break

                print(f"[实时] 捕获: {capture_time:.1f}ms "
                      f"处理: {resize_time + encode_time:.1f}ms "
                      f"队列延迟: {elapsed * 1000:.1f}ms")

    except Exception as e:
        print(f"处理客户端 {client_address} 时出错: {e}")
        traceback.print_exc()
    finally:
        client_socket.close()
        print(f"客户端 {client_address} 视频连接已关闭")


def handle_mouse_client(client_socket, client_address):
    try:
        print(f"开始处理客户端 {client_address} 的鼠标控制请求")
        screen_width, screen_height = pyautogui.size()
        pyautogui.PAUSE = 0.0
        pyautogui.FAILSAFE = True
        current_x, current_y = 0, 0
        is_mouse_down = False

        while True:
            data = client_socket.recv(1024)
            if not data:
                break
            messages = data.decode('utf-8').split('\n')

            for message in messages:
                if message.strip():
                    try:
                        mouse_event = json.loads(message)
                        abs_x = int(mouse_event["x"] * screen_width)
                        abs_y = int(mouse_event["y"] * screen_height)

                        if mouse_event["type"] == "move":
                            if abs_x != current_x or abs_y != current_y:
                                pyautogui.moveTo(abs_x, abs_y, duration=0.05)
                                current_x, current_y = abs_x, abs_y
                            if mouse_event.get("is_down", False) != is_mouse_down:
                                is_mouse_down = mouse_event["is_down"]
                                if is_mouse_down:
                                    pyautogui.mouseDown(button='left')
                                else:
                                    pyautogui.mouseUp(button='left')
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
                    except json.JSONDecodeError:
                        print("收到无效的JSON数据")
                    except Exception as e:
                        print(f"处理鼠标事件时出错: {e}")

    except Exception as e:
        print(f"处理客户端 {client_address} 鼠标控制时出错: {e}")
    finally:
        client_socket.close()
        print(f"客户端 {client_address} 鼠标控制连接已关闭")


def handle_keyboard_client(client_socket, client_address):
    pressed_keys = {}
    repeat_interval = 0.1
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
            data = client_socket.recv(1024)
            if not data:
                break
            messages = data.decode('utf-8').split('\n')

            for message in messages:
                if message.strip():
                    try:
                        key_event = json.loads(message)

                        if key_event.get("type") == "focus_lost":
                            for key in list(pressed_keys.keys()):
                                keyboard.release(key)
                            pressed_keys.clear()
                            continue

                        key_name = key_event["name"]
                        event_type = key_event["type"]
                        key_to_press = special_keys.get(key_name, key_name)

                        if event_type == "key_down":
                            if key_to_press not in pressed_keys:
                                keyboard.press(key_to_press)
                                pressed_keys[key_to_press] = time.time()
                        elif event_type == "key_up":
                            if key_to_press in pressed_keys:
                                keyboard.release(key_to_press)
                                del pressed_keys[key_to_press]

                    except Exception as e:
                        print(f"处理键盘事件时出错: {e}")

            current_time = time.time()
            for key in list(pressed_keys.keys()):
                if current_time - pressed_keys[key] >= repeat_interval:
                    keyboard.press(key)
                    pressed_keys[key] = current_time

    except Exception as e:
        print(f"处理客户端 {client_address} 键盘控制时出错: {e}")
    finally:
        for key in list(pressed_keys.keys()):
            keyboard.release(key)
        client_socket.close()
        print(f"客户端 {client_address} 键盘控制连接已关闭")


# 创建GUI窗口并使用事件标志通知主线程
def create_gui(stop_event):
    """创建简易GUI窗口"""
    root = Tk()
    root.title("远程控制服务端")
    root.geometry("500x120")
    root.iconbitmap('exe.ico')
    #root.resizable(False, False)

    # 居中显示窗口
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 500) // 2
    y = (screen_height - 120) // 2
    root.geometry(f"500x120+{x}+{y}")

    # 显示固定文本
    Label(root, text=get_public_ip(), font=('黑体', 14, 'bold')).pack(pady=10)
    #label.pack(pady=20)

    # 关闭窗口时设置事件并退出
    def on_close():
        print("GUI窗口关闭，程序将退出")
        stop_event.set()  # 设置事件标志通知主线程
        root.destroy()    # 销毁Tkinter窗口

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


def main(stop_event):
    """主函数（启动服务器，监听多端口连接）"""
    server_ip = get_public_ip()  # 获取公网IPv4地址
    video_port = 8585
    mouse_port = 8586
    keyboard_port = 8587

    print(
        "\n\n",
        f"服务器启动",
        f"公网IPv4地址: {server_ip}",
        "\n\n",
        f"视频端口: {video_port}, 鼠标控制端口: {mouse_port}, 键盘控制端口: {keyboard_port}"
        "\n\n",
    )

    # 创建并绑定视频服务器套接字（IPv4）
    video_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    video_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    video_socket.bind((server_ip, video_port))  # 绑定公网IP和端口
    video_socket.listen(1)
    print("视频服务器已启动,等待连接...")

    # 创建并绑定鼠标控制服务器套接字
    mouse_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    mouse_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    mouse_socket.bind((server_ip, mouse_port))
    mouse_socket.listen(1)
    print("鼠标控制服务器已启动,等待连接...")

    # 创建并绑定键盘控制服务器套接字
    keyboard_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    keyboard_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    keyboard_socket.bind((server_ip, keyboard_port))
    keyboard_socket.listen(1)
    print("键盘控制服务器已启动,等待连接...")
    print("\n")

    try:
        while not stop_event.is_set():  # 检查事件标志
            # 使用select实现带超时的监听，以便定期检查事件标志
            readable, _, _ = select.select([video_socket, mouse_socket, keyboard_socket], [], [], 1.0)
            for sock in readable:
                if sock is video_socket:
                    client_socket, addr = video_socket.accept()
                    print(f"视频客户端已连接: {addr}")
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

    except KeyboardInterrupt:
        print("\n服务器关闭")
    except Exception as e:
        print(f"主循环异常: {e}")
        traceback.print_exc()
    finally:
        # 关闭所有套接字
        video_socket.close()
        mouse_socket.close()
        keyboard_socket.close()
        print("所有服务器已关闭")


if __name__ == "__main__":
    # 创建一个事件对象用于线程间通信
    stop_event = threading.Event()

    # 设置主线程的信号处理
    signal.signal(signal.SIGINT, lambda signum, frame: stop_event.set())
    print("提示: 按Ctrl+C也可以退出程序")

    # 启动GUI线程，传入事件对象
    gui_thread = threading.Thread(target=create_gui, args=(stop_event,), daemon=True)
    gui_thread.start()

    # 启动服务器主线程，传入事件对象
    server_thread = threading.Thread(target=main, args=(stop_event,), daemon=True)
    server_thread.start()

    # 主线程等待事件标志被设置
    try:
        while not stop_event.is_set():
            time.sleep(0.1)
        print("收到退出信号，程序即将退出")
    except KeyboardInterrupt:
        print("\n用户中断，程序退出")
    finally:
        # 设置事件标志（如果尚未设置）
        stop_event.set()
        # 等待服务器线程完成清理工作
        server_thread.join(timeout=2.0)
        print("程序已完全退出")