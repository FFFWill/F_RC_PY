import socket
import cv2
import numpy as np
import threading
import time
import json
import keyboard
import win32gui  # 用于窗口焦点检测和设置窗口图标
import win32con  # 用于窗口常量
import os

# ==========================================================
# 全局变量定义
# ==========================================================
window_name = 'F_RC'
exit_event = threading.Event()
is_fullscreen = False
last_window_size = (0, 0)
server_address = None
server_port = 8585
mouse_socket = None
keyboard_socket = None
is_mouse_down = False
window_has_focus = False  # 窗口焦点状态

# 新增：焦点状态变更标志
focus_change_lock = threading.Lock()
last_focus_state = False  # 记录上一次焦点状态

# 新增：鼠标事件队列和处理逻辑
mouse_event_queue = []
mouse_event_lock = threading.Lock()
last_mouse_move_time = 0
MOUSE_MOVE_THROTTLE = 0.01  # 10ms，限制鼠标移动事件发送频率


# ==========================================================
# 窗口焦点处理函数
# ==========================================================
def get_focused_window_title():
    """获取当前焦点窗口的标题"""
    try:
        return win32gui.GetWindowText(win32gui.GetForegroundWindow())
    except:
        return ""


def check_window_focus():
    """持续检查窗口焦点状态并同步到服务器"""
    global window_has_focus, last_focus_state, keyboard_socket
    while not exit_event.is_set():
        focused_title = get_focused_window_title()
        current_focus = focused_title == window_name

        # 检测到焦点状态变化时
        with focus_change_lock:
            if current_focus != last_focus_state:
                last_focus_state = current_focus
                if keyboard_socket and not current_focus:  # 失去焦点时发送释放指令
                    try:
                        # 发送特殊焦点丢失事件
                        release_event = json.dumps({"type": "focus_lost"}).encode('utf-8') + b'\n'
                        keyboard_socket.sendall(release_event)
                        print("已通知服务器释放所有按键")
                    except Exception as e:
                        print(f"焦点状态同步失败: {e}")

        window_has_focus = current_focus
        time.sleep(0.1)  # 每100ms检查一次


# ==========================================================
# 鼠标事件处理函数
# ==========================================================
def process_mouse_events():
    """处理鼠标事件队列，合并高频移动事件"""
    global mouse_event_queue, mouse_socket, window_has_focus, last_mouse_move_time

    while not exit_event.is_set():
        if not window_has_focus or not mouse_socket:
            time.sleep(0.01)
            continue

        with mouse_event_lock:
            if not mouse_event_queue:
                time.sleep(0.01)
                continue

            # 优先处理非移动事件
            non_move_events = [e for e in mouse_event_queue if e["type"] != "move"]
            move_events = [e for e in mouse_event_queue if e["type"] == "move"]

            # 对于移动事件，只发送最新的一个（合并所有中间移动）
            final_move_event = move_events[-1] if move_events else None

            # 清空队列
            mouse_event_queue = []

        # 发送非移动事件
        for event in non_move_events:
            try:
                mouse_socket.sendall(json.dumps(event).encode('utf-8') + b'\n')
            except Exception as e:
                print(f"发送鼠标事件失败: {e}")

        # 限制移动事件发送频率
        current_time = time.time()
        if final_move_event and (current_time - last_mouse_move_time) >= MOUSE_MOVE_THROTTLE:
            try:
                mouse_socket.sendall(json.dumps(final_move_event).encode('utf-8') + b'\n')
                last_mouse_move_time = current_time
            except Exception as e:
                print(f"发送鼠标移动事件失败: {e}")

        # 短暂休眠，避免CPU占用过高
        time.sleep(0.001)


def mouse_callback(event, x, y, flags, param):
    """优化后的鼠标回调函数，使用事件队列"""
    global mouse_event_queue, is_mouse_down, window_has_focus, mouse_event_lock

    if not window_has_focus or not mouse_socket:
        return

    window_width = cv2.getWindowImageRect(window_name)[2]
    window_height = cv2.getWindowImageRect(window_name)[3]

    img_height, img_width = param[0], param[1]
    img_ratio = img_width / img_height
    window_ratio = window_width / window_height

    if img_ratio > window_ratio:
        display_width = window_width
        display_height = int(window_width / img_ratio)
    else:
        display_height = window_height
        display_width = int(window_height * img_ratio)

    x_offset = (window_width - display_width) // 2
    y_offset = (window_height - display_height) // 2

    if x >= x_offset and x < x_offset + display_width and y >= y_offset and y < y_offset + display_height:
        rel_x = (x - x_offset) / display_width
        rel_y = (y - y_offset) / display_height

        event_type = None
        if event == cv2.EVENT_LBUTTONDOWN:
            event_type = "left_click"
            is_mouse_down = True
        elif event == cv2.EVENT_RBUTTONDOWN:
            event_type = "right_click"
        elif event == cv2.EVENT_LBUTTONUP:
            event_type = "left_release"
            is_mouse_down = False
        elif event == cv2.EVENT_LBUTTONDBLCLK:
            event_type = "left_double_click"
        elif event == cv2.EVENT_MOUSEMOVE:
            event_type = "move"
        elif event == cv2.EVENT_MOUSEWHEEL:
            event_type = "wheel"
            wheel_direction = "up" if flags > 0 else "down"
        elif event == cv2.EVENT_MOUSEHWHEEL:
            event_type = "hwheel"
            wheel_direction = "right" if flags > 0 else "left"

        if event_type:
            mouse_event = {
                "type": event_type,
                "x": rel_x,
                "y": rel_y,
                "is_down": is_mouse_down
            }

            if event_type in ["wheel", "hwheel"]:
                mouse_event["direction"] = wheel_direction

            # 将事件添加到队列
            with mouse_event_lock:
                mouse_event_queue.append(mouse_event)


# ==========================================================
# 键盘事件处理函数
# ==========================================================
def keyboard_listener():
    """键盘监听函数，仅在窗口有焦点时发送事件"""
    global keyboard_socket, window_has_focus

    if not keyboard_socket:
        return

    def send_key_event(e):
        # 仅在窗口有焦点时发送键盘事件
        if window_has_focus:
            try:
                key_event = {
                    "type": "key_" + ("down" if e.event_type == keyboard.KEY_DOWN else "up"),
                    "name": e.name,
                    "scan_code": e.scan_code,
                    "time": e.time
                }
                keyboard_socket.sendall(json.dumps(key_event).encode('utf-8') + b'\n')
            except Exception as e:
                print(f"发送键盘事件失败: {e}")

    # 注册键盘事件回调
    keyboard.hook(send_key_event)

    # 保持线程运行
    while not exit_event.is_set():
        time.sleep(0.1)


# ==========================================================
# 窗口设置函数
# ==========================================================
def set_window_icon():
    """设置窗口图标"""
    try:
        # 获取窗口句柄
        hwnd = win32gui.FindWindow(None, window_name)

        # 检查图标文件是否存在
        icon_path = "exe.ico"
        if not os.path.exists(icon_path):
            print(f"警告: 图标文件 '{icon_path}' 不存在，使用默认图标")
            return

        # 加载图标资源
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            hicon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except Exception as e:
            print(f"加载图标失败: {e}，使用默认图标")
            return

        # 设置窗口图标（大图标和小图标）
        if hicon:
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon)
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon)
            print("窗口图标已设置")
    except Exception as e:
        print(f"设置窗口图标时出错: {e}")


# ==========================================================
# 视频帧接收和处理函数
# ==========================================================
def receive_frames():
    """接收视频帧并显示"""
    global is_fullscreen, last_window_size, mouse_socket, keyboard_socket, window_has_focus

    # 连接服务器的三个不同端口（视频、鼠标、键盘）
    video_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    video_socket.connect((server_address, server_port, 0, 0))

    mouse_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    mouse_socket.connect((server_address, server_port + 1, 0, 0))

    keyboard_socket = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    keyboard_socket.connect((server_address, server_port + 2, 0, 0))

    # 创建OpenCV窗口
    cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)

    # 启动窗口焦点检查线程
    focus_thread = threading.Thread(target=check_window_focus)
    focus_thread.daemon = True
    focus_thread.start()

    # 启动键盘监听线程
    keyboard_thread = threading.Thread(target=keyboard_listener)
    keyboard_thread.daemon = True
    keyboard_thread.start()

    # 启动鼠标事件处理线程
    mouse_thread = threading.Thread(target=process_mouse_events)
    mouse_thread.daemon = True
    mouse_thread.start()

    # 设置窗口图标（新增）
    # 延迟设置，确保窗口已创建
    threading.Timer(1.0, set_window_icon).start()

    try:
        while True:
            # 接收视频帧大小信息
            size_data = video_socket.recv(4)
            if not size_data:
                break
            size = int.from_bytes(size_data, byteorder='big')

            # 接收完整视频帧数据
            data = b''
            while len(data) < size:
                packet = video_socket.recv(size - len(data))
                if not packet:
                    break
                data += packet

            if not data:
                break

            # 解码并显示视频帧
            img_np = np.frombuffer(data, dtype=np.uint8)
            frame = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
            img_height, img_width = frame.shape[:2]

            # 获取窗口尺寸
            window_width = cv2.getWindowImageRect(window_name)[2]
            window_height = cv2.getWindowImageRect(window_name)[3]

            # 检测窗口大小变化
            if (window_width, window_height) != last_window_size:
                last_window_size = (window_width, window_height)
                print(f"窗口大小已调整为: {window_width}x{window_height}")

            # 根据窗口大小调整视频帧显示
            if window_width > 10 and window_height > 10:
                img_ratio = img_width / img_height
                window_ratio = window_width / window_height

                if img_ratio > window_ratio:
                    new_width = window_width
                    new_height = int(window_width / img_ratio)
                else:
                    new_height = window_height
                    new_width = int(window_height * img_ratio)

                # 调整视频帧大小并居中显示
                resized_frame = cv2.resize(frame, (new_width, new_height))
                background = np.zeros((window_height, window_width, 3), dtype=np.uint8)
                x_offset = (window_width - new_width) // 2
                y_offset = (window_height - new_height) // 2
                background[y_offset:y_offset + new_height, x_offset:x_offset + new_width] = resized_frame

                # 设置鼠标回调函数并显示图像
                cv2.setMouseCallback(window_name, mouse_callback, (img_height, img_width))
                cv2.imshow(window_name, background)
            else:
                # 窗口太小时直接显示原始帧
                cv2.setMouseCallback(window_name, mouse_callback, (img_height, img_width))
                cv2.imshow(window_name, frame)

            # 窗口刷新
            cv2.waitKey(1)

            # 检测窗口是否关闭
            window_visible = cv2.getWindowProperty(window_name, cv2.WND_PROP_VISIBLE)
            if window_visible < 1:
                exit_event.set()
                break

    except Exception as e:
        print(f"连接错误: {e}")
    finally:
        # 清理资源
        if video_socket:
            video_socket.close()
        if mouse_socket:
            mouse_socket.close()
        if keyboard_socket:
            keyboard_socket.close()
        cv2.destroyAllWindows()


# ==========================================================
# 主函数
# ==========================================================
def main():
    global server_address

    # 获取服务器地址并启动接收线程
    server_address = input("请输入服务器IPv6地址: ").strip()
    print(f"正在连接到服务器: [{server_address}]:{server_port}")

    receive_thread = threading.Thread(target=receive_frames)
    receive_thread.daemon = True
    receive_thread.start()

    # 等待退出事件
    try:
        while not exit_event.is_set():
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("程序已退出")
    finally:
        # 清理资源
        cv2.destroyAllWindows()
        exit_event.set()


# ==========================================================
# 程序入口
# ==========================================================
if __name__ == "__main__":
    main()

"""
程序初始化
从用户输入获取服务器 IPv6 地址
启动接收线程处理视频流和用户输入
视频流接收与显示receive_frames函数
连接服务器的三个端口：视频 (8585)、鼠标 (8586)、键盘 (8587)
创建 OpenCV 窗口显示远程桌面
启动焦点检查线程和键盘监听线程
循环接收视频帧：
接收帧大小信息（4 字节）
接收完整帧数据
解码并根据窗口大小自适应显示
窗口焦点管理check_window_focus函数
每 100ms 检查一次窗口焦点状态
当窗口失去焦点时，发送特殊事件通知服务器释放所有按键
鼠标事件处理mouse_callback函数
仅在窗口有焦点时处理鼠标事件
计算鼠标在远程桌面中的相对位置
处理点击、双击、移动和滚轮事件
将事件发送到服务器
键盘事件处理keyboard_listener函数
仅在窗口有焦点时处理键盘事件
监听所有键盘按键的按下和释放事件
将按键事件发送到服务器
窗口管理set_window_icon函数
设置窗口图标（如果存在 exe.ico 文件）
处理窗口大小变化和关闭事件
"""