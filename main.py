#作者：Github@YuLuoQAQ
import time
import cv2
import threading
import win32gui
import pyautogui
import sys

def pretreatmentHandle():
    hwnd = win32gui.FindWindow("Progman", "Program Manager")
    win32gui.SendMessageTimeout(hwnd, 0x052C, 0, None, 0, 0x03E8)
    hwnd_WorkW = None
    while 1:
        hwnd_WorkW = win32gui.FindWindowEx(None, hwnd_WorkW, "WorkerW", None)
        print('hwmd_workw: ', hwnd_WorkW)
        if not hwnd_WorkW:
            continue
        hView = win32gui.FindWindowEx(hwnd_WorkW, None, "SHELLDLL_DefView", None)
        print('hwmd_hView: ', hView)
        if not hView:
            continue
        h = win32gui.FindWindowEx(None, hwnd_WorkW, "WorkerW", None)
        print('h_1: ', h)
        while h:
            win32gui.SendMessage(h, 0x0010, 0, 0)  # WM_CLOSE
            h = win32gui.FindWindowEx(None, hwnd_WorkW, "WorkerW", None)
            print(h)
        break
    return hwnd

def play_video(video_path):
    try:
        cap = cv2.VideoCapture(video_path)
        cv2.namedWindow('Video Player', cv2.WINDOW_NORMAL)
        cv2.setWindowProperty('Video Player', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
        while True:
            while cap.isOpened():
                ret, frame = cap.read()
                if ret:
                    cv2.imshow('Video Player', frame)
                    if cv2.waitKey(25) & 0xFF == ord('q'):
                        break
                else:
                    # 重置视频帧到起始位置
                    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                    break
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
    except Exception as e:
        pyautogui.alert(e)
    finally:
        cv2.destroyAllWindows()
        default_video_path = ''
        play_video(default_video_path)

def get_hwnd(window_class, window_title):
    for i in range(60):
        hwnd = win32gui.FindWindow(window_class, window_title)
        if hwnd != 0:
            return hwnd
        else:
            print("未找到指定窗口")
            time.sleep(1)
    else:
        raise TimeoutError('未找到指定窗口')

if __name__ == "__main__":
    video_path = pyautogui.prompt('请输入视频文件路径/URL。\n支持图片和视频。', '选择壁纸')
    if not video_path:
        video_path = ''
    video_thread = threading.Thread(target=play_video, args=(video_path,))
    video_thread.start()
    pretreatmentHandle()
    window_class = "Main HighGUI class"  # 窗口类名
    window_title = "Video Player"  # 窗口标题
    video_h = get_hwnd(window_class, window_title)
    print("Window Handle:", video_h)
    h = win32gui.FindWindow(("Progman"), ("Program Manager"))
    win32gui.SetParent(video_h, h)
