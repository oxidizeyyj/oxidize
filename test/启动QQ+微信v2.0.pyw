import os
import subprocess
import time

try:
    import pyautogui
    import pygetwindow as gw
except Exception:
    raise SystemExit("缺少依赖，请先运行：pip install pyautogui pygetwindow")

# 快捷方式路径
qq_shortcut = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\腾讯软件\QQ\QQ.lnk"
wechat_shortcut = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\微信\微信.lnk"

# 微信窗口配置
WAIT_TIMEOUT = 30          # 等待窗口出现时间
AFTER_WINDOW_APPEAR = 2.0  # 窗口出现后等待几秒让界面完全加载
ENTER_RETRIES = 2          # 微信回车次数
DELAY_BETWEEN_KEYS = 0.2

def start_app(shortcut_path, name):
    """启动程序，不发送回车"""
    if not os.path.exists(shortcut_path):
        print(f"未找到 {name} 快捷方式，请检查路径：{shortcut_path}")
        return

    print(f"正在启动 {name}...")
    subprocess.Popen(shortcut_path, shell=True)
    # 可以根据需要等待几秒保证启动，但不发送回车
    time.sleep(1)
    print(f"{name} 已启动。")

def start_wechat_and_enter(shortcut_path, name, window_keys):
    """启动微信，等待窗口出现并发送回车"""
    if not os.path.exists(shortcut_path):
        print(f"未找到 {name} 快捷方式，请检查路径：{shortcut_path}")
        return

    print(f"正在启动 {name}...")
    subprocess.Popen(shortcut_path, shell=True)

    # 等待微信窗口出现
    win = None
    start_time = time.time()
    while time.time() - start_time < WAIT_TIMEOUT:
        for key in window_keys:
            wins = gw.getWindowsWithTitle(key)
            if wins:
                win = wins[0]
                break
        if win:
            break
        time.sleep(0.5)

    if not win:
        print(f"{name} 窗口未检测到")
        return

    # 确保窗口置顶并激活
    print(f"{name} 窗口检测到，激活窗口...")
    try:
        win.restore()
        win.activate()
        win.maximize()
    except Exception:
        pass

    # 等待界面完全加载
    time.sleep(AFTER_WINDOW_APPEAR)

    # 发送回车尝试登录
    print(f"向 {name} 发送回车尝试登录...")
    for _ in range(ENTER_RETRIES):
        pyautogui.press('enter')
        time.sleep(DELAY_BETWEEN_KEYS)

    print(f"{name} 登录操作完成。")

if __name__ == "__main__":
    # 1. 启动 QQ（不发送回车）
    start_app(qq_shortcut, "QQ")
    # 等待几秒，确保 QQ 已经启动完成
    time.sleep(2)

    # 2. 启动微信并发送回车登录
    start_wechat_and_enter(wechat_shortcut, "微信", ["WeChat", "微信"])

    print("全部操作完成。")
