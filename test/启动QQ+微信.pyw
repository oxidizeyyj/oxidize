import os
import subprocess
import time

qq_shortcut = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\腾讯软件\QQ\QQ.lnk"
wechat_shortcut = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\微信\微信.lnk"

def start_app(shortcut_path, name):
    if os.path.exists(shortcut_path):
        print(f"正在启动 {name}...")
        subprocess.Popen(shortcut_path, shell=True)
        time.sleep(1)
    else:
        print(f"未找到 {name} 快捷方式，请检查路径：{shortcut_path}")

start_app(qq_shortcut, "QQ")
start_app(wechat_shortcut, "微信")

print("全部启动完成。")
