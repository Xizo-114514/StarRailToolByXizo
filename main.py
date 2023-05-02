import ctypes
import json
import re
import subprocess
import requests
import time
import zipfile
import os
import webbrowser
import winreg
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as msg
import base64
from files import fileszip

def readreg():
    key_path = r"Software\miHoYo\崩坏：星穹铁道"
    value_name1 = "GraphicsSettings_Model_h2986158309"
    value_name2 = "GraphicsSettings_PCResolution_h431323223"

    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)

    original_value1, _ = winreg.QueryValueEx(key, value_name1)
    original_value2, _ = winreg.QueryValueEx(key, value_name2)

    winreg.CloseKey(key)

    print("初始：", original_value1)
    print("初始：", original_value2)

    value1 = original_value1.decode().rstrip('\x00')
    value2 = original_value2.decode().rstrip('\x00')

    json_value1 = json.loads(value1)
    json_value2 = json.loads(value2)

    FPS = json_value1["FPS"]
    RenderScale = json_value1["RenderScale"]
    width = json_value2["width"]
    height = json_value2["height"]
    resolution = str(width)+'x'+str(height)
    isFullScreen = json_value2["isFullScreen"]

    return FPS,RenderScale,resolution,isFullScreen

def change(FPS = 120, RenderScale = 1.4, width = 2560, height = 1440, isFullScreen = 'True'):
    key_path = r"Software\miHoYo\崩坏：星穹铁道"
    value_name1 = "GraphicsSettings_Model_h2986158309"
    value_name2 = "GraphicsSettings_PCResolution_h431323223"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        original_value1, _ = winreg.QueryValueEx(key, value_name1)
        original_value2, _ = winreg.QueryValueEx(key, value_name2)
    except:
        msg.showerror("错误", "修改失败，没有找到注册表项。\n请先进入游戏设置，手动修改任意画质选项。")

    print("当前：", original_value1)
    print("当前：", original_value2)

    value1 = original_value1.decode().rstrip('\x00')
    value2 = original_value2.decode().rstrip('\x00')

    json_value1 = json.loads(value1)
    json_value2 = json.loads(value2)

    json_value1["FPS"] = FPS
    json_value1["RenderScale"] = RenderScale
    json_value2["width"] = width
    json_value2["height"] = height
    json_value2["isFullScreen"] = isFullScreen

    str_value1 = json.dumps(json_value1, separators=(',', ':')) + '\x00'
    bytes_value1 = str_value1.encode()
    str_value2 = json.dumps(json_value2, separators=(',', ':')) + '\x00'
    bytes_value2 = str_value2.encode()

    winreg.SetValueEx(key, value_name1, 0, winreg.REG_BINARY, bytes_value1)
    print("修改：", bytes_value1)
    winreg.SetValueEx(key, value_name2, 0, winreg.REG_BINARY, bytes_value2)
    print("修改：", bytes_value2)

    winreg.CloseKey(key)

def install():
    resolution = choose_resolution_combobox.get()
    width = resolution.split('x')[0]
    height = resolution.split('x')[1]
    isFullScreen = str(bool(CheckIsFullScreen.get()))
    fps = choose_fps_combobox.get()
    RenderScale = choose_RenderScale_combobox.get()
    change(fps, RenderScale, width, height, isFullScreen)

def find_exe():
    choose_exe_bar.configure(state='normal')
    choose_exe_bar.delete("1.0", "end")
    choose_exe_bar.insert("insert", "查找路径中...请稍等一下~\n此时请不要点击窗口")
    choose_exe_bar.configure(state='disabled')
    choose_exe_bar.update()
    exe_name = "StarRail.exe"
    for drive in ['C:\\', 'D:\\', 'E:\\', 'F:\\', 'G:\\', 'H:\\', 'I:\\', 'J:\\', 'K:\\', 'L:\\', 'M:\\', 'N:\\', 'O:\\', 'P:\\', 'Q:\\', 'R:\\', 'S:\\', 'T:\\', 'U:\\', 'V:\\', 'W:\\', 'X:\\', 'Y:\\', 'Z:\\']:
        for root, dirs, files in os.walk(drive):
            if exe_name in files:
                exe_path = os.path.join(root, exe_name)
                print(exe_path)
                choose_exe_bar.configure(state='normal')
                choose_exe_bar.delete("1.0", "end")
                choose_exe_bar.insert("insert", exe_path)
                choose_exe_bar.configure(state='disabled')
                config = open(appdata + r"\config.ini.xizo", "w")
                config.write(exe_path)
                config.close()
                return None
            else:
                choose_exe_bar.configure(state='normal')
                choose_exe_bar.delete("1.0", "end")
                choose_exe_bar.insert("insert", "未找到StarRail.exe\n请确认游戏是否安装")
                choose_exe_bar.configure(state='disabled')

def get_pid(name):
    pid = None
    for line in os.popen('tasklist /fi "imagename eq %s"' % name).readlines():
        m = re.match(r"(\S+)\s+(\d+)", line.strip())
        if m and m.group(1).lower() == name.lower():
            pid = int(m.group(2))
            break
    return pid

def start():
    os.startfile(str(choose_exe_bar.get(1.0, END))[:-1])

def end():
    pid = get_pid('StarRail.exe')
    print(str(pid))
    PROCESS_TERMINATE = 1
    handle = ctypes.windll.kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
    ctypes.windll.kernel32.TerminateProcess(handle, -1)
    ctypes.windll.kernel32.CloseHandle(handle)
def restart():
    end()
    time.sleep(1)
    start()

# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':

    # 初次启动生成目录释放文件 内有程序贴图 配置文件
    appdata = os.getenv("APPDATA") + r"\_XizoStarRailTool"
    if os.path.exists(appdata) == False:
        os.makedirs(appdata)
        filezip = open(appdata + r"\files.zip", "wb+")
        filezip.write(base64.b64decode(fileszip))
        filezip.close()
        zip_file = zipfile.ZipFile(appdata + r"\files.zip")
        zip_file.extractall(appdata)
        zip_file.close()
        os.remove(appdata + r"\files.zip")

    config = open(appdata + r"\config.ini.xizo", "r")
    path_config = config.read()
    config.close()

    try:
        init = readreg()
        fps = init[0]
        RenderScale = init[1]
        resolution = init[2]
        isFullScreen = init[3]
    except:
        None

    #主窗口
    root = tk.Tk()
    root.title("星球铁道画质修改器")

    root.overrideredirect(True)  # 无标题栏

    #窗口位置大小
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    width = 650
    height = 400
    x = int((screenwidth - width) / 2)
    y = int((screenheight - height) / 2)
    root.geometry("%dx%d+%d+%d" % (width, height, x, y))
    root.resizable(0, 0)

    #图标
    root.iconbitmap(appdata + r"\icon.ico")

    # 设置背景色使其替换为透明配合圆角
    root["background"] = "#f1eeee"
    root.attributes("-transparentcolor", "#f1eeee")

    # 创建移动窗口的函数
    def MouseXYonWindow(MouseXY):
        global MouseX, MouseY
        MouseX, MouseY = MouseXY.x, MouseXY.y

    def MouseXYonScreen(MouseXY):
        root.geometry(f'+{MouseXY.x_root - MouseX}+{MouseXY.y_root - MouseY}')

    # 圆角背景
    backgroundphoto = tk.PhotoImage(file=appdata + r'/bg.png')  # 图片里做了圆角添加了标题
    backgroundLabel = tk.Label(root, image=backgroundphoto, bg='#f1eeee')  # 将lable的底色改为背景色使其透明将上部图片的圆角显示出来
    backgroundLabel.place(width=650, height=400)
    backgroundLabel.bind("<Button-1>", MouseXYonWindow)  # 绑定函数
    backgroundLabel.bind("<B1-Motion>", MouseXYonScreen)
    backgroundLabel.backgroundphoto = backgroundphoto  # 不能没有这个 没有会有显示问题

    # 退出程序函数 破坏主窗体
    def exitprogram():
        root.destroy()

    def minimize():
        root.overrideredirect(False)
        root.iconify()

    def on_map(event):
        root.overrideredirect(True)
    root.bind("<Motion>", on_map)

    # 退出按钮
    exitbtnphoto = tk.PhotoImage(file=appdata + r'/exitbtn.png')
    exit_btn = tk.Button(root, image=exitbtnphoto, bd=0, activebackground='#ffffff', command=exitprogram)
    exit_btn.pack(padx=10, pady=10, anchor=tk.NE)
    exit_btn.exitbtnphoto = exitbtnphoto  # 不能没有这个 没有会有显示问题

    # 最小化按钮
    exitbtnphoto = tk.PhotoImage(file=appdata + r'/minimizebtn.png')
    exit_btn = tk.Button(root, image=exitbtnphoto, bd=0, activebackground='#ffffff', command=minimize)
    exit_btn.pack(padx=10, pady=20, anchor=tk.NE)
    exit_btn.exitbtnphoto = exitbtnphoto  # 不能没有这个 没有会有显示问题

    CheckIsInTop = tk.IntVar(master=root)
    IsInTop_checkbtn = tk.Checkbutton(root, variable=CheckIsInTop,onvalue=1, offvalue=0)

    def CheckIsInTopClick(event):
        if CheckIsInTop.get() == 1:
            root.attributes("-topmost", 1)
            open(appdata + r"\isintop.xizo", "w")
        if CheckIsInTop.get() == 0:
            root.attributes("-topmost", 0)
            os.remove(appdata + r"\isintop.xizo")
    IsInTop_checkbtn.bind("<Motion>", CheckIsInTopClick)  # 绑定函数

    if os.path.exists(appdata + r"\isintop.xizo") == False:
        root.attributes("-topmost", 0)
        IsInTop_checkbtn.deselect()
    else:
        root.attributes("-topmost", 1)
        IsInTop_checkbtn.select()

    errorcode = 0

    CheckIsFullScreen = tk.IntVar(master=root)
    IsFullScreen_checkbtn = tk.Checkbutton(root, variable=CheckIsFullScreen,onvalue=1, offvalue=0)
    try:
        if isFullScreen == 'True':
            IsFullScreen_checkbtn.select()
    except:
        errorcode = 1

    #下拉框
    resolution_list = ["3840x2160", "2560x1440", "1920x1440", "1920x1200", "1920x1080", "1680x1050", "1600x1200", "1600x1024", "1600x900", "1440x1080", "1440x900", "1366x768", "1360x768", "1280x1440", "1280x1024", "1280x960", "1280x800", "1280x768", "1176x664", "1152x864", "1024x768", "800x600", "720x576", "720x480", "640x480"]
    resolution_var = StringVar()
    try:
        resolution_var.set(resolution)
    except:
        errorcode = 1

    choose_resolution_combobox = ttk.Combobox(root, justify=tk.CENTER, height=10, width=9,
    state="readonly", font=("微软雅黑", 10), values=resolution_list, textvariable=resolution_var)

    fps_list = [60,120]
    fps_var = StringVar()
    try:
        fps_var.set(fps)
    except:
        errorcode = 1

    choose_fps_combobox = ttk.Combobox(root, justify=tk.CENTER, height=10, width=3,
    state="readonly", font=("微软雅黑", 10), values=fps_list, textvariable=fps_var)

    RenderScale_list = [0.6,0.8,1.0,1.2,1.4,1.6,1.8,2.0]
    RenderScale_var = StringVar()
    try:
        RenderScale_var.set(RenderScale)
    except:
        errorcode = 1

    if errorcode == 1:
        msg.showerror("错误", "没有找到注册表项。\n请先进入游戏设置，手动修改任意画质选项。")

    choose_RenderScale_combobox = ttk.Combobox(root, justify=tk.CENTER, height=10, width=8,
    state="readonly", font=("微软雅黑", 10), values=RenderScale_list, textvariable=RenderScale_var)

    choose_exe_bar = tk.Text(root, height=2,
    state="disabled", font=("微软雅黑", 10))
    choose_exe_bar.configure(state='normal')
    choose_exe_bar.insert("insert", "游戏StarRail.exe路径\n       (点击自动查找)→")
    choose_exe_bar.configure(state='disabled')

    #Buttons
    install_btn = tk.Button(root, text="应用", command=install, fg='black', bg='light blue', font=("微软雅黑", 10))
    find_exe_btn = tk.Button(root, text="查找\n路径", command=find_exe, font=("微软雅黑", 9))
    start_btn = tk.Button(root, text="启动", command=start, fg='black', bg='light green', font=("微软雅黑", 10))
    end_btn = tk.Button(root, text="结束", command=end, fg='black', bg='pink', font=("微软雅黑", 10))
    restart_btn = tk.Button(root, text="重启", command=restart, fg='black', bg='light yellow', font=("微软雅黑", 10))

    #Place elements
    choose_resolution_combobox.place(x=225, y=40)
    choose_fps_combobox.place(x=329, y=40)
    choose_RenderScale_combobox.place(x=385, y=40)
    IsFullScreen_checkbtn.place(x=490, y=38)
    install_btn.place(x=540, y=10, width=56, height=55)
    choose_exe_bar.place(x=445, y=110, width=160, height=40)
    find_exe_btn.place(x=605, y=110)
    start_btn.place(x=445, y=160, width=60)
    end_btn.place(x=513, y=160, width=60)
    restart_btn.place(x=580, y=160, width=60)
    IsInTop_checkbtn.place(x=610, y=250)

    if path_config != "":
        print(path_config,type(path_config))
        choose_exe_bar.configure(state='normal')
        choose_exe_bar.delete("1.0", "end")
        choose_exe_bar.insert("insert", path_config)
        choose_exe_bar.configure(state='disabled')

    root.mainloop()
