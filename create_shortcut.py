import os
import sys
import winreg
import pythoncom
from win32com.client import Dispatch
import tkinter as tk
from tkinter import messagebox

def create_shortcut():
    try:
        # 获取当前脚本的完整路径
        script_path = os.path.abspath("context_menu_manager.py")
        current_dir = os.path.dirname(script_path)
        
        # 创建快捷方式的目标路径（放在同级目录）
        shortcut_path = os.path.join(current_dir, "右键菜单管理器.lnk")
        
        # 创建快捷方式
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{script_path}"'
        shortcut.WorkingDirectory = current_dir
        shortcut.IconLocation = "shell32.dll,0"  # Windows 默认图标
        shortcut.save()
        
        # 创建一个批处理文件，用于以管理员权限运行
        bat_path = os.path.join(current_dir, "启动右键菜单管理器.bat")
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write('@echo off\n')
            f.write('cd /d "%~dp0"\n')  # 切换到批处理文件所在目录
            f.write(f'python "{os.path.basename(script_path)}"\n')
            f.write('pause')
        
        messagebox.showinfo("成功", "快捷方式和启动文件已创建在程序目录下！")
    except Exception as e:
        messagebox.showerror("错误", f"创建快捷方式失败：{str(e)}")

if __name__ == "__main__":
    create_shortcut() 