import tkinter as tk
from tkinter import ttk
import winreg
import datetime
import os
from tkinter import messagebox
import win32gui
import win32ui
import win32con
import win32api
from PIL import Image, ImageTk
import tempfile
import ctypes
import sys

class ModernContextMenuManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("右键菜单管理器")
        self.root.geometry("1000x700")
        
        # 存储图标引用，防止被垃圾回收
        self.icons = {}
        
        # 设置主题色
        self.style = ttk.Style()
        self.style.configure("Treeview", rowheight=40)
        self.style.configure("Custom.TButton", padding=10)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="Windows 右键菜单管理器", font=("Microsoft YaHei UI", 16, "bold"))
        title_label.pack(side=tk.LEFT)
        
        # 创建搜索框
        search_frame = ttk.Frame(self.main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_var.trace('w', self.filter_items)
        
        # 创建树形视图
        tree_frame = ttk.Frame(self.main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("图标", "名称", "来源程序", "创建时间", "注册表路径")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        # 设置列标题和宽度
        self.tree.heading("图标", text="图标")
        self.tree.heading("名称", text="名称")
        self.tree.heading("来源程序", text="来源程序")
        self.tree.heading("创建时间", text="创建时间")
        self.tree.heading("注册表路径", text="注册表路径")
        
        self.tree.column("图标", width=50)
        self.tree.column("名称", width=150)
        self.tree.column("来源程序", width=150)
        self.tree.column("创建时间", width=150)
        self.tree.column("注册表路径", width=400)
        
        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # 按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        # 创建按钮
        self.delete_btn = ttk.Button(button_frame, text="删除选中项", 
                                   command=self.delete_selected, style="Custom.TButton")
        self.refresh_btn = ttk.Button(button_frame, text="刷新列表", 
                                    command=self.refresh_menu_items, style="Custom.TButton")
        
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态栏
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var)
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # 创建右键菜单
        self.popup_menu = tk.Menu(self.root, tearoff=0)
        self.popup_menu.add_command(label="打开注册表位置", command=self.open_registry_location)
        self.popup_menu.add_command(label="复制注册表路径", command=self.copy_registry_path)
        self.popup_menu.add_separator()
        self.popup_menu.add_command(label="删除", command=self.delete_selected)
        
        # 绑定右键菜单事件
        self.tree.bind("<Button-3>", self.show_popup_menu)
        
        # 加载菜单项
        self.refresh_menu_items()
        
        # 绑定双击事件
        self.tree.bind("<Double-1>", self.show_item_details)

    def show_popup_menu(self, event):
        """显示右键菜单"""
        # 先选中点击的项
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.popup_menu.post(event.x_root, event.y_root)
            
    def open_registry_location(self):
        """打开注册表位置"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        item = selected_items[0]
        values = self.tree.item(item)['values']
        registry_path = values[4]  # 注册表路径
        
        try:
            # 使用 reg.exe 跳转到指定位置
            full_path = f"HKCR\\{registry_path}"
            os.system(f'start regedit.exe /m')  # 先启动注册表编辑器
            
            # 设置注册表编辑器的最后访问键值
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit") as key:
                winreg.SetValueEx(key, "LastKey", 0, winreg.REG_SZ, full_path)
                
        except Exception as e:
            messagebox.showerror("错误", f"打开注册表位置失败：{str(e)}")
        
    def copy_registry_path(self):
        """复制注册表路径到剪贴板"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        item = selected_items[0]
        values = self.tree.item(item)['values']
        registry_path = values[4]  # 注册表路径
        
        self.root.clipboard_clear()
        self.root.clipboard_append(f"HKEY_CLASSES_ROOT\\{registry_path}")
        messagebox.showinfo("成功", "注册表路径已复制到剪贴板")
        
    def get_registry_creation_time(self, key_path):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_READ)
            info = winreg.QueryInfoKey(key)
            winreg.CloseKey(key)
            timestamp = info[2] / 10000000 - 11644473600  # 转换 Windows 文件时间到 UNIX 时间戳
            return datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
        except Exception as e:
            return "未知"
            
    def get_icon_from_exe(self, exe_path, icon_index=0):
        try:
            if not os.path.exists(exe_path) and not exe_path.lower().endswith('shell32.dll'):
                return None
                
            # 使用文件的唯一标识作为键
            icon_key = f"{exe_path.lower()}_{icon_index}"
            
            # 如果已经有缓存的图标，直接返回
            if icon_key in self.icons:
                return self.icons[icon_key]

            # 提取图标
            if exe_path.lower().endswith('shell32.dll'):
                ico_x = win32gui.ExtractIcon(0, "shell32.dll", icon_index)
            else:
                ico_x = win32gui.ExtractIcon(0, exe_path, icon_index)

            if ico_x:
                try:
                    # 创建DC
                    dc = win32ui.CreateDC()
                    dc.CreateCompatibleDC()
                    
                    # 创建位图
                    bmp = win32ui.CreateBitmap()
                    bmp.CreateCompatibleBitmap(dc, 16, 16)
                    dc.SelectObject(bmp)
                    
                    # 绘制图标
                    win32gui.DrawIconEx(dc.GetHandleOutput(), 0, 0, ico_x, 16, 16, 0, None, win32con.DI_NORMAL)
                    
                    # 获取位图数据
                    bmpinfo = bmp.GetInfo()
                    bmpstr = bmp.GetBitmapBits(True)
                    
                    # 转换为PIL Image
                    img = Image.frombuffer(
                        'RGBA',
                        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                        bmpstr, 'raw', 'BGRA', 0, 1
                    )
                    
                    # 创建PhotoImage
                    photo = ImageTk.PhotoImage(img)
                    self.icons[icon_key] = photo
                    
                    return photo
                    
                finally:
                    # 清理资源
                    if 'dc' in locals():
                        dc.DeleteDC()
                    if 'bmp' in locals():
                        bmp.DeleteObject()
                    win32gui.DestroyIcon(ico_x)
                
        except Exception as e:
            print(f"获取图标失败 {exe_path}: {str(e)}")
        return None
        
    def filter_items(self, *args):
        search_text = self.search_var.get().lower()
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if any(search_text in str(value).lower() for value in values):
                self.tree.reattach(item, '', 'end')
            else:
                self.tree.detach(item)
                
    def show_item_details(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        item = selected_items[0]
        values = self.tree.item(item)['values']
        details = f"""
名称: {values[1]}
来源程序: {values[2]}
创建时间: {values[3]}
注册表路径: {values[4]}
        """
        messagebox.showinfo("详细信息", details)
        
    def refresh_menu_items(self):
        self.tree.delete(*self.tree.get_children())
        self.scan_registry("*\\shell")
        self.scan_registry("Directory\\shell")
        self.status_var.set(f"共找到 {len(self.tree.get_children())} 个右键菜单项")
        
    def scan_registry(self, base_path):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, base_path, 0, winreg.KEY_READ)
            index = 0
            
            while True:
                try:
                    menu_name = winreg.EnumKey(key, index)
                    menu_path = f"{base_path}\\{menu_name}"
                    
                    command = self.get_command(menu_path)
                    creation_time = self.get_registry_creation_time(menu_path)
                    source = self.get_program_source(command)
                    
                    # 获取图标
                    icon = None
                    try:
                        # 先尝试从注册表获取图标
                        icon_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{menu_path}", 0, winreg.KEY_READ)
                        try:
                            icon_path = winreg.QueryValue(icon_key, "Icon")
                            if icon_path:
                                if ',' in icon_path:
                                    path, idx = icon_path.split(',')
                                    icon = self.get_icon_from_exe(path, int(idx))
                                else:
                                    icon = self.get_icon_from_exe(icon_path)
                        except:
                            pass
                        winreg.CloseKey(icon_key)
                    except:
                        pass
                        
                    # 如果没有找到图标，尝试使用程序图标
                    if not icon and source != "未知":
                        icon = self.get_icon_from_exe(source)
                        
                    # 如果还是没有图标，使用默认图标
                    if not icon:
                        try:
                            # 使用默认的shell32图标
                            icon = self.get_icon_from_exe("shell32.dll", 2)  # 使用索引2，这是一个常见的文档图标
                        except:
                            pass
                    
                    values = ("", menu_name, source, creation_time, menu_path)
                    item = self.tree.insert("", tk.END, values=values)
                    
                    if icon:
                        self.tree.set(item, "图标", "")
                        self.tree.item(item, image=icon)
                        
                    index += 1
                except WindowsError:
                    break
                    
            winreg.CloseKey(key)
        except WindowsError as e:
            print(f"扫描注册表失败: {str(e)}")
            
    def get_command(self, menu_path):
        try:
            command_key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{menu_path}\\command", 0, winreg.KEY_READ)
            command = winreg.QueryValue(command_key, "")
            winreg.CloseKey(command_key)
            return command
        except:
            return "未知"
            
    def get_program_source(self, command):
        if command == "未知":
            return "未知"
        try:
            command = command.strip('"')
            exe_path = command.split('"')[0]
            return exe_path
        except:
            return "未知"
            
    def delete_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的项目")
            return
            
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selected_items)} 个右键菜单项吗？"):
            for item in selected_items:
                values = self.tree.item(item)['values']
                registry_path = values[4]
                try:
                    self.delete_registry_key(registry_path)
                    self.tree.delete(item)
                except Exception as e:
                    messagebox.showerror("错误", f"删除失败：{str(e)}")
                    
            self.status_var.set(f"成功删除 {len(selected_items)} 个菜单项")
            
    def delete_registry_key(self, key_path):
        try:
            command_path = f"{key_path}\\command"
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, command_path)
            except:
                pass
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
        except Exception as e:
            raise Exception(f"删除注册表项失败：{str(e)}")
            
    def run(self):
        # 请求管理员权限
        if not ctypes.windll.shell32.IsUserAnAdmin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
            
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernContextMenuManager()
    app.run() 