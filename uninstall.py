import winreg
import subprocess
import os
import shutil
import psutil
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from typing import List, Dict
import threading
import time
import ctypes  # 添加ctypes导入，用于管理员权限检查

def get_installed_software() -> List[Dict]:
    """获取已安装的软件列表"""
    software_list = []
    paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    for path in paths:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ)
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        software = {}
                        try:
                            software['name'] = winreg.QueryValueEx(subkey, 'DisplayName')[0]
                            software['uninstall_string'] = winreg.QueryValueEx(subkey, 'UninstallString')[0]
                            
                            # 这些值可能不存在于所有软件中
                            try:
                                software['version'] = winreg.QueryValueEx(subkey, 'DisplayVersion')[0]
                            except:
                                software['version'] = "未知"
                                
                            try:
                                software['publisher'] = winreg.QueryValueEx(subkey, 'Publisher')[0]
                            except:
                                software['publisher'] = "未知"
                                
                            try:
                                software['install_location'] = winreg.QueryValueEx(subkey, 'InstallLocation')[0]
                            except:
                                software['install_location'] = ""
                                
                            # 只添加有名称和卸载字符串的软件
                            if software['name'] and software['uninstall_string']:
                                software_list.append(software)
                        except:
                            continue
                    finally:
                        winreg.CloseKey(subkey)
                except WindowsError:
                    continue
        except WindowsError:
            continue
        finally:
            try:
                winreg.CloseKey(key)
            except:
                pass
    
    return software_list

def force_kill_process(process_name: str, log_func=print):
    """强制结束指定进程"""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'].lower() == process_name.lower():
                log_func(f"正在结束进程: {proc.info['name']}")
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

def uninstall_software(software: Dict, log_func=print) -> bool:
    """卸载指定软件"""
    try:
        log_func(f"执行卸载命令: {software['uninstall_string']}")
        # 执行卸载命令
        subprocess.run(software['uninstall_string'], shell=True, check=True)
        return True
    except subprocess.CalledProcessError as e:
        log_func(f"卸载命令执行失败: {e}")
        return False

def is_admin():
    """检查程序是否以管理员权限运行"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def force_delete_file(path):
    """使用系统命令强制删除文件"""
    try:
        if os.path.exists(path):
            if os.path.isfile(path):
                os.chmod(path, 0o777)  # 尝试更改文件权限
                os.remove(path)
            else:
                subprocess.run(f'rmdir /s /q "{path}"', shell=True, check=False)
        return True
    except Exception as e:
        return False

def clean_residual(software: Dict, log_func=print):
    """清理软件残留"""
    possible_names = [software['name']]
    # 添加没有版本号的名称变体
    if '(' in software['name']:
        possible_names.append(software['name'].split('(')[0].strip())
    
    paths_to_check = []
    
    # 添加安装位置
    if software['install_location'] and os.path.exists(software['install_location']):
        paths_to_check.append(software['install_location'])
    
    # 添加常见的安装路径
    for name in possible_names:
        paths_to_check.extend([
            os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), name),
            os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), name),
            os.path.join(os.environ.get('APPDATA', ''), name),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), name)
        ])
    
    # 检查管理员权限
    admin_status = "是" if is_admin() else "否"
    log_func(f"当前程序管理员权限状态: {admin_status}")
    
    for path in paths_to_check:
        if os.path.exists(path):
            try:
                log_func(f"正在删除残留文件夹: {path}")
                try:
                    # 首先尝试使用标准方法删除
                    shutil.rmtree(path)
                    log_func(f"成功删除: {path}")
                except (PermissionError, OSError) as e:
                    log_func(f"标准删除失败: {e}，尝试强制删除...")
                    # 如果标准方法失败，尝试强制删除
                    if force_delete_file(path):
                        log_func(f"强制删除成功: {path}")
                    else:
                        log_func(f"无法删除路径: {path}，请尝试以管理员身份运行程序")
            except Exception as e:
                log_func(f"删除过程中发生错误: {e}")
                log_func(f"无法删除路径: {path}, 错误: {e}")

class UninstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("软件卸载工具")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        self.software_list = []
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建顶部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="软件卸载工具", font=("Arial", 16)).pack(side=tk.LEFT)
        
        self.refresh_btn = ttk.Button(control_frame, text="刷新软件列表", command=self.refresh_software_list)
        self.refresh_btn.pack(side=tk.RIGHT)
        
        # 创建软件列表区域
        list_frame = ttk.LabelFrame(main_frame, text="已安装的软件")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建表格
        columns = ("name", "version", "publisher")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # 定义表头
        self.tree.heading("name", text="软件名称")
        self.tree.heading("version", text="版本")
        self.tree.heading("publisher", text="发布者")
        
        # 设置列宽
        self.tree.column("name", width=300)
        self.tree.column("version", width=100)
        self.tree.column("publisher", width=200)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        # 放置表格和滚动条
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_item_double_click)
        
        # 创建操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.uninstall_btn = ttk.Button(button_frame, text="卸载选中软件", command=self.uninstall_selected)
        self.uninstall_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 创建日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    def log(self, message):
        """向日志区域添加消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    def refresh_software_list(self):
        """刷新软件列表"""
        self.refresh_btn.config(state=tk.DISABLED)
        self.log("正在获取已安装软件列表...")
        
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 在后台线程中获取软件列表
        def fetch_software():
            self.software_list = get_installed_software()
            self.software_list.sort(key=lambda x: x['name'])
            
            # 更新UI（必须在主线程中进行）
            self.root.after(0, self.update_software_list)
        
        threading.Thread(target=fetch_software, daemon=True).start()
    def update_software_list(self):
        """更新软件列表UI"""
        for software in self.software_list:
            self.tree.insert("", tk.END, values=(
                software['name'],
                software['version'],
                software['publisher']
            ))
        
        self.log(f"已加载 {len(self.software_list)} 个软件")
        self.refresh_btn.config(state=tk.NORMAL)
    def get_selected_software(self):
        """获取选中的软件"""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showinfo("提示", "请先选择要卸载的软件")
            return None
        
        item_index = self.tree.index(selected_item[0])
        return self.software_list[item_index]
    def on_item_double_click(self, event):
        """双击软件项时的处理"""
        self.uninstall_selected()
    # 将ProgressWindow类移到UninstallerApp类外部
    class ProgressWindow:
        """卸载进度窗口"""
        def __init__(self, parent, software_name):
            self.window = tk.Toplevel(parent)
            self.window.title("卸载进度")
            self.window.geometry("400x150")
            self.window.resizable(False, False)
            self.window.transient(parent)  # 设置为父窗口的临时窗口
            self.window.grab_set()  # 模态窗口
            
            # 窗口居中
            self.window.update_idletasks()
            width = self.window.winfo_width()
            height = self.window.winfo_height()
            x = (self.window.winfo_screenwidth() // 2) - (width // 2)
            y = (self.window.winfo_screenheight() // 2) - (height // 2)
            self.window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            
            # 创建界面元素
            self.frame = ttk.Frame(self.window, padding="20")
            self.frame.pack(fill=tk.BOTH, expand=True)
            
            self.status_label = ttk.Label(self.frame, text=f"正在卸载 {software_name}...", font=("Arial", 12))
            self.status_label.pack(pady=(0, 20))
            
            self.progress = ttk.Progressbar(self.frame, orient=tk.HORIZONTAL, length=360, mode='determinate')
            self.progress.pack(pady=(0, 20))
            
            self.detail_label = ttk.Label(self.frame, text="准备中...", font=("Arial", 10))
            self.detail_label.pack()
            
            self.progress_value = 0
            self.is_complete = False
            self.logs = []
            
        def update_progress(self, value, detail=None):
            """更新进度条"""
            if detail:
                self.detail_label.config(text=detail)
                self.logs.append(detail)
            
            # 平滑更新进度条
            current = self.progress_value
            target = value
            step = (target - current) / 10
            
            for i in range(10):
                current += step
                self.progress_value = current
                self.progress['value'] = current
                self.window.update()
                time.sleep(0.02)
        def complete(self):
            """完成卸载"""
            self.is_complete = True
            self.status_label.config(text="卸载完成！")
            self.detail_label.config(text="所有操作已完成")
            self.progress['value'] = 100
            self.window.update()
            
            # 2秒后自动关闭
            self.window.after(2000, self.window.destroy)
        def get_logs(self):
            """获取所有日志"""
            return self.logs

class UninstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("软件卸载工具")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        self.software_list = []
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建顶部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="软件卸载工具", font=("Arial", 16)).pack(side=tk.LEFT)
        
        self.refresh_btn = ttk.Button(control_frame, text="刷新软件列表", command=self.refresh_software_list)
        self.refresh_btn.pack(side=tk.RIGHT)
        
        # 创建软件列表区域
        list_frame = ttk.LabelFrame(main_frame, text="已安装的软件")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建表格
        columns = ("name", "version", "publisher")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # 定义表头
        self.tree.heading("name", text="软件名称")
        self.tree.heading("version", text="版本")
        self.tree.heading("publisher", text="发布者")
        
        # 设置列宽
        self.tree.column("name", width=300)
        self.tree.column("version", width=100)
        self.tree.column("publisher", width=200)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        # 放置表格和滚动条
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_item_double_click)
        
        # 创建操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.uninstall_btn = ttk.Button(button_frame, text="卸载选中软件", command=self.uninstall_selected)
        self.uninstall_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 创建日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    
    def log(self, message):
        """向日志区域添加消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    
    def refresh_software_list(self):
        """刷新软件列表"""
        self.refresh_btn.config(state=tk.DISABLED)
        self.log("正在获取已安装软件列表...")
        
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 在后台线程中获取软件列表
        def fetch_software():
            self.software_list = get_installed_software()
            self.software_list.sort(key=lambda x: x['name'])
            
            # 更新UI（必须在主线程中进行）
            self.root.after(0, self.update_software_list)
        
        threading.Thread(target=fetch_software, daemon=True).start()
    
    def update_software_list(self):
        """更新软件列表UI"""
        for software in self.software_list:
            self.tree.insert("", tk.END, values=(
                software['name'],
                software['version'],
                software['publisher']
            ))
        
        self.log(f"已加载 {len(self.software_list)} 个软件")
        self.refresh_btn.config(state=tk.NORMAL)
    
    def get_selected_software(self):
        """获取选中的软件"""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showinfo("提示", "请先选择要卸载的软件")
            return None
        
        item_index = self.tree.index(selected_item[0])
        return self.software_list[item_index]
    
    def on_item_double_click(self, event):
        """双击软件项时的处理"""
        self.uninstall_selected()
    
    def uninstall_selected(self):
        """卸载选中的软件"""
        software = self.get_selected_software()
        if not software:
            return
        
        # 确认卸载
        if not messagebox.askyesno("确认", f"确定要卸载 {software['name']} 吗？"):
            return
        
        # 禁用按钮，防止重复操作
        self.uninstall_btn.config(state=tk.DISABLED)
        self.refresh_btn.config(state=tk.DISABLED)
        
        # 创建进度窗口
        progress_window = ProgressWindow(self.root, software['name'])
        
        # 在后台线程中执行卸载
        def do_uninstall():
            # 记录日志的函数
            def log_to_progress(message):
                self.log(message)  # 仍然记录到主界面日志
                self.root.after(0, lambda: progress_window.update_progress(
                    progress_window.progress_value, message))
            
            # 更新进度到10%
            self.root.after(0, lambda: progress_window.update_progress(10, f"正在卸载 {software['name']}..."))
            
            # 尝试结束相关进程
            process_name = software['name'].split()[0] + ".exe"
            log_to_progress(f"尝试结束相关进程: {process_name}")
            force_kill_process(process_name, log_to_progress)
            
            # 更新进度到30%
            self.root.after(0, lambda: progress_window.update_progress(30, "正在执行卸载命令..."))
            
            # 执行卸载
            uninstall_success = uninstall_software(software, log_to_progress)
            
            # 更新进度到70%
            self.root.after(0, lambda: progress_window.update_progress(
                70, "卸载命令执行" + ("成功" if uninstall_success else "失败")))
            
            if uninstall_success:
                # 清理残留
                self.root.after(0, lambda: progress_window.update_progress(80, "正在清理残留文件..."))
                clean_residual(software, log_to_progress)
                self.root.after(0, lambda: progress_window.update_progress(95, "清理完成"))
                
                # 刷新软件列表
                self.root.after(1000, self.refresh_software_list)
            
            # 完成
            self.root.after(0, progress_window.complete)
            
            # 恢复按钮状态
            self.root.after(0, lambda: self.uninstall_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.refresh_btn.config(state=tk.NORMAL))
        
        threading.Thread(target=do_uninstall, daemon=True).start()

def main():
    # 检查管理员权限
    if not is_admin():
        print("警告: 程序未以管理员权限运行，某些卸载和清理操作可能失败")
        print("建议: 右键点击程序，选择'以管理员身份运行'")
        
        # 可选: 自动请求管理员权限重启
        if messagebox.askyesno("权限提示", "程序需要管理员权限才能完全卸载软件。\n是否以管理员身份重新启动程序？"):
            try:
                import sys
                ctypes.windll.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
                sys.exit(0)
            except:
                pass
    
    root = tk.Tk()
    app = UninstallerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()