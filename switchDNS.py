import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import ctypes
import sys
import os
import json
from datetime import datetime

# 检查是否以管理员权限运行
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# 如果不是管理员权限，重新以管理员权限运行
if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

# 主应用类
class NetworkConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网络配置切换工具")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # 设置图标
        self.setup_icon()
        
        # 配置文件路径
        self.config_file = "network_configs.json"
        self.load_configs()
        
        # 创建主界面
        self.create_main_frame()
        
        # 创建系统托盘图标
        self.create_tray_icon()
        
        # 绑定事件
        self.bind_events()
    
    def setup_icon(self):
        # 尝试设置窗口图标
        try:
            # 这里可以添加自定义图标
            pass
        except:
            pass
    
    def load_configs(self):
        # 加载配置文件
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.configs = json.load(f)
        else:
            self.configs = {
                "profiles": []
            }
    
    def save_configs(self):
        # 保存配置文件
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.configs, f, ensure_ascii=False, indent=2)
    
    def create_main_frame(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="网络配置切换工具", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=10)
        
        # 创建配置表单
        form_frame = ttk.LabelFrame(main_frame, text="网络配置", padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 网络接口选择
        interface_frame = ttk.Frame(form_frame)
        interface_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(interface_frame, text="网络接口:").pack(side=tk.LEFT, padx=5)
        self.interface_var = tk.StringVar()
        self.interface_combo = ttk.Combobox(interface_frame, textvariable=self.interface_var, width=30)
        self.interface_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.update_interfaces()
        
        # IP地址
        ip_frame = ttk.Frame(form_frame)
        ip_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(ip_frame, text="IP地址:").pack(side=tk.LEFT, padx=5)
        self.ip_var = tk.StringVar()
        ttk.Entry(ip_frame, textvariable=self.ip_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 子网掩码
        subnet_frame = ttk.Frame(form_frame)
        subnet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(subnet_frame, text="子网掩码:").pack(side=tk.LEFT, padx=5)
        self.subnet_var = tk.StringVar(value="255.255.255.0")
        ttk.Entry(subnet_frame, textvariable=self.subnet_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 网关
        gateway_frame = ttk.Frame(form_frame)
        gateway_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(gateway_frame, text="网关:").pack(side=tk.LEFT, padx=5)
        self.gateway_var = tk.StringVar()
        ttk.Entry(gateway_frame, textvariable=self.gateway_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # DNS服务器
        dns_frame = ttk.Frame(form_frame)
        dns_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dns_frame, text="DNS服务器:").pack(side=tk.LEFT, padx=5)
        self.dns_var = tk.StringVar()
        ttk.Entry(dns_frame, textvariable=self.dns_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Label(dns_frame, text="(多个DNS用逗号分隔)").pack(side=tk.LEFT, padx=5)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.apply_button = ttk.Button(button_frame, text="应用配置", command=self.apply_config)
        self.apply_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.auto_button = ttk.Button(button_frame, text="自动获取配置", command=self.set_auto_config)
        self.auto_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.refresh_button = ttk.Button(button_frame, text="刷新接口", command=self.update_interfaces)
        self.refresh_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 退出按钮
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.on_close)
        self.exit_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    
    def update_interfaces(self):
        # 更新网络接口列表
        try:
            interfaces = []
            
            # 优先使用wmic命令获取网络接口，因为它更稳定
            try:
                # 获取网络适配器的详细信息
                result = subprocess.run(["wmic", "nic", "get", "PNPDeviceID,NetConnectionID"], 
                                      capture_output=True, text=True, encoding='gbk', errors='ignore')
                if result.stdout:
                    lines = result.stdout.split('\n')
                    for line in lines:
                        line = line.strip()
                        if line and 'PNPDeviceID' not in line:
                            # 分割行获取网络连接ID和PNP设备ID
                            # 注意：wmic输出的格式可能因系统而异
                            parts = line.split()
                            if len(parts) >= 2:
                                # 尝试从行中提取网络连接ID（可能包含空格）
                                net_connection_id = ''
                                pnp_device_id = ''
                                for i, part in enumerate(parts):
                                    if '\\' in part:
                                        # 找到PNP设备ID的开始
                                        net_connection_id = ' '.join(parts[:i])
                                        pnp_device_id = ' '.join(parts[i:])
                                        break
                                
                                # 判断是否为物理设备
                                if net_connection_id and pnp_device_id:
                                    # 检查PNPDeviceID是否以PCI开头
                                    if pnp_device_id.startswith('PCI'):
                                        interfaces.append(net_connection_id)
            except Exception as e:
                # 如果wmic失败，尝试使用netsh命令
                try:
                    # 使用netsh命令获取网络接口 - 这应该返回友好名称
                    result = subprocess.run(["netsh", "interface", "ipv4", "show", "interfaces"], 
                                          capture_output=True, text=True, encoding='utf-8', errors='ignore')
                    if result.stdout:
                        # 解析netsh命令输出
                        lines = result.stdout.split('\n')
                        for i, line in enumerate(lines):
                            # 跳过标题行和空行
                            if i == 0 or not line.strip():
                                continue
                            # 尝试从不同位置提取接口名称
                            line = line.strip()
                            if line:
                                # 尝试找到接口名称的起始位置
                                parts = line.split()
                                if len(parts) >= 4:
                                    interface_name = ' '.join(parts[3:])
                                    # 过滤掉GUID格式的名称
                                    if interface_name and not (interface_name.startswith('{') and interface_name.endswith('}')):
                                        interfaces.append(interface_name)
                except Exception as e2:
                    # 如果netsh也失败，尝试使用更简单的netsh命令
                    try:
                        result = subprocess.run(["netsh", "interface", "show", "interface"], 
                                              capture_output=True, text=True, encoding='utf-8', errors='ignore')
                        if result.stdout:
                            for line in result.stdout.split('\n'):
                                line = line.strip()
                                if line and 'Name' in line:
                                    # 尝试提取名称
                                    name_part = line.split('Name')[-1].strip()
                                    if name_part and not (name_part.startswith('{') and name_part.endswith('}')):
                                        interfaces.append(name_part)
                    except:
                        pass
            
            # 去重并过滤空值
            interfaces = list(filter(None, set(interfaces)))
            interfaces = [iface for iface in interfaces if iface and not (iface.startswith('{') and iface.endswith('}'))]
            
            # 如果仍然没有接口，添加一些默认值
            if not interfaces:
                interfaces = ["以太网", "WLAN", "本地连接", "以太网 2", "WLAN 2"]
            
            self.interface_combo['values'] = interfaces
            if interfaces:
                self.interface_var.set(interfaces[0])
            else:
                # 如果仍然没有接口，添加一个默认值
                self.interface_combo['values'] = ["请选择网络接口"]
                self.interface_var.set("请选择网络接口")
        except Exception as e:
            messagebox.showerror("错误", f"获取网络接口失败: {e}")
            # 发生错误时添加默认值
            self.interface_combo['values'] = ["以太网", "WLAN", "本地连接"]
            self.interface_var.set("以太网")
    
    def apply_config(self):
        # 应用网络配置
        interface = self.interface_var.get()
        ip = self.ip_var.get()
        subnet = self.subnet_var.get()
        gateway = self.gateway_var.get()
        dns = self.dns_var.get()
        
        if not interface:
            messagebox.showerror("错误", "请选择网络接口")
            return
        
        if not ip:
            messagebox.showerror("错误", "请输入IP地址")
            return
        
        try:
            # 设置IP地址和子网掩码
            subprocess.run(["netsh", "interface", "ipv4", "set", "address", 
                          "name=" + interface, "static", ip, subnet, gateway, "1"], 
                         check=True, capture_output=True)
            
            # 设置DNS服务器
            if dns:
                dns_servers = dns.split(',')
                primary_dns = dns_servers[0].strip()
                # 设置主DNS
                subprocess.run(["netsh", "interface", "ipv4", "set", "dns", 
                              "name=" + interface, "static", primary_dns], 
                             check=True, capture_output=True)
                
                # 设置备用DNS
                for i, dns_server in enumerate(dns_servers[1:], 2):
                    subprocess.run(["netsh", "interface", "ipv4", "add", "dns", 
                                  "name=" + interface, dns_server.strip(), str(i)], 
                                 check=True, capture_output=True)
            
            messagebox.showinfo("成功", "网络配置已应用")
            
            # 保存配置到历史记录
            self.save_config_to_history(interface, ip, subnet, gateway, dns)
            
        except Exception as e:
            messagebox.showerror("错误", f"应用配置失败: {e}")
    
    def set_auto_config(self):
        # 设置自动获取配置
        interface = self.interface_var.get()
        
        if not interface:
            messagebox.showerror("错误", "请选择网络接口")
            return
        
        try:
            # 设置自动获取IP
            subprocess.run(["netsh", "interface", "ipv4", "set", "address", 
                          "name=" + interface, "source=dhcp"], 
                         check=True, capture_output=True)
            
            # 设置自动获取DNS
            subprocess.run(["netsh", "interface", "ipv4", "set", "dns", 
                          "name=" + interface, "source=dhcp"], 
                         check=True, capture_output=True)
            
            messagebox.showinfo("成功", "已设置为自动获取配置")
            
        except Exception as e:
            messagebox.showerror("错误", f"设置自动配置失败: {e}")
    
    def save_config_to_history(self, interface, ip, subnet, gateway, dns):
        # 保存配置到历史记录
        config = {
            "name": f"配置 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "interface": interface,
            "ip": ip,
            "subnet": subnet,
            "gateway": gateway,
            "dns": dns,
            "timestamp": datetime.now().isoformat()
        }
        
        # 保持历史记录不超过10条
        self.configs["profiles"].insert(0, config)
        if len(self.configs["profiles"]) > 10:
            self.configs["profiles"] = self.configs["profiles"][:10]
        
        self.save_configs()
    
    def create_tray_icon(self):
        # 创建系统托盘图标
        try:
            import win32api
            import win32gui
            import win32con
            
            # 定义托盘图标类
            class TrayIcon:
                def __init__(self, app):
                    self.app = app
                    self.hwnd = None
                    self.icon = None
                    self.nid = None
                    # 尝试初始化系统托盘
                    try:
                        self.create_tray_icon()
                    except Exception as e:
                        # 只显示一次警告，避免干扰用户
                        if not hasattr(self.app, 'tray_warning_shown'):
                            messagebox.showwarning("警告", f"系统托盘功能暂时不可用: {e}")
                            self.app.tray_warning_shown = True
                
                def create_tray_icon(self):
                    # 注册窗口类
                    class_name = self.register_window_class()
                    if not class_name:
                        return
                    
                    # 创建窗口
                    try:
                        self.hwnd = win32gui.CreateWindow(
                            class_name,
                            "网络配置切换工具",
                            win32con.WS_OVERLAPPED | win32con.WS_SYSMENU,
                            win32con.CW_USEDEFAULT,
                            win32con.CW_USEDEFAULT,
                            win32con.CW_USEDEFAULT,
                            win32con.CW_USEDEFAULT,
                            0,
                            0,
                            win32api.GetModuleHandle(None),
                            None
                        )
                    except Exception as e:
                        if not hasattr(self.app, 'tray_warning_shown'):
                            messagebox.showwarning("警告", f"创建窗口失败: {e}")
                            self.app.tray_warning_shown = True
                        return
                    
                    # 创建托盘图标
                    try:
                        # 使用更兼容的方式创建NOTIFYICONDATA
                        # 避免使用字典形式，使用正确的参数格式
                        icon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
                        
                        # 尝试使用更简单的方式添加托盘图标
                        # 注意：这里可能需要根据pywin32版本调整
                        try:
                            # 方法1：使用正确的NOTIFYICONDATA结构
                            nid = (self.hwnd, 100, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP, 
                                   win32con.WM_USER + 100, icon, "网络配置切换工具")
                            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
                            self.nid = nid
                        except Exception as e1:
                            try:
                                # 方法2：使用更简单的参数格式
                                win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, {
                                    'hWnd': self.hwnd,
                                    'uID': 100,
                                    'uFlags': win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                                    'uCallbackMessage': win32con.WM_USER + 100,
                                    'hIcon': icon,
                                    'szTip': "网络配置切换工具"
                                })
                            except Exception as e2:
                                raise Exception(f"托盘图标创建失败: {e1}, {e2}")
                        
                        # 开始消息循环
                        import threading
                        self.message_thread = threading.Thread(target=self.pump_messages)
                        self.message_thread.daemon = True
                        self.message_thread.start()
                    except Exception as e:
                        if not hasattr(self.app, 'tray_warning_shown'):
                            messagebox.showwarning("警告", f"系统托盘图标创建失败: {e}")
                            self.app.tray_warning_shown = True
                
                def register_window_class(self):
                    # 注册窗口类
                    try:
                        # 使用正确的WNDCLASS结构
                        wc = win32gui.WNDCLASS()
                        wc.lpfnWndProc = self.window_proc
                        wc.hInstance = win32api.GetModuleHandle(None)
                        wc.lpszClassName = "NetworkConfigTray"
                        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
                        wc.hbrBackground = win32con.COLOR_WINDOW + 1
                        return win32gui.RegisterClass(wc)
                    except Exception as e:
                        if not hasattr(self.app, 'tray_warning_shown'):
                            messagebox.showwarning("警告", f"窗口类注册失败: {e}")
                            self.app.tray_warning_shown = True
                        return None
                
                def window_proc(self, hwnd, msg, wparam, lparam):
                    # 窗口消息处理
                    if msg == win32con.WM_USER + 100:
                        if lparam == win32con.WM_RBUTTONUP:
                            # 右键点击系统托盘图标
                            self.show_context_menu()
                        elif lparam == win32con.WM_LBUTTONDBLCLK:
                            # 双击系统托盘图标
                            self.app.root.deiconify()
                    elif msg == win32con.WM_DESTROY:
                        try:
                            if hasattr(self, 'nid') and self.nid:
                                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.nid)
                        except:
                            pass
                        win32gui.PostQuitMessage(0)
                    return 0
                
                def show_context_menu(self):
                    # 显示上下文菜单
                    try:
                        # 创建菜单
                        menu = win32gui.CreatePopupMenu()
                        
                        # 添加菜单项 - 确保所有字符串都是有效的
                        try:
                            win32gui.AppendMenu(menu, win32con.MF_STRING, 1, "显示主窗口")
                            win32gui.AppendMenu(menu, win32con.MF_STRING, 2, "恢复自动获取")
                            # 添加分隔线 - 使用正确的参数
                            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
                        except Exception as e:
                            messagebox.showwarning("警告", f"添加菜单项失败: {e}")
                            return
                        
                        # 添加快速配置选项
                        try:
                            if self.app.configs.get("profiles"):
                                win32gui.AppendMenu(menu, win32con.MF_STRING, 3, "快速配置")
                                submenu = win32gui.CreatePopupMenu()
                                for i, profile in enumerate(self.app.configs["profiles"][:5]):
                                    # 确保profile["name"]是有效的字符串
                                    profile_name = profile.get("name", f"配置{i+1}")
                                    if profile_name is None:
                                        profile_name = f"配置{i+1}"
                                    # 确保profile_name是字符串类型
                                    profile_name = str(profile_name)
                                    win32gui.AppendMenu(submenu, win32con.MF_STRING, 10 + i, profile_name)
                                # 确保submenu是有效的
                                if submenu:
                                    win32gui.AppendMenu(menu, win32con.MF_POPUP, submenu, "快速配置")
                        except Exception as e:
                            messagebox.showwarning("警告", f"添加快速配置菜单失败: {e}")
                        
                        try:
                            # 添加分隔线 - 使用正确的参数
                            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
                            win32gui.AppendMenu(menu, win32con.MF_STRING, 99, "退出")
                        except Exception as e:
                            messagebox.showwarning("警告", f"添加退出菜单项失败: {e}")
                            return
                        
                        # 获取鼠标位置
                        pos = win32gui.GetCursorPos()
                        
                        # 显示菜单并获取用户选择
                        try:
                            win32gui.SetForegroundWindow(self.hwnd)
                            cmd = win32gui.TrackPopupMenu(
                                menu,
                                win32con.TPM_RIGHTALIGN | win32con.TPM_BOTTOMALIGN | win32con.TPM_RETURNCMD,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None
                            )
                        except Exception as e:
                            messagebox.showwarning("警告", f"显示菜单失败: {e}")
                            return
                        
                        # 处理菜单命令
                        if cmd == 1:
                            # 显示主窗口
                            self.app.root.deiconify()
                        elif cmd == 2:
                            # 恢复自动获取
                            self.app.set_auto_config()
                        elif cmd == 99:
                            # 退出程序
                            self.app.on_close()
                        elif cmd >= 10:
                            # 快速配置选项
                            profile_index = cmd - 10
                            if profile_index < len(self.app.configs.get("profiles", [])):
                                profile = self.app.configs["profiles"][profile_index]
                                # 应用快速配置
                                self.app.interface_var.set(profile.get("interface", ""))
                                self.app.ip_var.set(profile.get("ip", ""))
                                self.app.subnet_var.set(profile.get("subnet", ""))
                                self.app.gateway_var.set(profile.get("gateway", ""))
                                self.app.dns_var.set(profile.get("dns", ""))
                                self.app.apply_config()
                                # 显示主窗口
                                self.app.root.deiconify()
                    except Exception as e:
                        # 显示错误信息以便调试
                        messagebox.showwarning("警告", f"显示上下文菜单失败: {e}")
                
                def get_default_icon(self):
                    # 获取默认图标 - 使用应用程序默认图标
                    try:
                        # 使用默认应用程序图标
                        return win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
                    except:
                        # 如果失败，返回None
                        return None
                
                def pump_messages(self):
                    # 消息循环
                    try:
                        win32gui.PumpMessages()
                    except:
                        pass
            
            # 尝试创建系统托盘图标
            try:
                self.tray_icon = TrayIcon(self)
            except Exception as e:
                if not hasattr(self, 'tray_warning_shown'):
                    messagebox.showwarning("警告", f"系统托盘功能初始化失败: {e}")
                    self.tray_warning_shown = True
        except ImportError as e:
            if not hasattr(self, 'tray_warning_shown'):
                messagebox.showwarning("警告", f"系统托盘功能需要pywin32库: {e}")
                self.tray_warning_shown = True
        except Exception as e:
            if not hasattr(self, 'tray_warning_shown'):
                messagebox.showwarning("警告", f"系统托盘功能初始化失败: {e}")
                self.tray_warning_shown = True
    
    def bind_events(self):
        # 绑定窗口事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        # 绑定最小化事件
        self.root.bind("<Unmap>", self.on_minimize)
    
    def on_close(self):
        # 窗口关闭事件 - 完全退出程序
        if hasattr(self, 'tray_icon'):
            try:
                # 尝试清理系统托盘图标
                if hasattr(self.tray_icon, 'nid') and self.tray_icon.nid:
                    import win32gui
                    win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.tray_icon.nid)
            except:
                pass
        self.root.destroy()
        sys.exit()
    
    def on_minimize(self, event=None):
        # 窗口最小化事件 - 最小化到托盘
        try:
            # 检查是否是最小化事件
            if event and event.widget == self.root and self.root.state() == 'iconic':
                # 隐藏主窗口
                self.root.withdraw()
                # 尝试显示系统托盘图标
                # 不再显示消息提示，直接最小化
        except Exception as e:
            messagebox.showwarning("警告", f"最小化到托盘失败: {e}")
            # 如果最小化到托盘失败，恢复显示窗口
            if self.root.state() == 'withdrawn':
                self.root.deiconify()
    
    def _is_physical_adapter(self, adapter_name):
        """通过注册表查询判断是否为物理网卡"""
        try:
            import winreg
            
            # 注册表路径
            reg_path = r'SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}'
            adapter_key = adapter_name.strip('\\')
            
            # 打开注册表
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as network_key:
                # 打开适配器子键
                with winreg.OpenKey(network_key, f'{adapter_key}\\Connection') as connection_key:
                    # 读取PnpInstanceID值
                    pnp_instance_id, _ = winreg.QueryValueEx(connection_key, 'PnpInstanceID')
                    
                    # 检查PnpInstanceID是否以PCI开头
                    if pnp_instance_id.startswith('PCI'):
                        # 读取MediaSubType值
                        media_sub_type, _ = winreg.QueryValueEx(connection_key, 'MediaSubType')
                        # MediaSubType为01是有线网卡，02是无线网卡
                        if media_sub_type in (1, 2):
                            return True
        except Exception:
            # 注册表查询失败，返回False
            pass
        return False
    
    def show_main_window(self):
        # 显示主窗口
        self.root.deiconify()
        self.root.lift()

# 创建主窗口
root = tk.Tk()
app = NetworkConfigApp(root)

# 启动应用
root.mainloop()