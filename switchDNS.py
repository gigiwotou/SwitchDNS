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
        
        # 配置文件路径 - 使用执行文件同目录
        import os
        import sys
        
        try:
            # 获取执行文件路径
            if getattr(sys, 'frozen', False):
                # 编译为exe的情况
                exe_path = sys.executable
                config_dir = os.path.dirname(exe_path)
            else:
                # 直接运行Python脚本的情况
                script_path = os.path.abspath(__file__)
                config_dir = os.path.dirname(script_path)
            
            # 配置文件路径
            self.config_file = os.path.join(config_dir, "network_configs.json")
        except:
            # 如果失败，使用当前目录作为后备
            config_dir = os.getcwd()
            self.config_file = os.path.join(config_dir, "network_configs.json")
        
        # 确保配置文件存在
        if not os.path.exists(self.config_file):
            # 创建默认配置文件
            default_config = {
                "profiles": [
                    {
                        "name": "默认配置",
                        "interface": "以太网",
                        "ip": "192.168.1.100",
                        "subnet": "255.255.255.0",
                        "gateway": "192.168.1.1",
                        "dns": "8.8.8.8,114.114.114.114"
                    }
                ]
            }
            try:
                # 确保目录存在
                os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
                # 写入默认配置
                import json
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, ensure_ascii=False, indent=2)
            except:
                pass
        
        # 加载配置文件
        self.load_configs()
        
        # 创建主界面
        self.create_main_frame()
        
        # 应用配置文件中的值到UI界面
        self.apply_config_from_file()
        
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
        # 默认配置
        default_config = {
            "profiles": [
                {
                    "name": "默认配置",
                    "interface": "以太网",
                    "ip": "192.168.1.100",
                    "subnet": "255.255.255.0",
                    "gateway": "192.168.1.1",
                    "dns": "8.8.8.8,114.114.114.114"
                }
            ]
        }
        
        try:
            import os
            import json
            
            print(f"开始加载配置文件: {self.config_file}")
            
            if os.path.exists(self.config_file):
                print("配置文件存在，开始读取...")
                try:
                    with open(self.config_file, 'r', encoding='utf-8') as f:
                        self.configs = json.load(f)
                        print(f"配置文件内容: {self.configs}")
                        
                        # 确保配置格式正确
                        if "profiles" not in self.configs:
                            self.configs["profiles"] = []
                            print("配置文件缺少profiles键，添加默认值")
                except Exception as e:
                    print(f"读取配置文件错误: {e}")
                    # 配置文件格式错误，使用默认配置
                    self.configs = default_config
            else:
                print("配置文件不存在，使用默认配置")
                # 配置文件不存在，使用默认配置
                self.configs = default_config
                
                # 创建默认配置文件
                try:
                    # 确保目录存在
                    os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
                    with open(self.config_file, 'w', encoding='utf-8') as f:
                        json.dump(self.configs, f, ensure_ascii=False, indent=2)
                    print("已创建默认配置文件")
                except Exception as e:
                    print(f"创建默认配置文件错误: {e}")
        except Exception as e:
            print(f"加载配置文件发生错误: {e}")
            # 任何错误都使用默认配置
            self.configs = default_config
        print(f"最终配置: {self.configs}")
    
    def save_configs(self):
        # 保存配置文件
        try:
            import os
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                import json
                json.dump(self.configs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # 保存失败，忽略错误
            pass
    
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
        # 不在这里调用update_interfaces()，而是在apply_config_from_file中调用
        
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
        
        # 检查IP和网关是否相同
        if gateway and ip == gateway:
            messagebox.showerror("错误", "IP地址和网关不能相同，请修改配置")
            return
        
        try:
            # 打印调试信息
            print(f"应用配置: 接口={interface}, IP={ip}, 子网={subnet}, 网关={gateway}, DNS={dns}")
            
            # 设置IP地址、子网掩码和网关
            if gateway:
                # 如果有网关，使用完整命令
                print(f"使用带网关的命令: netsh interface ipv4 set address name={interface} static {ip} {subnet} {gateway} 1")
                subprocess.run(["netsh", "interface", "ipv4", "set", "address", 
                              "name=" + interface, "static", ip, subnet, gateway, "1"], 
                             check=True, capture_output=True)
            else:
                # 如果没有网关，使用不带网关的命令
                print(f"使用不带网关的命令: netsh interface ipv4 set address name={interface} static {ip} {subnet}")
                subprocess.run(["netsh", "interface", "ipv4", "set", "address", 
                              "name=" + interface, "static", ip, subnet], 
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
    
    def get_actual_interfaces(self):
        # 获取系统中实际的网络接口列表
        interfaces = []
        try:
            # 使用netsh命令获取所有网络接口
            result = subprocess.run(["netsh", "interface", "show", "interface"], 
                                  capture_output=True, text=True, encoding='gbk', errors='ignore')
            if result.stdout:
                lines = result.stdout.split('\n')
                for line in lines:
                    line = line.strip()
                    if line and 'Name' not in line and '----' not in line:
                        # 提取接口名称（从状态和类型后的部分）
                        parts = line.split()
                        if len(parts) >= 4:
                            # 接口名称可能包含空格，从第4个部分开始
                            interface_name = ' '.join(parts[3:])
                            if interface_name and not (interface_name.startswith('{') and interface_name.endswith('}')):
                                interfaces.append(interface_name)
        except Exception as e:
            print(f"获取接口列表失败: {e}")
        return interfaces
    
    def set_auto_config(self):
        # 设置自动获取配置
        interface = self.interface_var.get()
        
        # 尝试获取实际的网络接口列表
        actual_interfaces = self.get_actual_interfaces()
        print(f"实际接口列表: {actual_interfaces}")
        
        # 如果没有选择接口，使用第一个实际接口
        if not interface and actual_interfaces:
            interface = actual_interfaces[0]
            print(f"使用第一个实际接口: {interface}")
        
        if not interface:
            messagebox.showerror("错误", "请先在主窗口中选择网络接口")
            return
        
        try:
            # 打印调试信息
            print(f"尝试设置自动配置: 接口={interface}")
            
            # 尝试使用管理员权限运行命令
            success = False
            error_message = ""
            
            # 尝试使用powershell命令，可能具有更高的权限
            try:
                # 使用powershell命令设置自动获取配置
                ps_cmd = f'netsh interface ipv4 set address name="{interface}" source=dhcp'
                print(f"尝试PowerShell命令: {ps_cmd}")
                # 使用powershell执行命令
                subprocess.run(["powershell", "-Command", ps_cmd], check=True, capture_output=True)
                success = True
            except Exception as e4:
                error_message += f"PowerShell命令失败: {e4}\n"
                print(error_message)
            
            # 如果PowerShell失败，尝试使用runas提升权限
            if not success:
                try:
                    # 尝试使用runas提升权限
                    runas_cmd = f'netsh interface ipv4 set address name="{interface}" source=dhcp'
                    print(f"尝试runas命令: {runas_cmd}")
                    # 使用ShellExecuteW以管理员权限运行
                    import ctypes
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", f"/c {runas_cmd}", None, 1)
                    success = True
                    # 延迟一下，让命令有时间执行
                    import time
                    time.sleep(2)
                except Exception as e5:
                    error_message += f"runas命令失败: {e5}\n"
                    print(error_message)
            
            # 如果IP设置成功，设置DNS
            if success:
                try:
                    # 设置自动获取DNS
                    dns_cmd = f'netsh interface ipv4 set dns name="{interface}" source=dhcp'
                    print(f"执行DNS命令: {dns_cmd}")
                    subprocess.run(["powershell", "-Command", dns_cmd], check=True, capture_output=True)
                    
                    messagebox.showinfo("成功", f"已为接口 '{interface}' 设置为自动获取配置")
                except Exception as dns_error:
                    print(f"DNS设置失败: {dns_error}")
                    messagebox.showinfo("部分成功", f"IP设置成功，但DNS设置失败: {dns_error}\n\n接口: {interface}")
            else:
                # 所有方法都失败，显示详细错误信息
                error_msg = f"设置自动配置失败: {error_message}\n\n接口: {interface}\n\n实际可用接口: {actual_interfaces}\n\n请尝试在主窗口中选择正确的网络接口后重试。"
                print(error_msg)
                messagebox.showerror("错误", error_msg)
            
        except Exception as e:
            # 提供更详细的错误信息
            error_msg = f"设置自动配置失败: {e}\n\n接口: {interface}\n\n实际可用接口: {actual_interfaces}"
            print(error_msg)
            messagebox.showerror("错误", error_msg)
    
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
        
        # 保存配置文件
        self.save_configs()
        print(f"配置已保存: {config}")
    
    def create_tray_icon(self):
        # 创建系统托盘图标
        try:
            # 尝试使用pystray库（如果可用）
            try:
                from pystray import Icon, Menu, MenuItem
                from PIL import Image, ImageDraw
                
                # 创建一个简单的图标
                def create_image():
                    width = 64
                    height = 64
                    image = Image.new('RGB', (width, height), color='white')
                    draw = ImageDraw.Draw(image)
                    draw.rectangle([16, 16, 48, 48], fill='black')
                    return image
                
                # 创建菜单
                menu = Menu(
                    MenuItem('Show Main Window', self.root.deiconify),
                    MenuItem('Auto DHCP', self.set_auto_config),
                    MenuItem('Exit', self.on_close)
                )
                
                # 创建系统托盘图标
                icon = Icon('NetworkConfigTool', create_image(), 'Network Config Tool', menu)
                
                # 在后台运行
                import threading
                def run_icon():
                    icon.run()
                
                threading.Thread(target=run_icon, daemon=True).start()
                
                self.tray_icon = icon
                return
            except ImportError:
                # 如果pystray不可用，使用原来的方法
                pass
            
            # 原来的系统托盘实现
            import win32api
            import win32gui
            import win32con
            
            # 定义托盘图标类
            class TrayIcon:
                def __init__(self, app):
                    self.app = app
                    self.hwnd = None
                    self.nid = None
                    self.create_tray_icon()
                
                def create_tray_icon(self):
                    # 注册窗口类
                    wc = win32gui.WNDCLASS()
                    wc.lpfnWndProc = self.window_proc
                    wc.hInstance = win32api.GetModuleHandle(None)
                    wc.lpszClassName = "NetworkConfigTray"
                    wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
                    wc.hbrBackground = win32con.COLOR_WINDOW + 1
                    class_atom = win32gui.RegisterClass(wc)
                    
                    # 创建窗口
                    self.hwnd = win32gui.CreateWindow(
                        class_atom,
                        "Network Config Tool",
                        win32con.WS_OVERLAPPED | win32con.WS_SYSMENU,
                        0, 0, 1, 1,
                        0, 0,
                        win32api.GetModuleHandle(None),
                        None
                    )
                    
                    # 创建托盘图标
                    icon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
                    nid = (self.hwnd, 100, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                           win32con.WM_USER + 100, icon, "Network Config Tool")
                    win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
                    self.nid = nid
                    
                    # 启动消息循环
                    import threading
                    self.message_thread = threading.Thread(target=self.pump_messages)
                    self.message_thread.daemon = True
                    self.message_thread.start()
                
                def window_proc(self, hwnd, msg, wparam, lparam):
                    if msg == win32con.WM_USER + 100:
                        if lparam == win32con.WM_RBUTTONUP:
                            # 右键点击，显示菜单
                            self.show_context_menu()
                        elif lparam == win32con.WM_LBUTTONDBLCLK:
                            # 双击，显示主窗口
                            self.app.root.deiconify()
                    elif msg == win32con.WM_COMMAND:
                        # 处理菜单命令
                        cmd = wparam
                        if cmd == 1:
                            self.app.root.deiconify()
                        elif cmd == 2:
                            self.app.set_auto_config()
                        elif cmd == 99:
                            self.app.on_close()
                    elif msg == win32con.WM_DESTROY:
                        try:
                            if self.nid:
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
                        
                        # 添加菜单项 - 使用完整的中文文本
                        win32gui.AppendMenu(menu, win32con.MF_STRING, 1, "显示主窗口")
                        win32gui.AppendMenu(menu, win32con.MF_STRING, 2, "自动获取配置")
                        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
                        win32gui.AppendMenu(menu, win32con.MF_STRING, 99, "退出程序")
                        
                        # 获取鼠标位置
                        pos = win32gui.GetCursorPos()
                        
                        # 显示菜单 - 使用TPM_RETURNCMD标志确保菜单正常显示
                        win32gui.SetForegroundWindow(self.hwnd)
                        result = win32gui.TrackPopupMenu(
                            menu,
                            win32con.TPM_LEFTALIGN | win32con.TPM_RETURNCMD,
                            pos[0],
                            pos[1],
                            0,
                            self.hwnd,
                            None
                        )
                        
                        # 直接处理菜单命令
                        if result == 1:
                            self.app.root.deiconify()
                        elif result == 2:
                            self.app.set_auto_config()
                        elif result == 99:
                            self.app.on_close()
                        
                        # 确保菜单消息被处理
                        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
                        
                    except Exception as e:
                        print(f"显示菜单失败: {e}")
                
                def pump_messages(self):
                    # 消息循环
                    try:
                        win32gui.PumpMessages()
                    except:
                        pass
            
            # 创建托盘图标
            self.tray_icon = TrayIcon(self)
            
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
    
    def apply_config_from_file(self):
        # 从配置文件应用配置到UI界面
        try:
            import os
            # 显示配置文件路径和内容
            config_status = f"配置文件路径: {self.config_file}\n"
            config_status += f"配置文件存在: {os.path.exists(self.config_file)}\n"
            
            if self.configs.get("profiles") and len(self.configs["profiles"]) > 0:
                config_status += "找到配置文件中的profiles，开始应用...\n"
                # 获取第一个配置（最新的配置）
                first_profile = self.configs["profiles"][0]
                config_status += f"第一个配置: {first_profile}\n"
                
                # 先应用非接口的配置值（IP、子网掩码、网关、DNS）
                # IP地址
                if "ip" in first_profile:
                    ip_value = first_profile["ip"]
                    config_status += f"设置IP地址: {ip_value}\n"
                    self.ip_var.set(ip_value)
                # 子网掩码
                if "subnet" in first_profile:
                    subnet_value = first_profile["subnet"]
                    config_status += f"设置子网掩码: {subnet_value}\n"
                    self.subnet_var.set(subnet_value)
                # 网关
                if "gateway" in first_profile:
                    gateway_value = first_profile["gateway"]
                    config_status += f"设置网关: {gateway_value}\n"
                    self.gateway_var.set(gateway_value)
                # DNS服务器
                if "dns" in first_profile:
                    dns_value = first_profile["dns"]
                    config_status += f"设置DNS: {dns_value}\n"
                    self.dns_var.set(dns_value)
                
                # 然后更新网络接口列表
                config_status += "更新网络接口列表...\n"
                self.update_interfaces()
                
                # 最后尝试设置接口值
                # 获取当前接口列表
                current_interfaces = self.interface_combo['values']
                config_status += f"当前接口列表: {current_interfaces}\n"
                
                # 接口
                if "interface" in first_profile and first_profile["interface"]:
                    interface_value = first_profile["interface"]
                    config_status += f"配置文件中的接口: {interface_value}\n"
                    if interface_value in current_interfaces:
                        config_status += "接口在列表中，设置为配置值\n"
                        self.interface_var.set(interface_value)
                    else:
                        config_status += "接口不在列表中，使用默认值\n"
            else:
                config_status += "配置文件中没有profiles或profiles为空\n"
            
            # 显示配置加载状态
            print(config_status)
            # 不显示消息框，避免干扰用户
        except Exception as e:
            error_msg = f"应用配置到UI时发生错误: {e}"
            print(error_msg)
            # 不显示错误消息框，避免干扰用户
            pass

# 创建主窗口
root = tk.Tk()
app = NetworkConfigApp(root)

# 启动应用
root.mainloop()