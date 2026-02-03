"""
Navicat Premium 试用期重置工具类
通过删除注册表项来重置 Navicat 试用期
"""

import winreg
import subprocess


class NavicatReset:
    """Navicat 试用期重置工具"""
    
    NAME = "Navicat"
    INTERVAL_DAYS = 10  # 10天自动重置一次
    
    # Navicat 进程名称
    PROCESSES = [
        "navicat.exe",
        "navicat_premium.exe", 
        "navicat premium.exe",
    ]
    
    @staticmethod
    def get_running_processes() -> list:
        """获取正在运行的 Navicat 进程列表"""
        running = []
        try:
            result = subprocess.run(
                ['tasklist', '/FO', 'CSV', '/NH'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            for line in result.stdout.strip().split('\n'):
                if line:
                    parts = line.split(',')
                    if parts:
                        proc_name = parts[0].strip('"').lower()
                        for nv_proc in NavicatReset.PROCESSES:
                            if proc_name == nv_proc.lower():
                                running.append(parts[0].strip('"'))
                                break
        except:
            pass
        return list(set(running))
    
    @staticmethod
    def kill_processes() -> list:
        """关闭所有正在运行的 Navicat 进程"""
        killed = []
        for proc in NavicatReset.PROCESSES:
            try:
                result = subprocess.run(
                    ['taskkill', '/F', '/IM', proc],
                    capture_output=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if result.returncode == 0:
                    killed.append(proc)
            except:
                pass
        return killed
    
    @staticmethod
    def _delete_key_recursive(hkey, path: str) -> bool:
        """递归删除注册表项及其所有子项"""
        try:
            key = winreg.OpenKey(hkey, path, 0, winreg.KEY_ALL_ACCESS)
            
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, 0)
                    NavicatReset._delete_key_recursive(hkey, f"{path}\\{subkey_name}")
                except OSError:
                    break
            
            winreg.CloseKey(key)
            winreg.DeleteKey(hkey, path)
            return True
        except (FileNotFoundError, PermissionError, Exception):
            return False
    
    @staticmethod
    def perform_reset() -> str:
        """
        执行 Navicat 试用期重置
        
        1. 删除 NavicatPremium\Update 注册表项
        2. 删除所有 Registration* 子项
        3. 删除包含 Info/ShellFolder 的 CLSID 条目
        
        返回：重置结果描述
        """
        deleted_items = []
        
        # === 步骤 1：删除 Update 项 ===
        path = r"Software\PremiumSoft\NavicatPremium\Update"
        if NavicatReset._delete_key_recursive(winreg.HKEY_CURRENT_USER, path):
            deleted_items.append("Update 项")
        
        # === 步骤 2：删除 Registration* 项 ===
        base_path = r"Software\PremiumSoft\NavicatPremium"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_READ)
            
            subkeys = []
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkeys.append(subkey_name)
                    i += 1
                except OSError:
                    break
            
            winreg.CloseKey(key)
            
            for subkey_name in subkeys:
                if subkey_name.startswith("Registration"):
                    full_path = f"{base_path}\\{subkey_name}"
                    if NavicatReset._delete_key_recursive(winreg.HKEY_CURRENT_USER, full_path):
                        deleted_items.append(subkey_name)
        except:
            pass
        
        # === 步骤 3：删除 CLSID 条目 ===
        clsid_base = r"Software\Classes\CLSID"
        clsid_deleted = 0
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_base, 0, winreg.KEY_READ)
            
            clsid_keys = []
            i = 0
            while True:
                try:
                    clsid_name = winreg.EnumKey(key, i)
                    clsid_keys.append(clsid_name)
                    i += 1
                except OSError:
                    break
            
            winreg.CloseKey(key)
            
            for clsid_name in clsid_keys:
                clsid_path = f"{clsid_base}\\{clsid_name}"
                should_delete = False
                
                try:
                    clsid_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path, 0, winreg.KEY_READ)
                    
                    j = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(clsid_key, j)
                            if subkey_name in ("Info", "ShellFolder"):
                                should_delete = True
                                break
                            j += 1
                        except OSError:
                            break
                    
                    winreg.CloseKey(clsid_key)
                    
                    if should_delete:
                        if NavicatReset._delete_key_recursive(winreg.HKEY_CURRENT_USER, clsid_path):
                            clsid_deleted += 1
                except:
                    continue
        except:
            pass
        
        if clsid_deleted > 0:
            deleted_items.append(f"CLSID ({clsid_deleted} 项)")
        
        if deleted_items:
            return "已删除: " + ", ".join(deleted_items)
        return "无需删除"


# 允许单独运行测试
if __name__ == "__main__":
    print("=" * 50)
    print("  Navicat Premium 试用期重置")
    print("=" * 50)
    print()
    
    running = NavicatReset.get_running_processes()
    if running:
        print(f"检测到运行中的进程: {', '.join(running)}")
        print("正在关闭...")
        NavicatReset.kill_processes()
        import time
        time.sleep(2)
    
    result = NavicatReset.perform_reset()
    print(f"\n结果: {result}")
    
    input("\n按回车键退出...")
