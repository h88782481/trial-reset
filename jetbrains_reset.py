"""
JetBrains 试用期重置工具类
通过删除相关文件和注册表来重置 JetBrains IDE 试用期，同时保留用户设置
"""

import os
import json
import shutil
import subprocess
from pathlib import Path


class JetBrainsReset:
    """JetBrains 试用期重置工具"""
    
    NAME = "JetBrains"
    INTERVAL_DAYS = 25  # 25天自动重置一次
    
    # JetBrains 进程名称
    PROCESSES = [
        "pycharm64.exe", "pycharm.exe",
        "webstorm64.exe", "webstorm.exe",
        "idea64.exe", "idea.exe",
        "clion64.exe", "clion.exe",
        "rider64.exe", "rider.exe",
        "goland64.exe", "goland.exe",
        "phpstorm64.exe", "phpstorm.exe",
        "rubymine64.exe", "rubymine.exe",
        "datagrip64.exe", "datagrip.exe",
        "aqua64.exe", "aqua.exe",
        "rustrover64.exe", "rustrover.exe",
        "fleet.exe",
        "dataspell64.exe", "dataspell.exe",
        "jetbrains-toolbox.exe",
    ]
    
    # 需要备份/恢复的设置项
    PRESERVE_ITEMS = [
        "options",           # 设置
        "codestyles",        # 代码格式化规则
        "colors",            # 配色方案
        "keymaps",           # 键盘快捷键
        "templates",         # 文件模板
        "fileTemplates",     # 文件模板
        "scratches",         # 临时文件
        "consoles",          # 数据库控制台
        "jdbc-drivers",      # 数据库驱动
        "extensions",        # 扩展
        "settingsSync",      # 设置同步
        "quicklists",        # 快捷列表
        "shelf",             # 搁置的更改
        "tasks",             # 任务
        "workspace",         # 工作区文件
        "plugins",           # 用户安装的插件
        "inspection",        # 检查配置
        "grazie",            # 语法检查数据
    ]
    
    # IDE 文件夹前缀
    IDE_PREFIXES = [
        "PyCharm", "WebStorm", "IntelliJ", "CLion", "Rider", "GoLand",
        "PhpStorm", "RubyMine", "DataGrip", "Aqua", "RustRover", "Fleet",
        "DataSpell", "Resharper", "dotMemory", "dotTrace"
    ]
    
    def __init__(self, backup_dir: Path):
        """
        初始化
        
        Args:
            backup_dir: 备份目录路径
        """
        self.backup_dir = backup_dir
    
    @staticmethod
    def get_running_processes() -> list:
        """获取正在运行的 JetBrains 进程列表"""
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
                        for jb_proc in JetBrainsReset.PROCESSES:
                            if proc_name == jb_proc.lower():
                                running.append(parts[0].strip('"'))
                                break
        except:
            pass
        return list(set(running))
    
    @staticmethod
    def kill_processes() -> list:
        """关闭所有正在运行的 JetBrains 进程"""
        killed = []
        for proc in JetBrainsReset.PROCESSES:
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
    def _is_ide_folder(name: str) -> bool:
        """检查文件夹是否为 IDE 配置文件夹"""
        for prefix in JetBrainsReset.IDE_PREFIXES:
            if name.startswith(prefix):
                return True
        return False
    
    @staticmethod
    def _clean_trial_from_options(options_dir: Path):
        """从 other.xml 中删除试用相关条目"""
        other_xml = options_dir / "other.xml"
        if not other_xml.exists():
            return
        
        try:
            with open(other_xml, 'r', encoding='utf-8') as f:
                content = f.read()
            
            import re
            
            match = re.search(
                r'<component name="PropertyService"><!\[CDATA\[(.*?)\]\]></component>',
                content, re.DOTALL
            )
            
            if match:
                json_str = match.group(1)
                try:
                    data = json.loads(json_str)
                    
                    remove_patterns = [
                        r'^evl',
                        r'^trial\.',
                        r'evalsprt',
                        r'\.runnable$',
                    ]
                    
                    if 'keyToString' in data:
                        keys_to_remove = [
                            key for key in data['keyToString']
                            if any(re.search(p, key, re.IGNORECASE) for p in remove_patterns)
                        ]
                        for key in keys_to_remove:
                            del data['keyToString'][key]
                    
                    new_json = json.dumps(data, indent=2)
                    content = re.sub(
                        r'<component name="PropertyService"><!\[CDATA\[.*?\]\]></component>',
                        f'<component name="PropertyService"><![CDATA[{new_json}]]></component>',
                        content, flags=re.DOTALL
                    )
                    
                    with open(other_xml, 'w', encoding='utf-8') as f:
                        f.write(content)
                except json.JSONDecodeError:
                    pass
        except:
            pass
    
    def perform_reset(self) -> str:
        """
        执行 JetBrains 试用期重置
        
        1. 备份 IDE 设置
        2. 删除整个 %APPDATA%\\JetBrains 文件夹
        3. 删除 %LOCALAPPDATA%\\JetBrains\\{IDE}* 文件夹
        4. 删除旧版 %USERPROFILE%\\.{IDE}* 文件夹
        5. 删除注册表 HKCU\\Software\\JavaSoft
        6. 从备份恢复 IDE 设置
        7. 清理已恢复设置中的试用相关条目
        
        返回：重置结果描述
        """
        appdata = Path(os.environ.get('APPDATA', ''))
        localappdata = Path(os.environ.get('LOCALAPPDATA', ''))
        user_profile = Path(os.environ.get('USERPROFILE', ''))
        
        jetbrains_folder = appdata / "JetBrains"
        
        deleted_items = []
        
        # === 步骤 1：备份 IDE 设置 ===
        if jetbrains_folder.exists():
            if self.backup_dir.exists():
                shutil.rmtree(self.backup_dir, ignore_errors=True)
            self.backup_dir.mkdir(parents=True, exist_ok=True)
            
            for item in jetbrains_folder.iterdir():
                if item.is_dir() and self._is_ide_folder(item.name):
                    ide_backup = self.backup_dir / item.name
                    ide_backup.mkdir(parents=True, exist_ok=True)
                    
                    for preserve_item in self.PRESERVE_ITEMS:
                        src = item / preserve_item
                        if src.exists():
                            dst = ide_backup / preserve_item
                            try:
                                if src.is_dir():
                                    shutil.copytree(src, dst)
                                else:
                                    shutil.copy2(src, dst)
                            except:
                                pass
        
        # === 步骤 2：删除整个 %APPDATA%\JetBrains ===
        if jetbrains_folder.exists():
            try:
                shutil.rmtree(jetbrains_folder)
                deleted_items.append("AppData/JetBrains")
            except Exception as e:
                deleted_items.append(f"错误: {e}")
        
        # === 步骤 3：删除 %LOCALAPPDATA%\JetBrains\{IDE}* ===
        jetbrains_local = localappdata / "JetBrains"
        if jetbrains_local.exists():
            for item in jetbrains_local.iterdir():
                if item.is_dir() and self._is_ide_folder(item.name):
                    try:
                        shutil.rmtree(item)
                        deleted_items.append(f"Local/{item.name}")
                    except:
                        pass
        
        # === 步骤 4：删除旧版文件夹 ===
        ide_names = ["WebStorm", "IntelliJ", "CLion", "Rider", "GoLand", "PhpStorm", "Resharper", "PyCharm", "DataGrip"]
        for ide in ide_names:
            for folder in user_profile.glob(f".{ide}*"):
                if folder.is_dir():
                    eval_path = folder / "config" / "eval"
                    if eval_path.exists():
                        try:
                            shutil.rmtree(eval_path)
                            deleted_items.append(f"{folder.name}/eval")
                        except:
                            pass
                    other_xml = folder / "config" / "options" / "other.xml"
                    if other_xml.exists():
                        try:
                            os.remove(other_xml)
                            deleted_items.append(f"{folder.name}/other.xml")
                        except:
                            pass
        
        # === 步骤 5：删除注册表 ===
        try:
            subprocess.run(
                ['reg', 'delete', 'HKEY_CURRENT_USER\\Software\\JavaSoft', '/f'],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            deleted_items.append("注册表: JavaSoft")
        except:
            pass
        
        # === 步骤 6：从备份恢复设置 ===
        restored = 0
        if self.backup_dir.exists():
            jetbrains_folder.mkdir(parents=True, exist_ok=True)
            
            for ide_backup in self.backup_dir.iterdir():
                if ide_backup.is_dir():
                    ide_folder = jetbrains_folder / ide_backup.name
                    ide_folder.mkdir(parents=True, exist_ok=True)
                    
                    for item in ide_backup.iterdir():
                        dst = ide_folder / item.name
                        try:
                            if item.is_dir():
                                shutil.copytree(item, dst)
                            else:
                                shutil.copy2(item, dst)
                            restored += 1
                        except:
                            pass
                    
                    # === 步骤 7：清理试用条目 ===
                    self._clean_trial_from_options(ide_folder / "options")
        
        if restored > 0:
            deleted_items.append(f"已恢复: {restored} 项")
        
        if deleted_items:
            display = deleted_items[:8]
            if len(deleted_items) > 8:
                display.append(f"...还有 {len(deleted_items) - 8} 项")
            return "\n".join(display)
        
        return "无需删除"


# 允许单独运行测试
if __name__ == "__main__":
    from pathlib import Path
    import os
    
    print("=" * 50)
    print("  JetBrains 试用期重置")
    print("=" * 50)
    print()
    
    running = JetBrainsReset.get_running_processes()
    if running:
        print(f"检测到运行中的进程: {', '.join(running)}")
        print("正在关闭...")
        JetBrainsReset.kill_processes()
        import time
        time.sleep(2)
    
    # 使用临时备份目录
    backup_dir = Path(os.environ.get('LOCALAPPDATA', '')) / "JetBrainsResetApp" / "backup"
    resetter = JetBrainsReset(backup_dir)
    result = resetter.perform_reset()
    print(f"\n结果:\n{result}")
    
    input("\n按回车键退出...")
