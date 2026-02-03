"""
试用期重置工具 - 系统托盘应用
整合 JetBrains 和 Navicat 的试用期重置功能

打包为 exe：
    pip install pyinstaller PyQt6
    pyinstaller --onefile --windowed --name="TrialReset" --icon=assets/icon.ico --add-data "assets;assets" --clean --noconfirm main.py
"""

import os
import sys
import json
import winreg
import ctypes
from datetime import datetime, timedelta
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QMessageBox
from PyQt6.QtGui import QIcon, QAction, QPixmap, QPainter, QColor, QFont
from PyQt6.QtCore import QTimer, Qt

from jetbrains_reset import JetBrainsReset
from navicat_reset import NavicatReset


APP_NAME = "Trial Reset"


def get_app_data_dir() -> Path:
    """获取应用数据目录"""
    app_data = Path(os.environ.get('LOCALAPPDATA', ''))
    app_dir = app_data / "TrialResetApp"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir


def get_config_path() -> Path:
    """获取配置文件路径"""
    return get_app_data_dir() / "config.json"


def get_config() -> dict:
    """读取配置"""
    config_path = get_config_path()
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "jetbrains": {"last_reset": None, "next_reset": None},
        "navicat": {"last_reset": None, "next_reset": None},
    }


def save_config(config: dict):
    """保存配置"""
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, default=str)


def get_exe_path() -> str:
    """获取当前可执行文件路径"""
    if getattr(sys, 'frozen', False):
        return sys.executable
    return os.path.abspath(sys.argv[0])


def is_in_autostart() -> bool:
    """检查是否已添加到开机自启"""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        try:
            winreg.QueryValueEx(key, APP_NAME)
            return True
        except FileNotFoundError:
            return False
        finally:
            winreg.CloseKey(key)
    except:
        return False


def add_to_autostart():
    """添加到开机自启"""
    try:
        exe_path = get_exe_path()
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{exe_path}"')
        winreg.CloseKey(key)
        return True
    except:
        return False


def remove_from_autostart():
    """从开机自启中移除"""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except:
        return False


def time_until_reset(tool_key: str, interval_days: int) -> str:
    """计算距离下次重置的时间"""
    config = get_config()
    tool_config = config.get(tool_key, {})
    
    if tool_config.get("next_reset"):
        try:
            next_reset = datetime.fromisoformat(tool_config["next_reset"])
            delta = next_reset - datetime.now()
            if delta.total_seconds() <= 0:
                return "需要重置！"
            days = delta.days
            hours = delta.seconds // 3600
            if days > 0:
                return f"{days}天 {hours}小时"
            elif hours > 0:
                return f"{hours}小时 {(delta.seconds % 3600) // 60}分钟"
            else:
                return f"{(delta.seconds % 3600) // 60}分钟"
        except:
            pass
    return "首次运行"


def create_icon() -> QIcon:
    """创建托盘图标"""
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).parent

    icon_path = base_path / "assets" / "icon.ico"

    if icon_path.exists():
        return QIcon(str(icon_path))

    # 动态生成图标
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor("#4CAF50"))
    painter.setPen(QColor("#388E3C"))
    painter.drawEllipse(4, 4, 56, 56)
    painter.setPen(QColor("white"))
    font = QFont("Arial", 16, QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "TR")
    painter.end()
    return QIcon(pixmap)


class TrayApp:
    """系统托盘应用"""
    
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        self.app.setApplicationName(APP_NAME)
        
        # 初始化工具
        self.jb_resetter = JetBrainsReset(get_app_data_dir() / "backup" / "jetbrains")
        
        # 首次运行自动添加到开机自启
        config = get_config()
        if not config.get("autostart_configured"):
            if add_to_autostart():
                config["autostart_configured"] = True
                save_config(config)
        
        # 创建托盘图标
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(create_icon())
        self._update_tooltip()
        
        # 创建右键菜单
        self._create_menu()
        
        self.tray.setContextMenu(self.menu)
        self.tray.show()
        
        # 状态更新定时器（每分钟）
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self._update_status)
        self.update_timer.start(60000)
        
        # 重置检查定时器（每小时）
        self.check_timer = QTimer()
        self.check_timer.timeout.connect(self._check_all_resets)
        self.check_timer.start(3600000)
        
        # 启动时检查
        self._check_all_resets()
    
    def _create_menu(self):
        """创建右键菜单"""
        self.menu = QMenu()
        
        # === JetBrains 部分 ===
        self.menu.addSection("JetBrains")
        
        self.jb_status_action = QAction(f"下次重置: {time_until_reset('jetbrains', JetBrainsReset.INTERVAL_DAYS)}")
        self.jb_status_action.setEnabled(False)
        self.menu.addAction(self.jb_status_action)
        
        self.jb_reset_action = QAction("重置 JetBrains")
        self.jb_reset_action.triggered.connect(self._manual_reset_jb)
        self.menu.addAction(self.jb_reset_action)
        
        # === Navicat 部分 ===
        self.menu.addSection("Navicat")
        
        self.nv_status_action = QAction(f"下次重置: {time_until_reset('navicat', NavicatReset.INTERVAL_DAYS)}")
        self.nv_status_action.setEnabled(False)
        self.menu.addAction(self.nv_status_action)
        
        self.nv_reset_action = QAction("重置 Navicat")
        self.nv_reset_action.triggered.connect(self._manual_reset_nv)
        self.menu.addAction(self.nv_reset_action)
        
        # === 通用选项 ===
        self.menu.addSeparator()
        
        self.autostart_action = QAction("开机自启 [开]" if is_in_autostart() else "开机自启 [关]")
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(is_in_autostart())
        self.autostart_action.triggered.connect(self._toggle_autostart)
        self.menu.addAction(self.autostart_action)
        
        self.menu.addSeparator()
        
        self.quit_action = QAction("退出")
        self.quit_action.triggered.connect(self._quit_app)
        self.menu.addAction(self.quit_action)
    
    def _update_tooltip(self):
        """更新托盘提示"""
        jb_time = time_until_reset('jetbrains', JetBrainsReset.INTERVAL_DAYS)
        nv_time = time_until_reset('navicat', NavicatReset.INTERVAL_DAYS)
        self.tray.setToolTip(f"{APP_NAME}\nJetBrains: {jb_time}\nNavicat: {nv_time}")
    
    def _update_status(self):
        """更新状态显示"""
        self.jb_status_action.setText(f"下次重置: {time_until_reset('jetbrains', JetBrainsReset.INTERVAL_DAYS)}")
        self.nv_status_action.setText(f"下次重置: {time_until_reset('navicat', NavicatReset.INTERVAL_DAYS)}")
        self._update_tooltip()
    
    def _toggle_autostart(self):
        """切换开机自启状态"""
        if is_in_autostart():
            if remove_from_autostart():
                self.autostart_action.setText("开机自启 [关]")
                self.autostart_action.setChecked(False)
                self.tray.showMessage(APP_NAME, "已禁用开机自启", QSystemTrayIcon.MessageIcon.Information, 2000)
        else:
            if add_to_autostart():
                self.autostart_action.setText("开机自启 [开]")
                self.autostart_action.setChecked(True)
                self.tray.showMessage(APP_NAME, "已启用开机自启", QSystemTrayIcon.MessageIcon.Information, 2000)
    
    def _check_all_resets(self):
        """检查所有工具是否需要重置"""
        self._check_reset('jetbrains', JetBrainsReset.INTERVAL_DAYS, JetBrainsReset)
        self._check_reset('navicat', NavicatReset.INTERVAL_DAYS, NavicatReset)
    
    def _check_reset(self, tool_key: str, interval_days: int, tool_class):
        """检查单个工具是否需要重置"""
        config = get_config()
        tool_config = config.get(tool_key, {})
        
        if tool_config.get("next_reset"):
            try:
                if datetime.now() >= datetime.fromisoformat(tool_config["next_reset"]):
                    self._show_auto_reset_dialog(tool_key, tool_class)
            except:
                pass
        elif tool_config.get("last_reset") is None:
            # 首次运行，设置下次重置时间
            if tool_key not in config:
                config[tool_key] = {}
            config[tool_key]["next_reset"] = (datetime.now() + timedelta(days=interval_days)).isoformat()
            save_config(config)
            self._update_status()
    
    def _show_auto_reset_dialog(self, tool_key: str, tool_class):
        """显示自动重置对话框"""
        tool_name = tool_class.NAME
        running = tool_class.get_running_processes()
        
        if running:
            running_str = ', '.join(running[:5])
            if len(running) > 5:
                running_str += f' 等 {len(running) - 5} 个'
            
            msg = QMessageBox()
            msg.setWindowTitle(APP_NAME)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText(f"{tool_name} 试用期即将到期！")
            msg.setInformativeText(
                f"检测到正在运行的进程：\n{running_str}\n\n"
                "点击「确定」关闭它们并重置。\n"
                "点击「稍后」推迟 1 小时。"
            )
            msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            msg.button(QMessageBox.StandardButton.Cancel).setText("稍后")
            msg.button(QMessageBox.StandardButton.Ok).setText("确定")
            
            reply = msg.exec()
            
            if reply == QMessageBox.StandardButton.Ok:
                tool_class.kill_processes()
                import time
                time.sleep(2)
                self._do_reset(tool_key, tool_class, auto=True)
            else:
                self._postpone_reset(tool_key)
        else:
            msg = QMessageBox()
            msg.setWindowTitle(APP_NAME)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText(f"{tool_name} 试用期即将到期！")
            msg.setInformativeText(
                f"没有检测到运行中的 {tool_name} 进程。\n\n"
                "点击「确定」立即重置。\n"
                "点击「稍后」推迟 1 小时。"
            )
            msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            msg.button(QMessageBox.StandardButton.Cancel).setText("稍后")
            msg.button(QMessageBox.StandardButton.Ok).setText("确定")
            
            reply = msg.exec()
            
            if reply == QMessageBox.StandardButton.Ok:
                self._do_reset(tool_key, tool_class, auto=True)
            else:
                self._postpone_reset(tool_key)
    
    def _postpone_reset(self, tool_key: str):
        """推迟重置 1 小时"""
        config = get_config()
        if tool_key not in config:
            config[tool_key] = {}
        config[tool_key]["next_reset"] = (datetime.now() + timedelta(hours=1)).isoformat()
        save_config(config)
        self._update_status()
        self.tray.showMessage(APP_NAME, "重置已推迟 1 小时", QSystemTrayIcon.MessageIcon.Information, 3000)
    
    def _manual_reset_jb(self):
        """手动重置 JetBrains"""
        self._manual_reset('jetbrains', JetBrainsReset)
    
    def _manual_reset_nv(self):
        """手动重置 Navicat"""
        self._manual_reset('navicat', NavicatReset)
    
    def _manual_reset(self, tool_key: str, tool_class):
        """手动重置"""
        tool_name = tool_class.NAME
        running = tool_class.get_running_processes()
        
        if running:
            running_str = ', '.join(running[:5])
            if len(running) > 5:
                running_str += f' 等 {len(running) - 5} 个'
            
            reply = QMessageBox.question(
                None, APP_NAME,
                f"检测到正在运行的 {tool_name} 进程：\n{running_str}\n\n"
                "点击「确定」关闭它们并重置。\n"
                "点击「取消」放弃操作。",
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
            )
            
            if reply == QMessageBox.StandardButton.Ok:
                tool_class.kill_processes()
                import time
                time.sleep(2)
                self._do_reset(tool_key, tool_class, auto=False)
        else:
            reply = QMessageBox.question(
                None, APP_NAME,
                f"确定要重置 {tool_name} 试用期吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._do_reset(tool_key, tool_class, auto=False)
    
    def _do_reset(self, tool_key: str, tool_class, auto=False):
        """执行重置"""
        try:
            # 执行重置
            if tool_key == 'jetbrains':
                result = self.jb_resetter.perform_reset()
            else:
                result = tool_class.perform_reset()
            
            # 更新配置
            config = get_config()
            if tool_key not in config:
                config[tool_key] = {}
            config[tool_key]["last_reset"] = datetime.now().isoformat()
            config[tool_key]["next_reset"] = (datetime.now() + timedelta(days=tool_class.INTERVAL_DAYS)).isoformat()
            save_config(config)
            
            self.tray.showMessage(
                APP_NAME,
                f"{tool_class.NAME} {'自动重置' if auto else '重置'}完成！\n{result}",
                QSystemTrayIcon.MessageIcon.Information, 5000
            )
            self._update_status()
        except Exception as e:
            self.tray.showMessage(APP_NAME, f"错误: {e}", QSystemTrayIcon.MessageIcon.Critical, 5000)
    
    def _quit_app(self):
        """退出应用"""
        self.tray.hide()
        self.app.quit()
    
    def run(self):
        """运行应用"""
        # 系统启动超过 2 分钟才显示通知
        uptime_ms = ctypes.windll.kernel32.GetTickCount64()
        if uptime_ms > 120000:
            jb_time = time_until_reset('jetbrains', JetBrainsReset.INTERVAL_DAYS)
            nv_time = time_until_reset('navicat', NavicatReset.INTERVAL_DAYS)
            self.tray.showMessage(
                APP_NAME,
                f"正在运行\nJetBrains: {jb_time}\nNavicat: {nv_time}",
                QSystemTrayIcon.MessageIcon.Information, 2000
            )
        sys.exit(self.app.exec())


def main():
    """主入口"""
    from PyQt6.QtNetwork import QLocalSocket, QLocalServer
    
    # 单例检查
    socket = QLocalSocket()
    socket.connectToServer(APP_NAME)
    if socket.waitForConnected(500):
        socket.close()
        sys.exit(0)
    
    server = QLocalServer()
    server.removeServer(APP_NAME)
    server.listen(APP_NAME)
    
    app = TrayApp()
    app.run()


if __name__ == "__main__":
    main()
