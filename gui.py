from __future__ import annotations

import asyncio
import json
import random
import sys
import threading
import time
from pathlib import Path
from typing import Dict, Optional
from urllib.parse import quote_plus

from PyQt6 import QtCore, QtGui, QtWidgets
from html import escape as html_escape

# 依赖于 main.py 中的逻辑函数与常量
from main import (
    USER_AGENTS,
    SEARCH_CONFIG,
    compute_remaining_searches,
    fetch_rewards_userinfo,
    generate_random_query,
    _perform_bing_search_like_human,
    _maybe_human_scroll,
    _maybe_click_one_result,
    get_cookie_files,
    _inject_mobile_spoofing,
    _sanitize_filename,
    _get_exe_dir,
    _get_bundle_dir,
    _format_console_dashboard,
)

# Playwright 异步 API（用于浏览器自动化）
from playwright.async_api import async_playwright


class RewardsGUI(QtWidgets.QMainWindow):
    # 跨线程 UI 更新信号
    log_signal = QtCore.pyqtSignal(str, str)           # (message, level)
    status_signal = QtCore.pyqtSignal(str, str)        # (text, color)
    dashboard_signal = QtCore.pyqtSignal(str, object)  # (dashboard text, account_name)
    refresh_accounts_signal = QtCore.pyqtSignal()      # 刷新账号列表


    def __init__(self):
        super().__init__()
        self.setWindowTitle("My QQ952904514")
        self.resize(1100, 700)

        # 设置窗口左上角标题图标（支持源码运行与打包运行的资源路径）
        try:
            icon_candidates = [
                _get_exe_dir() / "Assets" / "ico.png",
                _get_bundle_dir() / "Assets" / "ico.png",
            ]
            for icon_path in icon_candidates:
                if icon_path.exists():
                    self.setWindowIcon(QtGui.QIcon(str(icon_path)))
                    break
        except Exception:
            pass

        # 运行状态
        self.is_running: bool = False
        self.running_tasks: Dict[str, threading.Thread] = {}
        self.task_stop_flags: Dict[str, bool] = {}
        self.account_status: Dict[str, str] = {}
        # 各账号对应的最新仪表盘文本
        self.dashboard_texts: Dict[str, str] = {}

        # 配置与设备选择
        self.config_path = Path("Assets/config.json")
        self.cookies_dir = Path("Assets/cookies")
        # old GUI 默认非无头运行，这里与旧版保持一致
        self.headless_mode = False
        self.device_windows = True
        self.device_iphone = False

        # 选中账号
        self.selected_cookie: Optional[Path] = None

        self._build_ui()
        self._connect_signals()
        # 仪表盘显示刷新定时器（仅刷新界面显示，不改变任务逻辑），0.1 秒一次
        self.dashboard_timer = QtCore.QTimer(self)
        self.dashboard_timer.setInterval(100)
        self.dashboard_timer.timeout.connect(self._show_dashboard_for_selected)
        self.dashboard_timer.start()
        self.load_config()
        self.refresh_accounts()

    # region UI 构建
    def _build_ui(self):
        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 顶部选项区
        opts = QtWidgets.QHBoxLayout()
        self.cb_headless = QtWidgets.QCheckBox("后台模式")
        # 默认关闭无头模式，与旧版保持一致；加载配置后会覆盖
        self.cb_headless.setChecked(False)
        self.cb_windows = QtWidgets.QCheckBox("Windows")
        self.cb_windows.setChecked(True)
        self.cb_iphone = QtWidgets.QCheckBox("iPhone")
        self.cb_iphone.setChecked(False)
        opts.addWidget(self.cb_headless)
        opts.addWidget(self.cb_windows)
        opts.addWidget(self.cb_iphone)
        opts.addStretch(1)

        # 按钮区
        self.btn_refresh = QtWidgets.QPushButton("刷新账号")
        self.btn_add_account = QtWidgets.QPushButton("添加账号")
        self.btn_add_account.setToolTip("登录后点击右上角头像自动保存")
        self.btn_delete_account = QtWidgets.QPushButton("删除账号")
        self.btn_start = QtWidgets.QPushButton("开始 (单账号)")
        self.btn_batch = QtWidgets.QPushButton("批量开始 (全部账号)")
        self.btn_stop = QtWidgets.QPushButton("停止当前")
        self.btn_stop_all = QtWidgets.QPushButton("全部停止")
        opts.addWidget(self.btn_refresh)
        opts.addWidget(self.btn_add_account)
        opts.addWidget(self.btn_delete_account)
        opts.addWidget(self.btn_start)
        opts.addWidget(self.btn_batch)
        opts.addWidget(self.btn_stop)
        opts.addWidget(self.btn_stop_all)

        layout.addLayout(opts)

        # 中间分栏：左侧账号列表，右侧仪表板与日志
        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        layout.addWidget(splitter, 1)

        left = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        self.accounts_list = QtWidgets.QListWidget()
        self.accounts_list.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.accounts_list.currentItemChanged.connect(self._on_account_selection_changed)
        left_layout.addWidget(QtWidgets.QLabel("账号列表"))
        left_layout.addWidget(self.accounts_list, 1)

        splitter.addWidget(left)

        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        self.dashboard = QtWidgets.QTextEdit()
        self.dashboard.setReadOnly(True)
        self.dashboard.setMinimumHeight(160)
        # 仪表盘背景不可选中：禁用文本交互与焦点
        self.dashboard.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.dashboard.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        self.btn_clear_log = QtWidgets.QPushButton("清空日志")

        right_layout.addWidget(QtWidgets.QLabel("仪表板"))
        right_layout.addWidget(self.dashboard)

        log_header = QtWidgets.QHBoxLayout()
        log_header.setContentsMargins(0, 0, 0, 0)
        log_header.setSpacing(6)
        log_header.addWidget(QtWidgets.QLabel("日志"))
        log_header.addStretch(1)
        log_header.addWidget(self.btn_clear_log)
        right_layout.addLayout(log_header)
        right_layout.addWidget(self.log_view, 1)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        # 状态栏（默认隐藏，移除左下角日志展示）
        self.status_bar = QtWidgets.QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.hide()

        # 初始按钮状态（未运行）
        self._set_buttons_enabled(idle=True)

    def _connect_signals(self):
        # 自定义信号
        self.log_signal.connect(self.log)
        self.status_signal.connect(self.update_status)
        self.dashboard_signal.connect(self._update_dashboard_text)
        self.refresh_accounts_signal.connect(self.refresh_accounts)
        
        # UI 控件事件
        self.cb_headless.toggled.connect(self._on_headless_changed)
        self.cb_windows.toggled.connect(self._on_device_changed)
        self.cb_iphone.toggled.connect(self._on_device_changed)
        self.btn_refresh.clicked.connect(self.refresh_accounts)
        self.btn_add_account.clicked.connect(self.add_account)
        self.btn_delete_account.clicked.connect(self.delete_account)
        self.btn_start.clicked.connect(self.start_task)
        self.btn_batch.clicked.connect(self.start_batch_tasks)
        self.btn_stop.clicked.connect(self.stop_task)
        self.btn_stop_all.clicked.connect(self.stop_all_tasks)
        self.btn_clear_log.clicked.connect(self.clear_log)

    # endregion

    # region 添加账号（登录并保存 Cookie）
    def add_account(self):
        if self.is_running:
            self.log_signal.emit("任务运行中，无法添加账号", "WARN")
            return
        self.log_signal.emit("准备添加新账号...", "INFO")
        self._set_buttons_enabled(idle=False)
        
        # 在新线程中运行异步登录
        th = threading.Thread(target=self._run_add_account_thread, daemon=True)
        th.start()

    def _run_add_account_thread(self):
        try:
            asyncio.run(self._add_account_async())
        except Exception as e:
            self.log_signal.emit(f"添加账号流程异常: {e}", "ERROR")
        finally:
            self._set_buttons_enabled(idle=True)
            self.refresh_accounts_signal.emit()

    async def _add_account_async(self):
        """登录并保存cookie（异步版本）"""
        cookies_dir = _get_exe_dir() / "Assets" / "cookies"
        cookies_dir.mkdir(parents=True, exist_ok=True)
        
        async with async_playwright() as p:
            browser = await p.chromium.launch(channel="msedge", headless=False)
            context = await browser.new_context()
            page = await context.new_page()
            
            try:
                self.log_signal.emit("正在打开 Bing Rewards 页面...", "INFO")
                await page.goto("https://rewards.bing.com/?ref=rewardspanel")
                
                self.log_signal.emit("请在浏览器中登录您的Microsoft账户...", "INFO")
                
                # 等待用户登录（检测登录成功）
                try:
                    # 等待账户元素出现或超时30分钟
                    await page.wait_for_selector(
                        "#mectrl_currentAccount_secondary",
                        timeout=1800000
                    )
                except Exception:
                    self.log_signal.emit("登录超时或已取消", "WARN")
                    return
                
                # 获取账户名
                account_name = None
                try:
                    account_element = await page.query_selector("#mectrl_currentAccount_secondary")
                    if account_element:
                        text = await account_element.text_content()
                        if text and text.strip():
                            account_name = text.strip()
                except Exception:
                    pass
                
                if not account_name:
                    account_name = "未知账户"
                    
                account_name = _sanitize_filename(account_name)
                
                cookies = await context.cookies()
                cookie_file = cookies_dir / f"{account_name}.json"
                
                cookie_file.write_text(json.dumps(cookies, ensure_ascii=False, indent=2), encoding="utf-8")
                
                self.log_signal.emit(f"Cookie已保存: {account_name}", "SUCCESS")
                
            except Exception as e:
                self.log_signal.emit(f"登录失败: {str(e)}", "ERROR")
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass


    # region 删除账号
    def delete_account(self):
        item = self.accounts_list.currentItem()
        if not item:
            self.log_signal.emit("请先选择要删除的账号", "WARN")
            return

        account = item.text()
        cookie_path = Path(item.data(QtCore.Qt.ItemDataRole.UserRole))

        # 若有运行中的任务，先提醒并尝试停止
        running = [k for k in list(self.running_tasks.keys()) if k.startswith(account + "_")]
        if running:
            ret = QtWidgets.QMessageBox.warning(
                self,
                "删除账号",
                f"账号 {account} 正在运行任务，删除前会先停止这些任务。是否继续？",
                QtWidgets.QMessageBox.StandardButton.Ok,
                QtWidgets.QMessageBox.StandardButton.Cancel,
            )
            if ret != QtWidgets.QMessageBox.StandardButton.Ok:
                return
            for name in running:
                self.task_stop_flags[name] = True

        confirm = QtWidgets.QMessageBox.question(
            self,
            "删除账号",
            f"确定删除账号 {account} 的 Cookie 文件吗？\n{cookie_path}",
            QtWidgets.QMessageBox.StandardButton.Ok,
            QtWidgets.QMessageBox.StandardButton.Cancel,
        )
        if confirm != QtWidgets.QMessageBox.StandardButton.Ok:
            return

        try:
            if cookie_path.exists():
                cookie_path.unlink()
            self.account_status.pop(account, None)
            self.log_signal.emit(f"账号 {account} 已删除。", "SUCCESS")
        except Exception as e:
            self.log_signal.emit(f"删除账号失败: {e}", "ERROR")
            return

        self.refresh_accounts()
    # endregion

    # region 配置
    def load_config(self):
        try:
            if self.config_path.exists():
                cfg = json.loads(self.config_path.read_text(encoding="utf-8"))
                self.headless_mode = bool(cfg.get("headless", True))
                self.device_windows = bool(cfg.get("device_windows", True))
                self.device_iphone = bool(cfg.get("device_iphone", False))
                self.cb_headless.setChecked(self.headless_mode)
                self.cb_windows.setChecked(self.device_windows)
                self.cb_iphone.setChecked(self.device_iphone)
        except Exception as e:
            self.log_signal.emit(f"读取配置失败: {e}", "ERROR")

    def save_config(self):
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            cfg = {
                "headless": self.headless_mode,
                "device_windows": self.device_windows,
                "device_iphone": self.device_iphone,
            }
            self.config_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            self.log_signal.emit(f"保存配置失败: {e}", "ERROR")

    # endregion

    # region 账号列表
    def refresh_accounts(self):
        self.accounts_list.clear()
        try:
            # 与原始逻辑保持一致：内部自动从可写目录/资源目录查找
            files = get_cookie_files()
        except Exception:
            files = sorted(self.cookies_dir.glob("*.json"))
        for f in files:
            # 优先基于当前运行线程计算账号状态，避免一台设备结束后整体被标记为空闲
            running_devs = [
                name.split("_", 1)[1]
                for name in list(self.running_tasks.keys())
                if name.startswith(f.stem + "_")
            ]
            # 去重并排序（Windows 在前，iPhone 在后）
            order = {"windows": 0, "iphone": 1}
            running_devs = list(dict.fromkeys(running_devs))
            running_devs.sort(key=lambda d: order.get(d, 99))

            if running_devs:
                status = f"运行中: {', '.join(running_devs)}"
            else:
                status = self.account_status.get(f.stem) or "空闲"
            # 同步记录，便于其他位置读取
            self.account_status[f.stem] = status

            label = f.stem
            item = QtWidgets.QListWidgetItem(label)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, str(f))
            # 状态文本
            if status:
                item.setToolTip(status)
                # 运行中使用轻微绿色高亮
                if "运行中" in status:
                    item.setForeground(QtGui.QBrush(QtGui.QColor("#2e7d32")))
            self.accounts_list.addItem(item)

    # endregion

    # region 日志/状态/仪表板
    def log(self, message: str, level: str = "INFO"):
        ts = time.strftime("%H:%M:%S")
        color = {
            "INFO": "#333",
            "SUCCESS": "#2e7d32",
            "WARN": "#f57c00",
            "ERROR": "#c62828",
        }.get(level, "#333")
        safe = html_escape(message)
        html = f"<span style='color:{color}'>[{ts}] [{level}] {safe}</span>"
        self.log_view.append(html)
        self.log_view.ensureCursorVisible()

    def clear_log(self):
        self.log_view.clear()

    def update_status(self, text: str, color: str = "black"):
        # 使用状态栏显示简短状态
        self.status_bar.showMessage(text)
        # 可选：颜色标签（保留占位）
        _ = color

    def _current_selected_account(self) -> Optional[str]:
        item = self.accounts_list.currentItem()
        return item.text() if item else None

    def _update_dashboard_text(self, text: str, account_name: Optional[str] = None):
        # 记录每个账号对应的仪表盘文本
        if account_name:
            self.dashboard_texts[account_name] = text
        current = self._current_selected_account()
        if account_name is None or account_name == current:
            self.dashboard.setPlainText(text)

    def _show_dashboard_for_selected(self):
        name = self._current_selected_account()
        if not name:
            self.dashboard.clear()
            return
        text = self.dashboard_texts.get(name)
        if text:
            self.dashboard.setPlainText(text)
        else:
            self.dashboard.clear()

    def _on_account_selection_changed(self, current, previous):  # noqa: ARG002
        self._show_dashboard_for_selected()

    def _render_dashboard(self, *, account_name: Optional[str] = None, userinfo=None, stats=None, status_text: str = "", search_index=None, search_total=None):
        # 若未能识别 firstName，使用 account_name 作为替代显示
        safe_userinfo = dict(userinfo or {})
        if account_name and not safe_userinfo.get("firstName"):
            safe_userinfo["firstName"] = account_name

        try:
            text = _format_console_dashboard(
                userinfo=safe_userinfo if safe_userinfo else None,
                stats=stats,
                status_text=status_text,
                search_index=search_index,
                search_total=search_total,
            )
        except Exception:
            # 回退到本地简单格式
            lines = []
            if safe_userinfo:
                lines.append(f"账户: {safe_userinfo.get('firstName','')}  积分: {safe_userinfo.get('availablePoints','-')}")
            if stats:
                lines.append(f"设备: {stats.get('device_label','-')}  当前/最大: {stats.get('current_points','-')} / {stats.get('max_points','-')}")
                lines.append(f"剩余搜索: {stats.get('remaining_searches','-')} / {stats.get('total_searches','-')}")
            if status_text:
                lines.append(f"状态: {status_text}")
            # 统一状态仅展示剩余次数，移除进度数字
            text = "\n".join(lines)

        self.dashboard_signal.emit(text, account_name)

    async def _periodic_dashboard_refresh(self, context, device_type: str, account_name: Optional[str], task_name: Optional[str], interval: float = 0.1):
        """周期性拉取真实 Rewards 数据并刷新仪表盘，仅影响显示，不改变任务执行速度。"""
        while True:
            if self._should_stop(task_name, account_name):
                break
            try:
                userinfo_live = await fetch_rewards_userinfo(context)
                stats_live = compute_remaining_searches(userinfo_live, device_type)
                remaining_live = stats_live.get("remaining_searches", 0)
                total_live = stats_live.get("total_searches", 0)
                done_live = max(0, total_live - remaining_live) if total_live else 0

                self._render_dashboard(
                    account_name=account_name,
                    userinfo=userinfo_live,
                    stats=stats_live,
                    status_text=f"剩余 {remaining_live} 次",
                    search_index=None,
                    search_total=None,
                )
            except Exception:
                # 网络/接口失败时忽略，继续下一次
                pass
            await asyncio.sleep(interval)

    # endregion

    # region 任务控制
    def _selected_devices(self) -> list[str]:
        devices: list[str] = []
        if self.cb_windows.isChecked():
            devices.append("windows")
        if self.cb_iphone.isChecked():
            devices.append("iphone")
        return devices

    def start_task(self):
        item = self.accounts_list.currentItem()
        if not item:
            self.log_signal.emit("请先选择一个账号（cookies 文件）", "WARN")
            return
        cookie_path = Path(item.data(QtCore.Qt.ItemDataRole.UserRole))
        self.selected_cookie = cookie_path
        devices = self._selected_devices()
        if not devices:
            self.log_signal.emit("请至少选择一个设备（Windows 或 iPhone）", "WARN")
            return
        # 重置该账号的停止标记，避免上一次停止后遗留的 True 影响本次启动
        self._reset_stop_flags_for_account(cookie_path.stem)
        self.is_running = True
        # 启动前清除全局停止标志，并更新按钮状态
        self.task_stop_flags.pop("__global", None)
        self._set_buttons_enabled(idle=False)
        for dev in devices:
            task_name = f"{cookie_path.stem}_{dev}"
            self.task_stop_flags[task_name] = False
            self._launch_single(cookie_path, dev, task_name)
        self.update_status("任务已启动", "green")

    def start_batch_tasks(self):
        devices = self._selected_devices()
        if not devices:
            self.log_signal.emit("请至少选择一个设备（Windows 或 iPhone）", "WARN")
            return
        files = [Path(self.accounts_list.item(i).data(QtCore.Qt.ItemDataRole.UserRole)) for i in range(self.accounts_list.count())]
        if not files:
            self.log_signal.emit("未找到任何账号 cookies", "WARN")
            return
        self.is_running = True
        self.task_stop_flags.pop("__global", None)
        self._set_buttons_enabled(idle=False)
        for f in files:
            # 批量模式下也先重置账号级停止标记，防止历史 stop 影响
            self._reset_stop_flags_for_account(f.stem)
            for dev in devices:
                task_name = f"{f.stem}_{dev}"
                self.task_stop_flags[task_name] = False
                self._launch_single(f, dev, task_name)
        self.update_status("批量任务已启动", "green")

    def stop_task(self):
        # 停止当前选中账号的所有设备任务
        item = self.accounts_list.currentItem()
        if not item:
            self.log_signal.emit("未选择账号", "WARN")
            return
        account = item.text()
        to_stop = [k for k in list(self.running_tasks.keys()) if k.startswith(account + "_")]
        if not to_stop:
            self.log_signal.emit("该账号没有正在运行的任务", "WARN")
            return
        for name in to_stop:
            self.task_stop_flags[name] = True
        self.update_status("停止中...", "red")

    def stop_all_tasks(self):
        if not self.running_tasks:
            self.log_signal.emit("当前无任务在运行", "WARN")
            return
        self.is_running = False
        # 设置全局停止标志，所有循环会检测到
        self.task_stop_flags["__global"] = True
        for name in list(self.running_tasks.keys()):
            self.task_stop_flags[name] = True
        # 更新按钮状态
        self._set_buttons_enabled(idle=False)
        self.update_status("全部停止中...", "red")

    def _launch_single(self, cookie_file: Path, device_type: str, task_name: str):
        th = threading.Thread(target=self._run_task_thread, args=(cookie_file, device_type, task_name), daemon=True)
        th.start()
        self.running_tasks[task_name] = th
        # 启动后按仍在运行的设备集合更新账号状态
        running_devs = [
            name.split("_", 1)[1]
            for name in list(self.running_tasks.keys())
            if name.startswith(cookie_file.stem + "_")
        ]
        order = {"windows": 0, "iphone": 1}
        running_devs = list(dict.fromkeys(running_devs))
        running_devs.sort(key=lambda d: order.get(d, 99))
        self.account_status[cookie_file.stem] = f"运行中: {', '.join(running_devs)}" if running_devs else "空闲"
        self.refresh_accounts_signal.emit()

    def _run_task_thread(self, cookie_file: Path, device_type: str, task_name: str):
        try:
            asyncio.run(self._execute_single_account(cookie_file, device_type, task_name))
        except Exception as e:
            self.log_signal.emit(f"任务异常: {e}", "ERROR")
        finally:
            # 清理运行标记
            self.running_tasks.pop(task_name, None)
            # 若该账号仍有其他设备在运行，则保持运行中状态
            running_devs = [
                name.split("_", 1)[1]
                for name in list(self.running_tasks.keys())
                if name.startswith(cookie_file.stem + "_")
            ]
            order = {"windows": 0, "iphone": 1}
            running_devs = list(dict.fromkeys(running_devs))
            running_devs.sort(key=lambda d: order.get(d, 99))
            if running_devs:
                self.account_status[cookie_file.stem] = f"运行中: {', '.join(running_devs)}"
            else:
                self.account_status[cookie_file.stem] = "空闲"
            self.refresh_accounts_signal.emit()
            # 如果没有任何任务在运行，恢复按钮状态
            if not self.running_tasks:
                self.is_running = False
                self._set_buttons_enabled(idle=True)
                self.update_status("已停止或完成", "green")

    async def _execute_single_account(self, cookie_file: Path, device_type: str, task_name: str):
        prefix = f"[{cookie_file.stem}] "
        # 与 main.USER_AGENTS 对齐（iphone 为可调用的随机 UA）
        ua_value = USER_AGENTS[device_type]
        user_agent = ua_value() if callable(ua_value) else ua_value
        self.log_signal.emit(f"{prefix}启动浏览器，设备: {device_type}", "INFO")
        async with async_playwright() as p:
            # iPhone 模式下在 headless 模式添加额外启动参数以增强移动设备伪装
            launch_kwargs = {"channel": "msedge", "headless": self.headless_mode}
            if self.headless_mode and device_type == "iphone":
                # 添加启动参数强制移动设备特征
                launch_kwargs["args"] = [
                    "--disable-blink-features=AutomationControlled",  # 隐藏自动化特征
                    "--disable-features=IsolateOrigins,site-per-process",  # 减少检测
                    "--user-agent=" + user_agent,  # 在启动时就设置 UA
                ]
                self.log_signal.emit(f"{prefix}后台模式下使用增强型移动设备伪装", "INFO")
            browser = await p.chromium.launch(**launch_kwargs)
            ctx_kwargs = self._context_kwargs(p, device_type, user_agent)
            # 创建上下文后再注入移动伪装 & 添加 cookies（与 main.py 对齐）
            context = await browser.new_context(**ctx_kwargs)
            if device_type == "iphone":
                try:
                    await _inject_mobile_spoofing(context)
                except Exception:
                    pass
            # 加载 cookie JSON 并注入
            try:
                cookies = json.loads(Path(cookie_file).read_text(encoding="utf-8"))
                if isinstance(cookies, list) and cookies:
                    await context.add_cookies(cookies)
            except Exception:
                pass
            page = await context.new_page()
            
            # iPhone 模式下在页面级别也注入伪装脚本（双重保险）
            if device_type == "iphone":
                try:
                    await page.add_init_script("""
                        // 页面级别的移动设备伪装（补充 context 级别的注入）
                        Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 5});
                        Object.defineProperty(navigator, 'platform', {get: () => 'iPhone'});
                        Object.defineProperty(navigator, 'vendor', {get: () => 'Apple Computer, Inc.'});
                    """)
                    # 验证设备特征（用于诊断）
                    await page.goto("about:blank")
                    ua_check = await page.evaluate("navigator.userAgent")
                    platform_check = await page.evaluate("navigator.platform")
                    touch_check = await page.evaluate("navigator.maxTouchPoints")
                    self.log_signal.emit(f"{prefix}设备验证 - UA包含iPhone: {'iPhone' in ua_check}, Platform: {platform_check}, 触点: {touch_check}", "INFO")
                except Exception as e:
                    self.log_signal.emit(f"{prefix}页面级伪装警告: {e}", "WARN")
            
            try:
                await self._run_rewards_search_gui(page, context, device_type, cookie_file.stem, task_name)
                self.log_signal.emit(f"{prefix}任务执行完成！", "SUCCESS")
            except Exception as e:
                self.log_signal.emit(f"{prefix}执行错误: {e}", "ERROR")
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass

    async def _run_rewards_search_gui(
        self,
        page,
        context,
        device_type: str,
        account_name: Optional[str] = None,
        task_name: Optional[str] = None,
    ):
        prefix = f"[{account_name}] " if account_name else ""
        self.log_signal.emit(f"{prefix}正在获取 Rewards 信息...", "INFO")
        try:
            userinfo = await fetch_rewards_userinfo(context)
            stats = compute_remaining_searches(userinfo, device_type)
        except Exception as e:
            self.log_signal.emit(f"{prefix}无法获取 Rewards 数据: {e}", "ERROR")
            return

        remaining = stats["remaining_searches"]
        total = stats["total_searches"]
        self.log_signal.emit(f"{prefix}当前设备: {stats['device_label']}", "INFO")
        self.log_signal.emit(f"{prefix}剩余搜索次数: {remaining} / {total}", "INFO")
        # 初次渲染附带剩余次数，避免后续周期刷新覆盖状态前仍然显示空
        self._render_dashboard(
            account_name=account_name,
            userinfo=userinfo,
            stats=stats,
            status_text=f"剩余 {remaining} 次",
            search_index=None,
            search_total=None,
        )
        
        # 启动周期性真实数据刷新任务（仅更新显示）
        refresh_task = asyncio.create_task(self._periodic_dashboard_refresh(context, device_type, account_name, task_name, interval=0.1))

        # 检查当前设备是否需要执行任务
        if total <= 0 or remaining <= 0:
            self.log_signal.emit(f"{prefix}当前设备任务已完成或无需执行！", "SUCCESS")
            self._render_dashboard(account_name=account_name, userinfo=userinfo, stats=stats, status_text="剩余 0 次", search_index=None, search_total=None)
            try:
                refresh_task.cancel()
            except Exception:
                pass
            return

        if stats["max_points"] <= 0:
            self.log_signal.emit(f"{prefix}今日任务已完成！", "SUCCESS")
            self._render_dashboard(account_name=account_name, userinfo=userinfo, stats=stats, status_text="剩余 0 次", search_index=None, search_total=None)
            try:
                refresh_task.cancel()
            except Exception:
                pass
            return

        delay_cfg = SEARCH_CONFIG["pc"] if device_type == "windows" else SEARCH_CONFIG["mobile"]
        min_ms = delay_cfg["min_delay_ms"]
        max_ms = delay_cfg["max_delay_ms"]

        if page.url == "about:blank":
            await page.goto("https://www.bing.com", wait_until="domcontentloaded")
        await asyncio.sleep(2)

        # 使用实时剩余次数动态计算循环次数，避免因初始估计误差导致提前结束
        current_remaining = remaining
        executed = 0
        # 保险上限，防止意外无限循环（通常 total 就是每天的最大搜索次数）
        safety_cap = max(total, remaining, 30) * 3

        while current_remaining > 0 and executed < safety_cap:
            if self._should_stop(task_name, account_name):
                self.log_signal.emit(f"{prefix}任务已停止", "WARN")
                self._render_dashboard(
                    account_name=account_name,
                    userinfo=userinfo,
                    stats=stats,
                    status_text=f"已停止，剩余 {current_remaining} 次",
                    search_index=None,
                    search_total=None,
                )
                break

            executed += 1
            query = generate_random_query()
            search_url = f"https://www.bing.com/search?q={quote_plus(query)}"
            display_total = current_remaining if current_remaining > 0 else remaining
            try:
                self.log_signal.emit(f"{prefix}搜索: {query} | 剩余 {current_remaining} 次", "INFO")
                self.status_signal.emit(f"搜索中，剩余 {current_remaining} 次", "orange")

                ok = False
                try:
                    ok = await asyncio.wait_for(_perform_bing_search_like_human(page, query), timeout=25)
                except asyncio.TimeoutError:
                    self.log_signal.emit(f"{prefix}搜索超时，重试...", "WARN")
                    try:
                        await page.goto("https://www.bing.com", wait_until="domcontentloaded", timeout=20000)
                    except Exception:
                        pass
                    ok = False
                except Exception:
                    ok = False

                if not ok:
                    await page.goto(search_url, wait_until="domcontentloaded")
                await _maybe_human_scroll(page)
                await _maybe_click_one_result(page)

            except Exception as e:
                self.log_signal.emit(f"{prefix}搜索失败: {e}", "WARN")
                try:
                    await page.goto("https://www.bing.com", wait_until="domcontentloaded")
                except Exception:
                    pass

            # 执行完一次搜索后，实时获取最新剩余次数，避免因积分延迟导致的提前结束
            try:
                userinfo_live = await fetch_rewards_userinfo(context)
                stats_live = compute_remaining_searches(userinfo_live, device_type)
                current_remaining = stats_live.get("remaining_searches", current_remaining)
                total = stats_live.get("total_searches", total)
                done_count = max(0, total - current_remaining) if total else executed
                userinfo = userinfo_live
                stats = stats_live
                # 根据最新数据动态抬升安全上限，防止 total 变化导致过早退出
                safety_cap = max(safety_cap, max(total, current_remaining, 30) * 3)
                self._render_dashboard(
                    account_name=account_name,
                    userinfo=userinfo_live,
                    stats=stats_live,
                    status_text=f"剩余 {current_remaining} 次",
                    search_index=None,
                    search_total=None,
                )
            except Exception:
                # 若实时刷新失败，则按已完成次数递减估算，确保循环能推进
                current_remaining = max(0, current_remaining - 1)
                self._render_dashboard(
                    account_name=account_name,
                    userinfo=userinfo,
                    stats=stats,
                    status_text=f"剩余 {current_remaining} 次",
                    search_index=None,
                    search_total=None,
                )

            if current_remaining <= 0:
                self.log_signal.emit(f"{prefix}剩余搜索已清零，任务完成。", "SUCCESS")
                break

            delay_ms = random.randint(min_ms, max_ms)
            delay_sec = delay_ms / 1000
            self.log_signal.emit(f"{prefix}等待 {delay_sec:.1f} 秒...", "INFO")
            step = 0.5
            slept = 0.0
            while slept < delay_sec:
                await asyncio.sleep(step)
                slept += step
                if self._should_stop(task_name, account_name):
                    break
            if self._should_stop(task_name, account_name):
                self.log_signal.emit(f"{prefix}任务已停止", "WARN")
                self._render_dashboard(
                    account_name=account_name,
                    userinfo=userinfo,
                    stats=stats,
                    status_text=f"已停止，剩余 {current_remaining} 次",
                    search_index=None,
                    search_total=None,
                )
                break

        # 结束前停止周期刷新任务
        try:
            refresh_task.cancel()
        except Exception:
            pass

        try:
            userinfo2 = await fetch_rewards_userinfo(context)
            stats2 = compute_remaining_searches(userinfo2, device_type)
            self.log_signal.emit(f"{prefix}最新进度: {stats2['current_points']} / {stats2['max_points']}", "INFO")
            self._render_dashboard(account_name=account_name, userinfo=userinfo2, stats=stats2, status_text="剩余 0 次", search_index=None, search_total=None)
        except Exception as e:
            self.log_signal.emit(f"{prefix}刷新最新进度失败: {e}", "WARN")

    def _should_stop(self, task_name: Optional[str], account_name: Optional[str]) -> bool:
        if task_name and self.task_stop_flags.get(task_name):
            return True
        if account_name and self.task_stop_flags.get(f"{account_name}_windows"):
            return True
        if account_name and self.task_stop_flags.get(f"{account_name}_iphone"):
            return True
        if self.task_stop_flags.get("__global"):
            return True
        return False

    def _reset_stop_flags_for_account(self, account_name: str):
        """启动任务前重置该账号的停止标记，避免上次停止后直接终止本次任务。"""
        for key in (f"{account_name}_windows", f"{account_name}_iphone"):
            if key in self.task_stop_flags and self.task_stop_flags[key]:
                self.task_stop_flags[key] = False

    def _context_kwargs(self, p, device_type: str, user_agent: str):
        if device_type == "windows":
            return {
                "user_agent": user_agent,
                "viewport": {"width": 1920, "height": 1080},
                "screen": {"width": 1920, "height": 1080},
                "locale": "zh-CN",
                # 与旧版体验一致，设置系统时区
                "timezone_id": "Asia/Shanghai",
            }
        # iPhone 模拟
        device_profile = None
        for name in ["iPhone 15", "iPhone 14", "iPhone 13", "iPhone 12"]:
            try:
                device_profile = p.devices[name]
                break
            except Exception:
                continue
        ctx = dict(device_profile or {})
        ctx["user_agent"] = user_agent
        ctx.setdefault("viewport", {"width": 390, "height": 844})
        ctx.setdefault("screen", {"width": 390, "height": 844})
        ctx["is_mobile"] = True
        ctx["has_touch"] = True
        ctx["device_scale_factor"] = 3
        ctx["locale"] = "zh-CN"
        ctx["timezone_id"] = "Asia/Shanghai"
        return ctx

    # 事件响应
    def _on_headless_changed(self, checked: bool):
        self.headless_mode = checked
        self.save_config()

    def _on_device_changed(self, checked: bool):  # noqa: ARG002
        self.device_windows = self.cb_windows.isChecked()
        self.device_iphone = self.cb_iphone.isChecked()
        self.save_config()

    # Qt 生命周期
    def closeEvent(self, event: QtGui.QCloseEvent):  # noqa: N802
        if self.is_running:
            res = QtWidgets.QMessageBox.question(
                self,
                "退出",
                "任务正在运行，确定要退出吗？",
                QtWidgets.QMessageBox.StandardButton.Ok,
                QtWidgets.QMessageBox.StandardButton.Cancel,
            )
            if res != QtWidgets.QMessageBox.StandardButton.Ok:
                event.ignore()
                return
            self.stop_all_tasks()
        event.accept()

    # 统一控制按钮状态，贴近旧版 GUI 的交互逻辑
    def _set_buttons_enabled(self, *, idle: bool):
        # idle=True 表示没有任务在运行
        self.btn_refresh.setEnabled(idle)
        self.btn_add_account.setEnabled(idle)
        self.btn_delete_account.setEnabled(idle)
        self.btn_start.setEnabled(idle)
        self.btn_batch.setEnabled(idle)
        # 停止按钮仅在非 idle 时可用
        self.btn_stop.setEnabled(not idle)
        self.btn_stop_all.setEnabled(not idle)


def main():
    app = QtWidgets.QApplication(sys.argv)
    # 设置应用级图标，确保标题栏与任务栏图标一致
    try:
        icon_candidates = [
            _get_exe_dir() / "Assets" / "ico.png",
            _get_bundle_dir() / "Assets" / "ico.png",
        ]
        for icon_path in icon_candidates:
            if icon_path.exists():
                app.setWindowIcon(QtGui.QIcon(str(icon_path)))
                break
    except Exception:
        pass
    window = RewardsGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()


