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
    _sanitize_filename,
    _get_exe_dir,
    _get_bundle_dir,
    _format_console_dashboard,
    MSEDGE_CHANNEL,
    DOMCONTENTLOADED,
)

# Playwright 异步 API（用于浏览器自动化）
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError


class DashboardWidget(QtWidgets.QFrame):
    """结构化仪表盘组件，使用卡片式布局显示等级、积分和进度"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            DashboardWidget {
                background-color: #1a1a1a;
                border: 1px solid #333;
                border-radius: 8px;
            }
        """)
        self._build_ui()
    
    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        # 顶部时间信息卡片
        time_card = self._create_card()
        time_layout = QtWidgets.QHBoxLayout(time_card)
        time_layout.setContentsMargins(10, 8, 10, 8)
        self.lbl_time = QtWidgets.QLabel("当前时间: --:--:--")
        self.lbl_time.setStyleSheet("font-size: 12px; color: #aaa;")
        time_layout.addWidget(self.lbl_time)
        time_layout.addStretch()
        layout.addWidget(time_card)
        
        # 账户信息卡片
        account_card = self._create_card()
        account_layout = QtWidgets.QGridLayout(account_card)
        account_layout.setContentsMargins(10, 10, 10, 10)
        account_layout.setSpacing(8)
        
        # 等级
        lbl_level_title = QtWidgets.QLabel("等级:")
        lbl_level_title.setStyleSheet("font-weight: bold; color: #e0e0e0;")
        self.lbl_level = QtWidgets.QLabel("--")
        self.lbl_level.setStyleSheet("font-size: 14px; color: #1a73e8;")
        
        # 总积分
        lbl_points_title = QtWidgets.QLabel("总积分:")
        lbl_points_title.setStyleSheet("font-weight: bold; color: #e0e0e0;")
        self.lbl_points = QtWidgets.QLabel("--")
        self.lbl_points.setStyleSheet("font-size: 16px; font-weight: bold; color: #34a853;")
        
        account_layout.addWidget(lbl_level_title, 0, 0)
        account_layout.addWidget(self.lbl_level, 0, 1)
        account_layout.addWidget(lbl_points_title, 0, 2)
        account_layout.addWidget(self.lbl_points, 0, 3)
        account_layout.setColumnStretch(1, 1)
        account_layout.setColumnStretch(3, 1)
        layout.addWidget(account_card)
        
        # 今日进度卡片
        progress_card = self._create_card()
        progress_layout = QtWidgets.QVBoxLayout(progress_card)
        progress_layout.setContentsMargins(10, 10, 10, 10)
        progress_layout.setSpacing(8)
        
        # 今日获取标题
        today_header = QtWidgets.QHBoxLayout()
        lbl_today_title = QtWidgets.QLabel("今日获取")
        lbl_today_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #e0e0e0;")
        self.lbl_today_points = QtWidgets.QLabel("-- / --")
        self.lbl_today_points.setStyleSheet("font-size: 13px; color: #aaa;")
        today_header.addWidget(lbl_today_title)
        today_header.addStretch()
        today_header.addWidget(self.lbl_today_points)
        progress_layout.addLayout(today_header)
        
        # 今日进度条
        self.progress_today = QtWidgets.QProgressBar()
        self.progress_today.setMinimum(0)
        self.progress_today.setMaximum(100)
        self.progress_today.setValue(0)
        self.progress_today.setTextVisible(True)
        self.progress_today.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 5px;
                background-color: #3d3d3d;
                height: 20px;
                text-align: center;
                color: #e0e0e0;
            }
            QProgressBar::chunk {
                background-color: #4285f4;
                border-radius: 4px;
            }
        """)
        progress_layout.addWidget(self.progress_today)
        
        layout.addWidget(progress_card)
        
        # 状态卡片
        status_card = self._create_card()
        status_layout = QtWidgets.QHBoxLayout(status_card)
        status_layout.setContentsMargins(10, 8, 10, 8)
        lbl_status_title = QtWidgets.QLabel("状态:")
        lbl_status_title.setStyleSheet("font-weight: bold; color: #e0e0e0;")
        self.lbl_status = QtWidgets.QLabel("等待中...")
        self.lbl_status.setStyleSheet("font-size: 13px; color: #1a73e8;")
        status_layout.addWidget(lbl_status_title)
        status_layout.addWidget(self.lbl_status, 1)
        layout.addWidget(status_card)
        
        layout.addStretch()
    
    def _create_card(self) -> QtWidgets.QFrame:
        """创建一个卡片样式的容器"""
        card = QtWidgets.QFrame()
        card.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        card.setStyleSheet("""
            QFrame {
                background-color: #2d2d2d;
                border: 1px solid #404040;
                border-radius: 6px;
            }
        """)
        return card
    
    def update_data(self, data: dict):
        """更新仪表盘显示数据"""
        if not data:
            self.clear()
            return
        
        # 时间信息
        self.lbl_time.setText(f"当前时间: {data.get('time', '--:--:--')}")
        
        # 账户信息
        self.lbl_level.setText(data.get('level', '--'))
        self.lbl_points.setText(str(data.get('total_points', '--')))
        
        # 今日进度
        pc_cur = data.get('pc_current', 0)
        pc_max = data.get('pc_max', 0)
        today_cur = pc_cur
        today_max = pc_max
        
        self.lbl_today_points.setText(f"{today_cur} / {today_max}")
        today_pct = int((today_cur / today_max * 100) if today_max > 0 else 0)
        self.progress_today.setValue(today_pct)
        self.progress_today.setFormat(f"{today_pct}%")
        
        # 状态
        status_text = data.get('status', '等待中...')
        self.lbl_status.setText(status_text)
    
    def clear(self):
        """清空仪表盘"""
        self.lbl_time.setText("当前时间: --:--:--")
        self.lbl_level.setText("--")
        self.lbl_points.setText("--")
        self.lbl_today_points.setText("-- / --")
        self.progress_today.setValue(0)
        self.progress_today.setFormat("0%")
        self.lbl_status.setText("等待中...")


class RewardsGUI(QtWidgets.QMainWindow):
    # 跨线程 UI 更新信号
    log_signal = QtCore.pyqtSignal(str, str)           # (message, level)
    status_signal = QtCore.pyqtSignal(str, str)        # (text, color)
    dashboard_signal = QtCore.pyqtSignal(dict, object)  # (dashboard data dict, account_name)
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
        # 各账号对应的最新仪表盘数据
        self.dashboard_texts: Dict[str, dict] = {}

        # 配置与设备选择
        self.config_path = Path("Assets/config.json")
        self.cookies_dir = Path("Assets/cookies")
        # old GUI 默认非无头运行，这里与旧版保持一致
        self.headless_mode = False

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
        opts.addWidget(self.cb_headless)
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

        # 使用结构化仪表盘组件替代 QTextEdit
        self.dashboard = DashboardWidget()
        self.dashboard.setMinimumHeight(280)
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
        self.dashboard_signal.connect(self._update_dashboard_data)
        self.refresh_accounts_signal.connect(self.refresh_accounts)
        
        # UI 控件事件
        self.cb_headless.toggled.connect(self._on_headless_changed)
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
            browser = await p.chromium.launch(channel=MSEDGE_CHANNEL, headless=False)
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
                except PlaywrightTimeoutError:
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
                self.cb_headless.setChecked(self.headless_mode)
        except Exception as e:
            self.log_signal.emit(f"读取配置失败: {e}", "ERROR")

    def save_config(self):
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            cfg = {
                "headless": self.headless_mode,
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
            # 统一计算/同步状态
            self._recompute_account_status(f.stem)
            status = self.account_status.get(f.stem) or "空闲"

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

    def _update_dashboard_data(self, data: dict, account_name: Optional[str] = None):
        # 记录每个账号对应的仪表盘数据
        if account_name:
            self.dashboard_texts[account_name] = data
        current = self._current_selected_account()
        if account_name is None or account_name == current:
            self.dashboard.update_data(data)

    def _show_dashboard_for_selected(self):
        name = self._current_selected_account()
        if not name:
            self.dashboard.clear()
            return
        data = self.dashboard_texts.get(name)
        if data:
            self.dashboard.update_data(data)
        else:
            self.dashboard.clear()

    def _on_account_selection_changed(self, _current, _previous):
        self._show_dashboard_for_selected()

    def _render_dashboard(self, *, account_name: Optional[str] = None, userinfo=None, stats=None, status_text: str = "", search_index=None, search_total=None):
        """构建并发送结构化仪表盘数据"""
        import time as _time
        
        # 若未能识别 firstName，使用 account_name 作为替代显示
        safe_userinfo = dict(userinfo or {})
        if account_name and not safe_userinfo.get("firstName"):
            safe_userinfo["firstName"] = account_name

        # 构建结构化数据
        data = {
            "time": _time.strftime("%H:%M:%S"),
            "level": "--",
            "total_points": 0,
            "pc_current": 0,
            "pc_max": 0,
            "status": status_text or "等待中...",
        }
        
        if safe_userinfo:
            dashboard_info = safe_userinfo.get("dashboard") or safe_userinfo
            user_status = (dashboard_info or {}).get("userStatus") or {}
            level_info = user_status.get("levelInfo") or {}
            data["level"] = level_info.get("activeLevel") or level_info.get("level") or "未知"
            data["total_points"] = int(user_status.get("availablePoints") or 0)
        
        if stats:
            data["pc_current"] = stats.get("pc", {}).get("current", 0)
            data["pc_max"] = stats.get("pc", {}).get("max", 0)

        self.dashboard_signal.emit(data, account_name)

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
    def start_task(self):
        item = self.accounts_list.currentItem()
        if not item:
            self.log_signal.emit("请先选择一个账号（cookies 文件）", "WARN")
            return
        cookie_path = Path(item.data(QtCore.Qt.ItemDataRole.UserRole))
        device_type = "windows"
        task_name = f"{cookie_path.stem}_{device_type}"
        
        # 检查该账号是否已在运行
        if task_name in self.running_tasks:
            self.log_signal.emit(f"账号 {cookie_path.stem} 已在运行中", "WARN")
            return
        
        self.selected_cookie = cookie_path
        # 重置该账号的停止标记，避免上一次停止后遗留的 True 影响本次启动
        self._reset_stop_flags_for_account(cookie_path.stem)
        self.is_running = True
        # 启动前清除全局停止标志，并更新按钮状态
        self.task_stop_flags.pop("__global", None)
        self._set_buttons_enabled(idle=False)
        self.task_stop_flags[task_name] = False
        self._launch_single(cookie_path, device_type, task_name)
        self.update_status("任务已启动", "green")

    def start_batch_tasks(self):
        files = [Path(self.accounts_list.item(i).data(QtCore.Qt.ItemDataRole.UserRole)) for i in range(self.accounts_list.count())]
        if not files:
            self.log_signal.emit("未找到任何账号 cookies", "WARN")
            return
        device_type = "windows"
        self.is_running = True
        self.task_stop_flags.pop("__global", None)
        self._set_buttons_enabled(idle=False)
        for f in files:
            # 批量模式下也先重置账号级停止标记，防止历史 stop 影响
            self._reset_stop_flags_for_account(f.stem)
            task_name = f"{f.stem}_{device_type}"
            self.task_stop_flags[task_name] = False
            self._launch_single(f, device_type, task_name)
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
        self._recompute_account_status(cookie_file.stem)
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
            self._recompute_account_status(cookie_file.stem)
            self.refresh_accounts_signal.emit()
            # 如果没有任何任务在运行，恢复按钮状态
            if not self.running_tasks:
                self.is_running = False
                self._set_buttons_enabled(idle=True)
                self.update_status("已停止或完成", "green")

    async def _execute_single_account(self, cookie_file: Path, device_type: str, task_name: str):
        prefix = f"[{cookie_file.stem}] "
        ua_value = USER_AGENTS[device_type]
        user_agent = ua_value() if callable(ua_value) else ua_value
        self.log_signal.emit(f"{prefix}启动浏览器，设备: {device_type}", "INFO")
        async with async_playwright() as p:
            launch_kwargs = {"channel": MSEDGE_CHANNEL, "headless": self.headless_mode}
            browser = await p.chromium.launch(**launch_kwargs)
            ctx_kwargs = self._context_kwargs(p, device_type, user_agent)
            context = await browser.new_context(**ctx_kwargs)
            # 加载 cookie JSON 并注入
            try:
                cookies = json.loads(Path(cookie_file).read_text(encoding="utf-8"))
                if isinstance(cookies, list) and cookies:
                    await context.add_cookies(cookies)
            except Exception:
                pass
            page = await context.new_page()
            
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
            await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)
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
            try:
                self.log_signal.emit(f"{prefix}搜索: {query} | 剩余 {current_remaining} 次", "INFO")
                self.status_signal.emit(f"搜索中，剩余 {current_remaining} 次", "orange")

                ok = False
                try:
                    ok = await asyncio.wait_for(_perform_bing_search_like_human(page, query), timeout=25)
                except asyncio.TimeoutError:
                    self.log_signal.emit(f"{prefix}搜索超时，重试...", "WARN")
                    try:
                        await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED, timeout=20000)
                    except Exception:
                        pass
                    ok = False
                except Exception:
                    ok = False

                if not ok:
                    await page.goto(search_url, wait_until=DOMCONTENTLOADED)
                await _maybe_human_scroll(page)
                await _maybe_click_one_result(page)

            except Exception as e:
                self.log_signal.emit(f"{prefix}搜索失败: {e}", "WARN")
                try:
                    await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)
                except Exception:
                    pass

            # 执行完一次搜索后，实时获取最新剩余次数，避免因积分延迟导致的提前结束
            try:
                userinfo_live = await fetch_rewards_userinfo(context)
                stats_live = compute_remaining_searches(userinfo_live, device_type)
                current_remaining = stats_live.get("remaining_searches", current_remaining)
                total = stats_live.get("total_searches", total)
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
        if self.task_stop_flags.get("__global"):
            return True
        return False

    def _reset_stop_flags_for_account(self, account_name: str):
        """启动任务前重置该账号的停止标记，避免上次停止后直接终止本次任务。"""
        key = f"{account_name}_windows"
        if key in self.task_stop_flags and self.task_stop_flags[key]:
            self.task_stop_flags[key] = False

    @staticmethod
    def _context_kwargs(p, device_type: str, user_agent: str):
        return {
            "user_agent": user_agent,
            "viewport": {"width": 1920, "height": 1080},
            "screen": {"width": 1920, "height": 1080},
            "locale": "zh-CN",
            # 与旧版体验一致，设置系统时区
            "timezone_id": "Asia/Shanghai",
        }

    # 事件响应
    def _on_headless_changed(self, _checked: bool):
        self.headless_mode = self.cb_headless.isChecked()
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
        self.btn_refresh.setEnabled(True)  # 刷新账号始终可用
        self.btn_add_account.setEnabled(True)  # 添加账号始终可用（运行时也可添加）
        self.btn_delete_account.setEnabled(idle)  # 删除账号仅空闲时可用
        self.btn_start.setEnabled(True)  # 开始按钮始终可用（可运行其他账号）
        self.btn_batch.setEnabled(idle)  # 批量开始仅空闲时可用
        # 停止按钮仅在非 idle 时可用
        self.btn_stop.setEnabled(not idle)
        self.btn_stop_all.setEnabled(not idle)

    def _recompute_account_status(self, account_name: str) -> None:
        """根据当前运行线程，更新指定账号的状态文本。"""
        running_devs = [
            name.split("_", 1)[1]
            for name in list(self.running_tasks.keys())
            if name.startswith(account_name + "_")
        ]
        order = {"windows": 0, "iphone": 1}
        running_devs = list(dict.fromkeys(running_devs))
        running_devs.sort(key=lambda d: order.get(d, 99))
        self.account_status[account_name] = f"运行中: {', '.join(running_devs)}" if running_devs else "空闲"


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


