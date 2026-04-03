import sys
import socket
import requests
import re
import threading
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QLineEdit, 
                             QLabel, QFileDialog, QProgressBar, QTableWidget,
                             QTableWidgetItem, QHeaderView, QSplitter, QGroupBox,
                             QSpinBox, QComboBox, QCheckBox, QMessageBox, QMenu,
                             QAction, QShortcut)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMutex
from PyQt5.QtGui import QFont, QColor, QTextCursor, QKeySequence, QClipboard
import pandas as pd
from openpyxl.styles import Alignment

# 线程安全打印锁
print_lock = threading.Lock()

# 定义占位符
PLACEHOLDER = "*"

def format_size(size):
    """格式化文件大小显示"""
    if size == PLACEHOLDER or not isinstance(size, int):
        return PLACEHOLDER
    
    if size < 1024:
        return f"{size}B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f}KB"
    elif size < 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024):.1f}MB"
    else:
        return f"{size / (1024 * 1024 * 1024):.1f}GB"

def resolve_host(host):
    """解析主机名为IP地址"""
    try:
        socket.inet_aton(host)
        return host
    except socket.error:
        try:
            return socket.gethostbyname(host)
        except socket.gaierror:
            return None

def query_ip_location(ip, session):
    """通过cip.cc查询IP地理位置，提取"数据三"字段内容"""
    try:
        response = session.get(f"http://cip.cc/{ip}", timeout=5)
        response.encoding = 'utf-8'
        content = response.text
        
        # 提取"数据三"字段
        match = re.search(r'数据三\s*:\s*(.+?)(?:\n|$)', content, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        match = re.search(r'数据二\s*:\s*(.+?)(?:\n|$)', content, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        match = re.search(r'地址\s*:\s*(.+?)(?:\n|$)', content, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        return PLACEHOLDER
    except:
        return PLACEHOLDER

def get_url_headers_info(url, session):
    """获取URL的HTTP响应头信息"""
    try:
        response = session.head(url, timeout=5, allow_redirects=True)
        status_code = response.status_code
        server_header = response.headers.get('Server', PLACEHOLDER)
        if server_header == '':
            server_header = PLACEHOLDER
        
        content_length = response.headers.get('Content-Length', PLACEHOLDER)
        if content_length == '':
            content_length = PLACEHOLDER
        
        if content_length != PLACEHOLDER and content_length.isdigit():
            content_length = int(content_length)
            
        return status_code, server_header, content_length
    except:
        return PLACEHOLDER, PLACEHOLDER, PLACEHOLDER

def get_url_title(url, session, timeout=10):
    """获取URL页面的<title>标签内容"""
    try:
        response = session.get(url, timeout=timeout, allow_redirects=True, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        if response.status_code != 200:
            return PLACEHOLDER
        
        try:
            soup = BeautifulSoup(response.content, 'html.parser')
            title_tag = soup.find('title')
            if title_tag and title_tag.string:
                return title_tag.string.strip()
        except:
            pass
        
        try:
            response.encoding = response.apparent_encoding or 'utf-8'
            content = response.text
            title_match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE | re.DOTALL)
            if title_match:
                title = title_match.group(1).strip()
                title = re.sub(r'&[a-z]+;', '', title)
                if title:
                    return title
        except:
            pass
        
        return PLACEHOLDER
    except:
        return PLACEHOLDER

class WorkerThread(QThread):
    """工作线程类"""
    update_progress = pyqtSignal(int, int)  # 当前进度，总数
    update_result = pyqtSignal(int, dict)   # 行号，结果数据
    update_log = pyqtSignal(str)            # 日志信息
    finished_all = pyqtSignal()
    
    def __init__(self, urls, threads, method, timeout):
        super().__init__()
        self.urls = urls
        self.threads = threads
        self.method = method
        self.timeout = timeout
        self.is_running = True
        self.mutex = QMutex()
    
    def stop(self):
        """停止线程"""
        self.mutex.lock()
        self.is_running = False
        self.mutex.unlock()
    
    def create_session(self):
        """创建会话对象"""
        session = requests.Session()
        retry_strategy = Retry(
            total=2,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(
            max_retries=retry_strategy,
            pool_connections=50,
            pool_maxsize=50
        )
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        return session
    
    def process_url(self, idx, url, session):
        """处理单个URL"""
        # 提取主机名并获取物理地址
        parsed = urlparse(url)
        host = parsed.hostname
        if not host:
            location = PLACEHOLDER
        else:
            ip = resolve_host(host)
            if not ip:
                location = PLACEHOLDER
            else:
                location = query_ip_location(ip, session)
        
        # 获取页面标题
        title = get_url_title(url, session, self.timeout)
        
        # 获取HTTP响应头信息
        if self.method == 'HEAD':
            status_code, server_header, content_length = get_url_headers_info(url, session)
        else:
            # GET方法
            try:
                response = session.get(url, timeout=5, stream=True, allow_redirects=True)
                status_code = response.status_code
                server_header = response.headers.get('Server', PLACEHOLDER)
                if server_header == '':
                    server_header = PLACEHOLDER
                
                content_length = response.headers.get('Content-Length', PLACEHOLDER)
                if content_length == '':
                    content_length = PLACEHOLDER
                if content_length != PLACEHOLDER and content_length.isdigit():
                    content_length = int(content_length)
                
                response.close()
            except:
                status_code, server_header, content_length = PLACEHOLDER, PLACEHOLDER, PLACEHOLDER
        
        # 格式化Content-Length
        if content_length != PLACEHOLDER and isinstance(content_length, int):
            content_length_display = f"{content_length}字节 ({format_size(content_length)})"
        else:
            content_length_display = PLACEHOLDER
        
        return {
            'url': url,
            'location': location,
            'title': title,
            'status_code': status_code,
            'server': server_header,
            'content_length': content_length_display
        }
    
    def run(self):
        """线程运行函数"""
        import concurrent.futures
        
        total = len(self.urls)
        results = [None] * total
        completed = 0
        
        # 创建线程池
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.threads) as executor:
            futures = {}
            
            for idx, url in enumerate(self.urls):
                if not self.is_running:
                    break
                
                session = self.create_session()
                future = executor.submit(self.process_url, idx, url, session)
                futures[future] = (idx, session)
            
            for future in concurrent.futures.as_completed(futures):
                if not self.is_running:
                    break
                
                idx, session = futures[future]
                try:
                    result = future.result()
                    results[idx] = result
                    self.update_result.emit(idx, result)
                except Exception as e:
                    self.update_log.emit(f"处理URL时发生错误: {e}")
                    results[idx] = {
                        'url': self.urls[idx],
                        'location': PLACEHOLDER,
                        'title': PLACEHOLDER,
                        'status_code': PLACEHOLDER,
                        'server': PLACEHOLDER,
                        'content_length': PLACEHOLDER
                    }
                finally:
                    session.close()
                    completed += 1
                    self.update_progress.emit(completed, total)
        
        self.finished_all.emit()

class CustomTextEdit(QTextEdit):
    """自定义文本编辑框，自动调整高度"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFont(QFont("Consolas", 10))
        self.setLineWrapMode(QTextEdit.NoWrap)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.document().contentsChanged.connect(self.adjust_height)
        
    def adjust_height(self):
        """自动调整高度以适应内容"""
        doc_height = self.document().size().height()
        # 设置最小高度和最大高度限制
        min_height = 100
        max_height = 800
        new_height = max(min_height, min(doc_height + 20, max_height))
        self.setMinimumHeight(new_height)
        self.setMaximumHeight(new_height)

class LogTextEdit(QTextEdit):
    """日志文本框，支持自动滚动和内容管理"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFont(QFont("Consolas", 10))
        self.setLineWrapMode(QTextEdit.NoWrap)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
    def append_log(self, text):
        """添加日志并自动滚动"""
        self.append(text)
        # 自动滚动到底部
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.setTextCursor(cursor)
        
    def clear_log(self):
        """清空日志"""
        self.clear()

class TableWidget(QTableWidget):
    """自定义表格组件，支持右键菜单和快捷键复制"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
        # 添加快捷键 Ctrl+C 复制
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self)
        self.copy_shortcut.activated.connect(self.copy_selected)
        
        # 添加快捷键 Ctrl+A 全选
        self.select_all_shortcut = QShortcut(QKeySequence.SelectAll, self)
        self.select_all_shortcut.activated.connect(self.select_all_cells)
        
        # 获取主窗口引用用于显示状态栏消息
        self.main_window = parent
        
    def show_context_menu(self, position):
        """显示右键菜单"""
        menu = QMenu()
        
        copy_action = QAction("复制选中内容", self)
        copy_action.triggered.connect(self.copy_selected)
        menu.addAction(copy_action)
        
        copy_url_action = QAction("复制URL地址", self)
        copy_url_action.triggered.connect(self.copy_urls)
        menu.addAction(copy_url_action)
        
        menu.addSeparator()
        
        select_all_action = QAction("全选", self)
        select_all_action.triggered.connect(self.select_all_cells)
        menu.addAction(select_all_action)
        
        menu.exec_(self.viewport().mapToGlobal(position))
    
    def copy_selected(self):
        """复制选中的内容"""
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            # 如果没有选中任何内容，提示用户
            if self.main_window:
                self.main_window.statusBar().showMessage("请先选中要复制的内容", 2000)
            return
        
        clipboard = QApplication.clipboard()
        text = ""
        
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                row_texts = []
                for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    item = self.item(row, col)
                    if item:
                        row_texts.append(item.text())
                    else:
                        row_texts.append("")
                text += "\t".join(row_texts) + "\n"
        
        clipboard.setText(text.strip())
        # 使用状态栏显示成功信息，不弹出弹窗
        if self.main_window:
            count = len(selected_ranges)
            self.main_window.statusBar().showMessage(f"✓ 已复制 {count} 个区域的内容到剪贴板", 2000)
    
    def copy_urls(self):
        """复制URL地址"""
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            # 如果没有选中，复制所有URL
            urls = []
            for row in range(self.rowCount()):
                item = self.item(row, 0)  # URL地址在第0列
                if item and item.text():
                    urls.append(item.text())
        else:
            # 复制选中的URL
            urls = []
            for selected_range in selected_ranges:
                for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                    item = self.item(row, 0)  # URL地址在第0列
                    if item and item.text():
                        urls.append(item.text())
        
        if urls:
            clipboard = QApplication.clipboard()
            clipboard.setText("\n".join(urls))
            # 使用状态栏显示成功信息，不弹出弹窗
            if self.main_window:
                self.main_window.statusBar().showMessage(f"✓ 已复制 {len(urls)} 个URL地址到剪贴板", 2000)
        else:
            if self.main_window:
                self.main_window.statusBar().showMessage("没有可复制的URL地址", 2000)
    
    def select_all_cells(self):
        """全选所有单元格"""
        self.selectAll()
    
    def keyPressEvent(self, event):
        """键盘事件处理"""
        if event.key() == Qt.Key_A and event.modifiers() == Qt.ControlModifier:
            # Ctrl+A 全选
            self.select_all_cells()
            event.accept()
        else:
            super().keyPressEvent(event)

class MainWindow(QMainWindow):
    """主窗口类"""
    def __init__(self):
        super().__init__()
        self.urls = []
        self.results = []
        self.worker_thread = None
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("F4a1o Hacking Tools - URL信息提取工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建状态栏
        self.statusBar().showMessage("就绪")
        
        # 设置样式表
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
            }
            QLabel {
                color: #ffffff;
                font-size: 12px;
            }
            QPushButton {
                background-color: #4a4a4a;
                color: #ffffff;
                border: 1px solid #5a5a5a;
                padding: 5px 10px;
                border-radius: 3px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #5a5a5a;
            }
            QPushButton:pressed {
                background-color: #3a3a3a;
            }
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: Consolas, monospace;
                font-size: 11px;
                border: 1px solid #3a3a3a;
            }
            QTableWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                gridline-color: #3a3a3a;
                font-size: 11px;
                selection-background-color: #3a3a3a;
                selection-color: #ffffff;
            }
            QTableWidget::item {
                padding: 3px;
            }
            QTableWidget::item:selected {
                background-color: #3a3a3a;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #3a3a3a;
                color: #ffffff;
                padding: 5px;
                border: 1px solid #4a4a4a;
            }
            QGroupBox {
                color: #ffffff;
                border: 1px solid #4a4a4a;
                border-radius: 5px;
                margin-top: 10px;
                font-size: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QSpinBox, QComboBox, QLineEdit {
                background-color: #1e1e1e;
                color: #ffffff;
                border: 1px solid #4a4a4a;
                padding: 3px;
                border-radius: 3px;
            }
            QProgressBar {
                border: 1px solid #4a4a4a;
                border-radius: 3px;
                text-align: center;
                color: #ffffff;
            }
            QProgressBar::chunk {
                background-color: #4caf50;
                border-radius: 2px;
            }
            QStatusBar {
                color: #ffffff;
                background-color: #1e1e1e;
            }
        """)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建分割器
        splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(splitter)
        
        # 顶部控制面板
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QHBoxLayout()
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择URL列表文件...")
        self.file_path_edit.setReadOnly(True)
        file_layout.addWidget(self.file_path_edit)
        
        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.select_file_btn)
        
        file_group.setLayout(file_layout)
        top_layout.addWidget(file_group)
        
        # 配置区域
        config_group = QGroupBox("配置选项")
        config_layout = QHBoxLayout()
        
        config_layout.addWidget(QLabel("线程数:"))
        self.threads_spin = QSpinBox()
        self.threads_spin.setRange(1, 50)
        self.threads_spin.setValue(20)
        config_layout.addWidget(self.threads_spin)
        
        config_layout.addWidget(QLabel("HTTP方法:"))
        self.method_combo = QComboBox()
        self.method_combo.addItems(["HEAD", "GET"])
        config_layout.addWidget(self.method_combo)
        
        config_layout.addWidget(QLabel("超时时间(秒):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(5, 30)
        self.timeout_spin.setValue(10)
        config_layout.addWidget(self.timeout_spin)
        
        self.save_excel_btn = QPushButton("保存到Excel")
        self.save_excel_btn.clicked.connect(self.save_to_excel)
        self.save_excel_btn.setEnabled(False)
        config_layout.addWidget(self.save_excel_btn)
        
        config_layout.addStretch()
        config_group.setLayout(config_layout)
        top_layout.addWidget(config_group)
        
        # 控制按钮区域
        control_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始检测")
        self.start_btn.clicked.connect(self.start_detection)
        self.start_btn.setStyleSheet("background-color: #4caf50;")
        control_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("停止检测")
        self.stop_btn.clicked.connect(self.stop_detection)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("background-color: #f44336;")
        control_layout.addWidget(self.stop_btn)
        
        self.clear_btn = QPushButton("清空结果")
        self.clear_btn.clicked.connect(self.clear_results)
        control_layout.addWidget(self.clear_btn)
        
        control_layout.addStretch()
        top_layout.addLayout(control_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        top_layout.addWidget(self.progress_bar)
        
        splitter.addWidget(top_widget)
        
        # 底部结果显示区域 - 使用水平布局
        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(5)
        
        # 左侧表格 - 占3/5宽度
        self.table_widget = TableWidget(self)
        self.table_widget.setColumnCount(6)
        self.table_widget.setHorizontalHeaderLabels(['URL地址', '物理地址', '页面标题', '状态码', 'Server字段', 'Content-Length'])
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #252525;
                background-color: #1e1e1e;
            }
            QTableWidget::item {
                color: #ffffff;
            }
            QTableWidget::item:selected {
                background-color: #3a3a3a;
                color: #ffffff;
            }
        """)
        bottom_layout.addWidget(self.table_widget, 3)
        
        # 右侧日志区域 - 占2/5宽度，使用自定义日志文本框
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        self.log_text = LogTextEdit()
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        bottom_layout.addWidget(log_group, 2)
        
        splitter.addWidget(bottom_widget)
        splitter.setSizes([300, 500])
        
        # 添加启动横幅
        self.add_banner()
    
    def add_banner(self):
        """添加启动横幅 - F4a1o（完整显示）"""
        banner = (
            "╔════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗\n"
            "║                                                                                                                        ║\n"
            "║     ███████╗ █████╗  █████╗ ██╗      ██████╗                                                                          ║\n"
            "║     ██╔════╝██╔══██╗██╔══██╗██║     ██╔═══██╗                                                                         ║\n"
            "║     █████╗  ███████║███████║██║     ██║   ██║                                                                         ║\n"
            "║     ██╔══╝  ██╔══██║██╔══██║██║     ██║   ██║                                                                         ║\n"
            "║     ██║     ██║  ██║██║  ██║███████╗╚██████╔╝                                                                         ║\n"
            "║     ╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝ ╚═════╝                                                                          ║\n"
            "║                                                                                                                        ║\n"
            "║                           ⚡ F4a1o Hacking Tools ⚡                                                                    ║\n"
            "║                              Author: 法老                                                                              ║\n"
            "║                                                                                                                        ║\n"
            "╚════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
        )
        self.log_text.append_log(banner)
        self.log_text.append_log("\n工具已启动，请选择URL文件开始检测...\n")
        self.log_text.append_log("💡 提示：在结果表格中点击右键可以复制URL地址或选中内容\n")
        self.log_text.append_log("💡 快捷键：Ctrl+C 复制选中内容 | Ctrl+A 全选\n")
        # 滚动到顶部
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.Start)
        self.log_text.setTextCursor(cursor)
    
    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择URL列表文件", "", "文本文件 (*.txt);;所有文件 (*)")
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_urls(file_path)
    
    def load_urls(self, file_path):
        """加载URL列表"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.urls = [line.strip() for line in f if line.strip()]
            
            self.log_text.append_log(f"✓ 成功加载 {len(self.urls)} 条URL")
            self.start_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件失败: {e}")
            self.log_text.append_log(f"✗ 读取文件失败: {e}")
    
    def start_detection(self):
        """开始检测"""
        if not self.urls:
            QMessageBox.warning(self, "警告", "请先选择URL文件")
            return
        
        # 清空现有结果
        self.clear_results()
        
        # 禁用控件
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_file_btn.setEnabled(False)
        self.save_excel_btn.setEnabled(False)
        
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self.urls))
        self.progress_bar.setValue(0)
        
        # 创建并启动工作线程
        self.worker_thread = WorkerThread(
            self.urls,
            self.threads_spin.value(),
            self.method_combo.currentText(),
            self.timeout_spin.value()
        )
        self.worker_thread.update_progress.connect(self.update_progress)
        self.worker_thread.update_result.connect(self.update_result)
        self.worker_thread.update_log.connect(self.update_log)
        self.worker_thread.finished_all.connect(self.detection_finished)
        self.worker_thread.start()
        
        self.log_text.append_log("🚀 开始检测URL...")
    
    def stop_detection(self):
        """停止检测"""
        if self.worker_thread:
            self.worker_thread.stop()
            self.log_text.append_log("⚠ 用户停止检测")
            self.detection_finished()
    
    def update_progress(self, current, total):
        """更新进度条"""
        self.progress_bar.setValue(current)
        self.progress_bar.setFormat(f"处理进度: {current}/{total} ({current*100//total}%)")
    
    def update_result(self, idx, result):
        """更新结果表格 - 不高亮显示"""
        row = self.table_widget.rowCount()
        self.table_widget.insertRow(row)
        
        # 设置单元格内容，不设置任何背景色
        url_item = QTableWidgetItem(result['url'])
        url_item.setFlags(url_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 0, url_item)
        
        location_item = QTableWidgetItem(result['location'])
        location_item.setFlags(location_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 1, location_item)
        
        title_item = QTableWidgetItem(result['title'])
        title_item.setFlags(title_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 2, title_item)
        
        status_item = QTableWidgetItem(str(result['status_code']))
        status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 3, status_item)
        
        server_item = QTableWidgetItem(result['server'])
        server_item.setFlags(server_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 4, server_item)
        
        length_item = QTableWidgetItem(result['content_length'])
        length_item.setFlags(length_item.flags() & ~Qt.ItemIsEditable)
        self.table_widget.setItem(row, 5, length_item)
        
        # 保存结果
        self.results.append(result)
        
        # 实时更新日志
        self.log_text.append_log(f"✓ [{row+1}] {result['url'][:50]}... -> {result['location']}")
    
    def update_log(self, message):
        """更新日志"""
        self.log_text.append_log(message)
    
    def detection_finished(self):
        """检测完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_file_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if self.results:
            self.save_excel_btn.setEnabled(True)
            
            success_count = sum(1 for r in self.results if r['location'] != PLACEHOLDER)
            title_success = sum(1 for r in self.results if r['title'] != PLACEHOLDER)
            
            self.log_text.append_log(f"\n{'='*60}")
            self.log_text.append_log(f"✅ 检测完成！")
            self.log_text.append_log(f"📊 统计信息:")
            self.log_text.append_log(f"   • 成功获取物理地址: {success_count}/{len(self.urls)} ({success_count*100//len(self.urls)}%)")
            self.log_text.append_log(f"   • 成功获取页面标题: {title_success}/{len(self.urls)} ({title_success*100//len(self.urls)}%)")
            self.log_text.append_log(f"{'='*60}")
            self.log_text.append_log(f"✨ 工具执行完成！感谢使用 F4a1o Hacking Tools | Author: 法老")
    
    def save_to_excel(self):
        """保存到Excel"""
        if not self.results:
            QMessageBox.warning(self, "警告", "没有结果可保存")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(self, "保存Excel文件", "", "Excel文件 (*.xlsx)")
        if not file_path:
            return
        
        try:
            # 准备数据
            data = []
            for result in self.results:
                data.append([
                    result['url'],
                    result['location'],
                    result['title'],
                    result['status_code'],
                    result['server'],
                    result['content_length']
                ])
            
            # 创建DataFrame
            df = pd.DataFrame(data, columns=['URL地址', '物理地址', '页面标题', '状态码', 'Server字段', 'Content-Length'])
            
            # 保存Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='URL信息')
                
                # 设置对齐和列宽
                workbook = writer.book
                worksheet = writer.sheets['URL信息']
                
                from openpyxl.styles import Alignment
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                
                worksheet.column_dimensions['A'].width = 50
                worksheet.column_dimensions['B'].width = 35
                worksheet.column_dimensions['C'].width = 40
                worksheet.column_dimensions['D'].width = 10
                worksheet.column_dimensions['E'].width = 20
                worksheet.column_dimensions['F'].width = 25
            
            self.log_text.append_log(f"✓ 结果已保存至: {file_path}")
            self.statusBar().showMessage(f"✓ 结果已保存至: {file_path}", 3000)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存Excel失败: {e}")
            self.log_text.append_log(f"✗ 保存Excel失败: {e}")
    
    def clear_results(self):
        """清空结果"""
        self.table_widget.setRowCount(0)
        self.results.clear()
        self.save_excel_btn.setEnabled(False)

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()