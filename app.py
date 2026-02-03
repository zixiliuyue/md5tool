import concurrent.futures
import hashlib
import logging
from logging.handlers import RotatingFileHandler
import os
import sys
import threading
import time
from typing import Dict, Iterable, List, Optional, Set
from pathlib import Path

from PyQt6 import QtCore, QtGui, QtWidgets
from openpyxl import Workbook
from send2trash import send2trash

CHUNK_SIZE = 128 * 1024
COL_SELECT = 0
COL_PATH = 1
COL_SIZE = 2
COL_MD5 = 3
COL_GROUP = 4
COL_DURATION = 5
COL_STATUS = 6
TABLE_COLUMNS = 7
BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = (BASE_DIR / "logs").resolve()
LOG_FILE = LOG_DIR / "md5tool.log"
LOGGER = logging.getLogger("md5tool")


def setup_logging() -> None:
    """配置日志输出（自动滚动文件）。

    作用：避免日志文件无限增长，超过大小会自动分卷。
    """
    os.makedirs(LOG_DIR, exist_ok=True)
    LOGGER.setLevel(logging.INFO)
    if LOGGER.handlers:
        return
    handler = RotatingFileHandler(os.fspath(LOG_FILE), maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    handler.setFormatter(formatter)
    LOGGER.addHandler(handler)


def create_app_icon() -> QtGui.QIcon:
    """优先使用 img/image.png 作为应用图标。

    说明：如果图片不存在，则回退到内置绘制图标，避免程序崩溃。
    """
    image_path = BASE_DIR / "img" / "image.png"
    if image_path.exists():
        return QtGui.QIcon(os.fspath(image_path))

    # 回退：使用内置绘制图标
    size = 256
    pixmap = QtGui.QPixmap(size, size)
    pixmap.fill(QtCore.Qt.GlobalColor.transparent)

    painter = QtGui.QPainter(pixmap)
    painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)

    shadow_rect = QtCore.QRectF(18, 22, size - 36, size - 36)
    painter.setBrush(QtGui.QColor(0, 0, 0, 40))
    painter.setPen(QtCore.Qt.PenStyle.NoPen)
    painter.drawRoundedRect(shadow_rect, 48, 48)

    card_rect = QtCore.QRectF(12, 12, size - 24, size - 24)
    gradient = QtGui.QLinearGradient(card_rect.topLeft(), card_rect.bottomRight())
    gradient.setColorAt(0, QtGui.QColor("#6c5ce7"))
    gradient.setColorAt(0.5, QtGui.QColor("#00cec9"))
    gradient.setColorAt(1, QtGui.QColor("#74b9ff"))
    painter.setBrush(QtGui.QBrush(gradient))
    painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 160), 2))
    painter.drawRoundedRect(card_rect, 48, 48)

    highlight_rect = QtCore.QRectF(26, 24, size - 52, (size - 48) * 0.45)
    highlight = QtGui.QLinearGradient(highlight_rect.topLeft(), highlight_rect.bottomRight())
    highlight.setColorAt(0, QtGui.QColor(255, 255, 255, 120))
    highlight.setColorAt(1, QtGui.QColor(255, 255, 255, 10))
    painter.setBrush(QtGui.QBrush(highlight))
    painter.setPen(QtCore.Qt.PenStyle.NoPen)
    painter.drawRoundedRect(highlight_rect, 36, 36)

    badge_rect = QtCore.QRectF(size * 0.24, size * 0.24, size * 0.52, size * 0.52)
    badge_gradient = QtGui.QRadialGradient(badge_rect.center(), badge_rect.width() / 2)
    badge_gradient.setColorAt(0, QtGui.QColor("#ffffff"))
    badge_gradient.setColorAt(1, QtGui.QColor("#dfe6e9"))
    painter.setBrush(badge_gradient)
    painter.setPen(QtGui.QPen(QtGui.QColor(0, 0, 0, 60), 1))
    painter.drawEllipse(badge_rect)

    painter.setPen(QtGui.QPen(QtGui.QColor("#2d3436")))
    font = QtGui.QFont("Helvetica", int(size / 7.5), QtGui.QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(badge_rect, QtCore.Qt.AlignmentFlag.AlignCenter, "MD5")

    painter.end()
    return QtGui.QIcon(pixmap)


def iter_files(paths: Iterable[str]) -> List[str]:
    """把“文件/目录”输入展开为唯一文件列表。

    Demo:
        输入: ["/a/file1.txt", "/b/dir"]
        输出: ["/a/file1.txt", "/b/dir/x.png", "/b/dir/y.csv", ...]

    说明：
    - 目录会递归展开；
    - 路径会做绝对路径归一化；
    - 重复文件会去重。
    """
    seen: Set[str] = set()
    collected: List[str] = []
    input_paths = list(paths)
    for path in input_paths:
        if not path:
            continue
        # Normalize path first
        path = os.path.abspath(path)
        if os.path.isdir(path):
            LOGGER.info("Scanning directory: %s", path)
            for root, _, files in os.walk(path):
                for name in files:
                    full_path = os.path.abspath(os.path.join(root, name))
                    if full_path not in seen:
                        seen.add(full_path)
                        collected.append(full_path)
        elif os.path.isfile(path):
            if path not in seen:
                seen.add(path)
                collected.append(path)
        else:
            LOGGER.warning("Path does not exist or is not accessible: %s", path)
    LOGGER.info("Collected %d file(s) from %d path(s)", len(collected), len(input_paths))
    return collected


def format_size(size_value: Optional[int]) -> str:
    """把字节数转换成“人类可读”的大小字符串。

    Demo:
        512     -> "512 B"
        2048    -> "2.00 KB"
        1048576 -> "1.00 MB"
    """
    if size_value is None:
        return "N/A"
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(size_value)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.2f} {unit}" if unit != "B" else f"{int(size)} {unit}"
        size /= 1024
    return f"{size_value} B"


def format_duration(seconds: float) -> str:
    """把耗时（秒）转换成人类可读时间。

    Demo:
        0.032 -> "32 ms"
        0.8   -> "800 ms"
        3.21  -> "3.210 s"
        75    -> "1 m 15 s"
        3700  -> "1 h 1 m"
    """
    if seconds < 1:
        return f"{seconds * 1000:.0f} ms"
    if seconds < 60:
        return f"{seconds:.3f} s"
    minutes, sec = divmod(seconds, 60)
    if minutes < 60:
        return f"{int(minutes)} m {sec:.0f} s"
    hours, minutes = divmod(minutes, 60)
    return f"{int(hours)} h {int(minutes)} m"


def compute_md5(path: str, cancel: threading.Event) -> dict:
    """计算文件 MD5（支持取消）。

    说明：
    - 每读取一块数据就检查 cancel 事件；
    - 发生错误会返回 error 字段给 UI 显示。
    """
    start = time.perf_counter()
    size = 0
    digest = hashlib.md5()
    try:
        with open(path, "rb") as handle:
            while True:
                if cancel.is_set():
                    raise RuntimeError("cancelled")
                chunk = handle.read(CHUNK_SIZE)
                if not chunk:
                    break
                size += len(chunk)
                digest.update(chunk)
    except Exception as exc:  # pragma: no cover - GUI utility
        LOGGER.warning("Failed to hash %s: %s", path, exc)
        return {"path": path, "error": str(exc), "size": size, "duration": time.perf_counter() - start}
    duration = time.perf_counter() - start
    return {"path": path, "md5": digest.hexdigest(), "size": size, "duration": duration}


def safe_size(path: str) -> Optional[int]:
    """安全获取文件大小，失败返回 None。"""
    try:
        return os.path.getsize(path)
    except OSError as exc:  # pragma: no cover - GUI utility
        LOGGER.warning("Size check failed for %s: %s", path, exc)
        return None


class HashWorker(QtCore.QObject):
    """后台哈希工作器（线程池并发）。

    说明：
    - UI 线程只负责显示；
    - 实际计算放在后台线程池里，避免界面卡死。
    """
    progress = QtCore.pyqtSignal(int, int)
    result = QtCore.pyqtSignal(dict)
    finished = QtCore.pyqtSignal()

    def __init__(self, paths: List[str], workers: int):
        super().__init__()
        self.paths = paths
        self.workers = workers
        self._cancel = threading.Event()

    def cancel(self) -> None:
        """请求取消（设置事件标志）。"""
        self._cancel.set()

    def start(self) -> None:
        """启动后台线程（再由它创建线程池）。"""
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()

    def _run(self) -> None:
        """执行哈希任务，并逐个发出结果。

        Demo（事件流）:
            progress(0,total)
            result({path, md5, size, duration})
            progress(1,total)
            ...
            finished()
        """
        total = len(self.paths)
        self.progress.emit(0, total)
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.workers) as pool:
            futures = {pool.submit(compute_md5, path, self._cancel): path for path in self.paths}
            completed = 0
            for future in concurrent.futures.as_completed(futures):
                path = futures[future]
                if self._cancel.is_set():
                    break
                try:
                    data = future.result()
                except Exception as exc:  # pragma: no cover - GUI utility
                    data = {"path": path, "error": str(exc)}
                self.result.emit(data)
                completed += 1
                self.progress.emit(completed, total)
        self.finished.emit()


class MainWindow(QtWidgets.QMainWindow):
    """主窗口与 UI 控制器。"""
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MD5 Tool")
        self.setWindowIcon(create_app_icon())
        self.resize(900, 600)

        self.paths: Set[str] = set()
        self.path_items: Dict[str, QtWidgets.QTableWidgetItem] = {}
        self.worker: HashWorker | None = None
        self.md5_groups: Dict[str, Set[str]] = {}
        self.group_assignments: Dict[str, int] = {}
        self.group_colors: List[QtGui.QColor] = [
            QtGui.QColor("#f6e58d"),
            QtGui.QColor("#81ecec"),
            QtGui.QColor("#74b9ff"),
            QtGui.QColor("#fab1a0"),
            QtGui.QColor("#dfe6e9"),
            QtGui.QColor("#ffeaa7"),
            QtGui.QColor("#a29bfe"),
            QtGui.QColor("#b2bec3"),
        ]

        central = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        buttons_row = QtWidgets.QHBoxLayout()
        self.add_files_btn = QtWidgets.QPushButton("Add Files")
        self.add_dir_btn = QtWidgets.QPushButton("Add Folder")
        self.clear_btn = QtWidgets.QPushButton("Clear")
        self.export_btn = QtWidgets.QPushButton("Export")
        self.delete_selected_btn = QtWidgets.QPushButton("Delete Selected")
        self.exit_btn = QtWidgets.QPushButton("Exit")
        buttons_row.addWidget(self.add_files_btn)
        buttons_row.addWidget(self.add_dir_btn)
        buttons_row.addWidget(self.clear_btn)
        buttons_row.addWidget(self.export_btn)
        buttons_row.addWidget(self.delete_selected_btn)
        buttons_row.addWidget(self.exit_btn)
        buttons_row.addStretch()

        control_row = QtWidgets.QHBoxLayout()
        self.start_btn = QtWidgets.QPushButton("Start")
        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.cancel_btn.setEnabled(False)
        self.worker_label = QtWidgets.QLabel("")
        control_row.addWidget(self.start_btn)
        control_row.addWidget(self.cancel_btn)
        control_row.addStretch()
        control_row.addWidget(self.worker_label)

        self.table = QtWidgets.QTableWidget(0, TABLE_COLUMNS)
        self.table.setHorizontalHeaderLabels(["Select", "Path", "Size (bytes)", "MD5", "Group", "Duration (s)", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(COL_SELECT, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_PATH, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(COL_SIZE, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_MD5, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_GROUP, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_DURATION, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(COL_STATUS, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)

        self.progress = QtWidgets.QProgressBar()
        self.status = QtWidgets.QLabel("Ready")

        layout.addLayout(buttons_row)
        layout.addLayout(control_row)
        layout.addWidget(self.table)
        layout.addWidget(self.progress)
        layout.addWidget(self.status)
        central.setLayout(layout)
        self.setCentralWidget(central)

        self.add_files_btn.clicked.connect(self.add_files)
        self.add_dir_btn.clicked.connect(self.add_directory)
        self.clear_btn.clicked.connect(self.clear_paths)
        self.export_btn.clicked.connect(self.export_results)
        self.delete_selected_btn.clicked.connect(self.delete_selected)
        self.start_btn.clicked.connect(self.start_hashing)
        self.cancel_btn.clicked.connect(self.cancel_hashing)
        self.exit_btn.clicked.connect(QtWidgets.QApplication.quit)

        self.update_worker_label()
        self.set_hashing_state(False)

    def set_hashing_state(self, active: bool) -> None:
        """根据哈希状态启用/禁用按钮，避免中途改队列。"""
        self.start_btn.setEnabled(not active)
        self.add_files_btn.setEnabled(not active)
        self.add_dir_btn.setEnabled(not active)
        self.clear_btn.setEnabled(not active)
        self.cancel_btn.setEnabled(active)
        self.delete_selected_btn.setEnabled(not active)

    def is_hashing(self) -> bool:
        """当前是否处于哈希计算中。"""
        return self.worker is not None

    def update_worker_label(self) -> None:
        """更新并发线程数显示。"""
        workers = self.default_workers()
        self.worker_label.setText(f"Workers: {workers}")

    def add_files(self) -> None:
        """打开文件选择框并加入队列（支持多选）。"""
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cannot add new files.")
            return
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, 
            "Select files",
            "",
            "All Files (*.*)"
        )
        if files:
            LOGGER.info("User selected %d file(s)", len(files))
            self.append_paths(files)
        else:
            LOGGER.info("File selection cancelled")

    def add_directory(self) -> None:
        """打开目录选择框并递归加入所有文件。"""
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cannot add new folders.")
            return
        directory = QtWidgets.QFileDialog.getExistingDirectory(
            self, 
            "Select folder",
            "",
            QtWidgets.QFileDialog.Option.ShowDirsOnly
        )
        if directory:
            LOGGER.info("User selected directory: %s", directory)
            self.append_paths([directory])
        else:
            LOGGER.info("Directory selection cancelled")

    def append_paths(self, paths: Iterable[str]) -> None:
        """把路径添加到内部队列与表格中。

        说明：
        - 会自动去重；
        - 会在表格中插入新行。
        """
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cannot add new paths.")
            return
        sorting = self.table.isSortingEnabled()
        if sorting:
            self.table.setSortingEnabled(False)
        for path in iter_files(paths):
            if path not in self.paths:
                self.paths.add(path)
                self.add_row(path)
        if sorting:
            self.table.setSortingEnabled(True)
        self.status.setText(f"Queued {len(self.paths)} file(s)")
        LOGGER.info("Queued %d file(s)", len(self.paths))

    def add_row(self, path: str) -> None:
        """为指定文件路径插入一行。"""
        row = self.table.rowCount()
        self.table.insertRow(row)

        select_item = QtWidgets.QTableWidgetItem()
        select_item.setFlags(QtCore.Qt.ItemFlag.ItemIsUserCheckable | QtCore.Qt.ItemFlag.ItemIsEnabled)
        select_item.setCheckState(QtCore.Qt.CheckState.Unchecked)
        self.table.setItem(row, COL_SELECT, select_item)

        path_item = QtWidgets.QTableWidgetItem(path)
        self.table.setItem(row, COL_PATH, path_item)
        self.path_items[path] = path_item
        size_value = safe_size(path)
        self.table.setItem(row, COL_SIZE, self.make_size_item(size_value))
        for col in (COL_MD5, COL_GROUP, COL_DURATION, COL_STATUS):
            default_text = "Pending" if col == COL_STATUS else ""
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(default_text))

    def make_size_item(self, size_value: Optional[int]) -> QtWidgets.QTableWidgetItem:
        """创建文件大小单元格（人类可读）。"""
        item = QtWidgets.QTableWidgetItem(format_size(size_value))
        return item

    def clear_paths(self) -> None:
        """清空队列并重置 UI。"""
        if self.worker:
            return
        self.paths.clear()
        self.path_items.clear()
        self.md5_groups.clear()
        self.group_assignments.clear()
        self.table.setRowCount(0)
        self.progress.reset()
        self.status.setText("Cleared")
        LOGGER.info("Cleared queue")

    def start_hashing(self) -> None:
        """开始计算队列中所有文件的 MD5。"""
        if self.worker or not self.paths:
            return
        path_list = list(self.paths)
        workers = self.default_workers()
        self.worker = HashWorker(path_list, workers)
        self.worker.result.connect(self.on_result)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.progress.setMaximum(len(path_list))
        self.progress.setValue(0)
        self.status.setText("Running...")
        self.set_hashing_state(True)
        self.worker.start()
        LOGGER.info("Started hashing %d file(s) with %d workers", len(path_list), workers)

    def cancel_hashing(self) -> None:
        """取消计算（通过事件通知后台）。"""
        if self.worker:
            self.worker.cancel()
            self.status.setText("Cancelling...")
            LOGGER.info("Cancel requested")

    def on_progress(self, current: int, total: int) -> None:
        """更新进度条。"""
        self.progress.setMaximum(total)
        self.progress.setValue(current)

    def on_result(self, data: dict) -> None:
        """处理哈希结果并刷新表格行。

        Demo（成功结果）:
            data = {"path": "/a.txt", "md5": "...", "size": 123, "duration": 0.05}
            -> Size 更新为 "123 B"
            -> Duration 更新为 "50 ms"
            -> Status 更新为 "Done"

        Demo（失败结果）:
            data = {"path": "/a.txt", "error": "permission denied"}
            -> Status 更新为 "Error: permission denied"
        """
        path = data.get("path", "")
        item = self.path_items.get(path)
        if not item:
            LOGGER.warning("Result for unknown path: %s", path)
            return
        row = item.row()
        size_value = data.get("size")
        if size_value is not None:
            self.table.setItem(row, COL_SIZE, self.make_size_item(size_value))
        if "md5" in data:
            self.table.setItem(row, COL_MD5, QtWidgets.QTableWidgetItem(data["md5"]))
            duration_value = float(data.get("duration", 0.0))
            self.table.setItem(row, COL_DURATION, QtWidgets.QTableWidgetItem(format_duration(duration_value)))
            status_item = self.table.item(row, COL_STATUS)
            if status_item is None:
                status_item = QtWidgets.QTableWidgetItem()
                self.table.setItem(row, COL_STATUS, status_item)
            status_item.setText("Done")
            self.update_grouping(path, data["md5"])
            LOGGER.info("Hashed %s", path)
        else:
            self.table.setItem(row, COL_MD5, QtWidgets.QTableWidgetItem(""))
            self.table.setItem(row, COL_DURATION, QtWidgets.QTableWidgetItem(""))
            status_item = self.table.item(row, COL_STATUS)
            if status_item is None:
                status_item = QtWidgets.QTableWidgetItem()
                self.table.setItem(row, COL_STATUS, status_item)
            status_item.setText(f"Error: {data.get('error', 'unknown')}")
            LOGGER.warning("Failed %s: %s", path, data.get("error", "unknown"))

    def on_finished(self) -> None:
        """哈希完成后恢复 UI 状态。"""
        self.worker = None
        self.set_hashing_state(False)
        self.status.setText("Finished")
        LOGGER.info("Finished")

    def default_workers(self) -> int:
        """计算默认并发数。

        说明：CPU 核数 * 2（上限 32，下限 2）。
        """
        cpu = os.cpu_count() or 4
        return max(2, min(cpu * 2, 32))

    def group_color_for_index(self, index: int) -> QtGui.QColor:
        """根据分组编号返回背景色。"""
        if not self.group_colors:
            return QtGui.QColor("#dfe6e9")
        return self.group_colors[(index - 1) % len(self.group_colors)]

    def update_grouping(self, path: str, md5: str) -> None:
        """更新“重复文件”分组。

        说明：同一个 MD5 代表内容相同，会显示为同一组。
        """
        paths = self.md5_groups.setdefault(md5, set())
        paths.add(path)
        if md5 not in self.group_assignments:
            self.group_assignments[md5] = len(self.group_assignments) + 1
        self.refresh_group_visuals()

    def refresh_group_visuals(self) -> None:
        """刷新分组标签和背景色。

        Demo:
            当某 MD5 出现 2 个以上文件时，
            会显示 "Group 1" 并着色。
        """
        for md5, paths in self.md5_groups.items():
            label = ""
            brush = QtGui.QBrush()
            if len(paths) > 1:
                group_index = self.group_assignments.get(md5, 1)
                label = f"Group {group_index}"
                color = self.group_color_for_index(group_index)
                brush = QtGui.QBrush(color)
            for path in paths:
                item = self.path_items.get(path)
                if not item:
                    continue
                row = item.row()
                group_item = self.table.item(row, COL_GROUP)
                if group_item is None:
                    group_item = QtWidgets.QTableWidgetItem()
                    self.table.setItem(row, COL_GROUP, group_item)
                group_item.setText(label)
                for col in range(self.table.columnCount()):
                    cell = self.table.item(row, col)
                    if cell:
                        cell.setBackground(brush)
                        cell.setForeground(QtGui.QBrush(QtGui.QColor("black")))

    def delete_selected(self) -> None:
        """把选中的“已完成”文件移入回收站。"""
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cancel or wait before deleting.")
            return
        paths = self._selected_done_paths()
        if not paths:
            QtWidgets.QMessageBox.information(self, "No Selection", "Select completed rows to delete.")
            return
        if not self._confirm_delete(paths):
            return
        self._trash_paths(paths)

    def delete_all_done(self) -> None:
        """把所有“已完成”文件移入回收站。"""
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cancel or wait before deleting.")
            return
        paths = [self.table.item(row, COL_PATH).text() for row in range(self.table.rowCount())
             if self._is_done_row(row)]
        if not paths:
            QtWidgets.QMessageBox.information(self, "No Completed", "No completed entries to delete.")
            return
        if not self._confirm_delete(paths):
            return
        self._trash_paths(paths)

    def _selected_done_paths(self) -> List[str]:
        """返回“勾选 + 已完成”的文件路径。"""
        rows = []
        for row in range(self.table.rowCount()):
            select_item = self.table.item(row, COL_SELECT)
            if select_item and select_item.checkState() == QtCore.Qt.CheckState.Checked:
                rows.append(row)
        paths: List[str] = []
        for row in rows:
            if self._is_done_row(row):
                item = self.table.item(row, COL_PATH)
                if item:
                    paths.append(item.text())
        return paths

    def _is_done_row(self, row: int) -> bool:
        """判断某行是否已完成（状态为 Done）。"""
        status_item = self.table.item(row, COL_STATUS)
        return bool(status_item and status_item.text().startswith("Done"))

    def _confirm_delete(self, paths: List[str]) -> bool:
        """弹出删除确认对话框。"""
        msg = QtWidgets.QMessageBox(self)
        msg.setWindowIcon(create_app_icon())
        msg.setIcon(QtWidgets.QMessageBox.Icon.NoIcon)
        msg.setIconPixmap(create_app_icon().pixmap(64, 64))
        msg.setWindowTitle("Confirm Delete")
        preview = "\n".join(paths[:5])
        extra = "" if len(paths) <= 5 else f"\n...and {len(paths) - 5} more"
        msg.setText(f"Move selected file(s) to system trash?\n{preview}{extra}")
        msg.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.Cancel)
        return msg.exec() == QtWidgets.QMessageBox.StandardButton.Yes

    def _trash_paths(self, paths: List[str]) -> None:
        """把文件移入系统回收站并更新表格。"""
        errors: List[str] = []
        for path in paths:
            try:
                send2trash(path)
            except Exception as exc:  # pragma: no cover - GUI utility
                errors.append(f"{path}: {exc}")
        self._remove_paths_from_table(paths)
        if errors:
            QtWidgets.QMessageBox.warning(self, "Delete Issues", "\n".join(errors))
            for err in errors:
                LOGGER.warning("Delete failed: %s", err)
        else:
            QtWidgets.QMessageBox.information(self, "Deleted", "Selected file(s) moved to trash.")
        LOGGER.info("Deleted %d file(s) to trash", len(paths) - len(errors))

    def _remove_paths_from_table(self, paths: Iterable[str]) -> None:
        """从数据结构和表格中移除路径。"""
        paths_set = set(paths)
        # Remove from data maps first
        for path in paths_set:
            self.paths.discard(path)
            self.path_items.pop(path, None)
        # Update MD5 grouping
        for md5, path_set in list(self.md5_groups.items()):
            path_set.difference_update(paths_set)
            if not path_set:
                self.md5_groups.pop(md5, None)
                self.group_assignments.pop(md5, None)
        # Reassign group numbers to keep labels compact
        for idx, md5 in enumerate(sorted(self.md5_groups.keys()), start=1):
            self.group_assignments[md5] = idx
        # Remove table rows bottom-up to keep indexes valid
        rows_to_remove = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, COL_PATH)
            if item and item.text() in paths_set:
                rows_to_remove.append(row)
        for row in sorted(rows_to_remove, reverse=True):
            self.table.removeRow(row)
        self.refresh_group_visuals()

    def export_results(self) -> None:
        """导出表格内容到 Excel 文件。"""
        if self.is_hashing():
            QtWidgets.QMessageBox.information(self, "Busy", "Hashing in progress, cancel or wait to export results.")
            return
        if self.table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, "No Data", "Nothing to export yet.")
            return
        downloads = Path.home() / "Downloads"
        target_dir = downloads if downloads.exists() else BASE_DIR
        suggested = os.fspath(target_dir / "md5-results.xlsx")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Export Results", suggested, "Excel Files (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "MD5 Results"
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            sheet.append(headers)
            for row in range(self.table.rowCount()):
                values: List[str] = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    values.append(item.text() if item else "")
                sheet.append(values)
            workbook.save(path)
            QtWidgets.QMessageBox.information(self, "Exported", f"Saved results to:\n{path}")
            LOGGER.info("Exported results to %s", path)
        except Exception as exc:  # pragma: no cover - GUI utility
            QtWidgets.QMessageBox.warning(self, "Export Failed", f"Could not export results: {exc}")
            LOGGER.exception("Export failed: %s", exc)


def main() -> None:
    """程序入口。"""
    setup_logging()
    LOGGER.info("App started")
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(create_app_icon())
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
