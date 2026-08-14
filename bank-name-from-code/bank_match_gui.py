"""联行号匹配收款行的 PyQt5 图形界面。"""

from __future__ import annotations

import sys
import time
import traceback
import os
from datetime import datetime
from pathlib import Path

try:
    from PyQt5.QtCore import QSettings, QThread, QTimer, Qt, pyqtSignal
    from PyQt5.QtWidgets import (
        QApplication,
        QCheckBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QProgressBar,
        QPushButton,
        QScrollArea,
        QStyle,
        QToolButton,
        QVBoxLayout,
        QWidget,
    )
except ImportError:
    QApplication = None

from match_bank_names import (
    MatchingCancelled,
    process_target_file,
    read_master_data,
    read_target_prefixes,
)


APPLICATION_DIRECTORY = (
    Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
)
RUNTIME_LOG_PATH = APPLICATION_DIRECTORY / "gui_runtime.log"


def write_runtime_event(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with RUNTIME_LOG_PATH.open("a", encoding="utf-8") as log_file:
            log_file.write(f"{timestamp} | {message}\n")
    except OSError:
        pass


if QApplication is not None:

    class MatchWorker(QThread):
        stage_changed = pyqtSignal(str)
        progressed = pyqtSignal(int, int, dict)
        completed = pyqtSignal(dict, str, int)
        cancelled = pyqtSignal()
        failed = pyqtSignal(str)

        def __init__(
            self,
            master_path: Path,
            target_path: Path,
            output_path: Path,
            target_sheet_names: set[str] | None,
        ) -> None:
            super().__init__()
            self.master_path = master_path
            self.target_path = target_path
            self.output_path = output_path
            self.target_sheet_names = target_sheet_names

        def run(self) -> None:
            try:
                self.stage_changed.emit("正在分析待匹配文件...")
                prefixes = read_target_prefixes(self.target_path, self.target_sheet_names)
                if self.isInterruptionRequested():
                    raise MatchingCancelled()

                self.stage_changed.emit(f"正在读取 {len(prefixes)} 个主数据类别...")
                unique_matches, duplicate_matches = read_master_data(
                    self.master_path, prefixes
                )
                if self.isInterruptionRequested():
                    raise MatchingCancelled()

                self.stage_changed.emit("正在匹配记录...")
                summary = process_target_file(
                    self.target_path,
                    self.output_path,
                    unique_matches,
                    duplicate_matches,
                    "匹配收款行名称",
                    "匹配状态",
                    progress_callback=self.report_progress,
                    cancel_callback=self.isInterruptionRequested,
                    target_sheet_names=self.target_sheet_names,
                )
                self.completed.emit(summary, str(self.output_path), len(prefixes))
            except MatchingCancelled:
                self.cancelled.emit()
            except Exception as error:
                self.failed.emit(str(error))

        def report_progress(self, processed: int, total: int, summary: dict[str, int]) -> None:
            self.progressed.emit(processed, total, summary)


    class FileField(QWidget):
        path_changed = pyqtSignal()

        def __init__(self, *, save_file: bool = False) -> None:
            super().__init__()
            self.save_file = save_file
            self.path_input = QLineEdit()
            self.path_input.setClearButtonEnabled(True)
            self.path_input.setPlaceholderText("请选择文件")
            self.path_input.editingFinished.connect(self.path_changed.emit)

            self.browse_button = QToolButton()
            self.browse_button.setIcon(
                self.style().standardIcon(
                    QStyle.SP_DialogSaveButton
                    if save_file
                    else QStyle.SP_DirOpenIcon
                )
            )
            self.browse_button.setToolTip("选择结果文件" if save_file else "选择文件")
            self.browse_button.clicked.connect(self.select_file)

            layout = QHBoxLayout(self)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(6)
            layout.addWidget(self.path_input, 1)
            layout.addWidget(self.browse_button)

        def value(self) -> Path | None:
            text = self.path_input.text().strip()
            return Path(text) if text else None

        def set_value(self, path: Path | str | None) -> None:
            self.path_input.setText("" if not path else str(path))

        def select_file(self) -> None:
            current = self.value()
            start_dir = Path.home()
            if current and current.parent.is_dir():
                start_dir = current.parent

            file_filter = "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*.*)"
            if self.save_file:
                file_name, _ = QFileDialog.getSaveFileName(
                    self, "选择结果文件", str(start_dir), file_filter
                )
                if file_name:
                    selected = Path(file_name)
                    self.set_value(selected if selected.suffix else selected.with_suffix(".xlsx"))
                    self.path_changed.emit()
            else:
                file_name, _ = QFileDialog.getOpenFileName(
                    self, "选择 Excel 文件", str(start_dir), file_filter
                )
                if file_name:
                    self.set_value(Path(file_name))
                    self.path_changed.emit()

        def set_fields_enabled(self, enabled: bool) -> None:
            self.path_input.setEnabled(enabled)
            self.browse_button.setEnabled(enabled)


    class Metric(QFrame):
        def __init__(self, title: str) -> None:
            super().__init__()
            self.setObjectName("metric")
            layout = QVBoxLayout(self)
            layout.setContentsMargins(14, 10, 14, 10)
            layout.setSpacing(3)

            label = QLabel(title)
            label.setObjectName("metricLabel")
            self.value = QLabel("--")
            self.value.setObjectName("metricValue")
            layout.addWidget(label)
            layout.addWidget(self.value)

        def set_value(self, value: int | str) -> None:
            self.value.setText(f"{value:,}" if isinstance(value, int) else value)


    class MainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.settings = QSettings("PaySysTools", "BankNameMatcher")
            self.worker: MatchWorker | None = None
            self.started_at: float | None = None
            self.last_processed = 0
            self.last_total = 0
            self.last_summary: dict[str, int] = {}

            self.setWindowTitle("联行号匹配收款行")
            self.resize(980, 670)
            self.setMinimumSize(840, 580)
            self.build_ui()
            self.restore_last_paths()

            self.clock = QTimer(self)
            self.clock.setInterval(500)
            self.clock.timeout.connect(self.refresh_timing)

        def build_ui(self) -> None:
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setFrameShape(QFrame.Shape.NoFrame)
            self.setCentralWidget(scroll_area)

            content = QWidget()
            scroll_area.setWidget(content)
            layout = QVBoxLayout(content)
            layout.setContentsMargins(40, 34, 40, 30)
            layout.setSpacing(18)

            heading = QHBoxLayout()
            title = QLabel("联行号匹配收款行")
            title.setObjectName("title")
            heading.addWidget(title)
            heading.addStretch(1)
            self.run_state = QLabel("等待开始")
            self.run_state.setObjectName("runState")
            heading.addWidget(self.run_state)
            layout.addLayout(heading)

            file_section = QFrame()
            file_section.setObjectName("section")
            file_layout = QVBoxLayout(file_section)
            file_layout.setContentsMargins(22, 20, 22, 20)
            file_layout.setSpacing(14)
            section_title = QLabel("文件")
            section_title.setObjectName("sectionTitle")
            file_layout.addWidget(section_title)

            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
            form.setHorizontalSpacing(18)
            form.setVerticalSpacing(12)
            self.master_field = FileField()
            self.target_field = FileField()
            self.output_field = FileField(save_file=True)
            form.addRow("行名行号主数据", self.master_field)
            form.addRow("待匹配文件", self.target_field)

            scope_widget = QWidget()
            scope_layout = QHBoxLayout(scope_widget)
            scope_layout.setContentsMargins(0, 0, 0, 0)
            scope_layout.setSpacing(12)
            self.all_sheets_checkbox = QCheckBox("遍历所有工作表")
            self.all_sheets_checkbox.setToolTip("处理所有包含联行号列的工作表")
            self.sheet_input = QLineEdit()
            self.sheet_input.setMinimumWidth(260)
            self.sheet_input.setPlaceholderText("例如：Sheet1")
            scope_layout.addWidget(self.all_sheets_checkbox)
            scope_layout.addWidget(self.sheet_input, 1)
            form.addRow("匹配范围", scope_widget)
            form.addRow("结果文件", self.output_field)
            file_layout.addLayout(form)
            layout.addWidget(file_section)

            self.all_sheets_checkbox.toggled.connect(self.update_sheet_mode)

            progress_section = QFrame()
            progress_section.setObjectName("section")
            progress_layout = QVBoxLayout(progress_section)
            progress_layout.setContentsMargins(22, 20, 22, 20)
            progress_layout.setSpacing(12)
            progress_title = QLabel("执行进度")
            progress_title.setObjectName("sectionTitle")
            progress_layout.addWidget(progress_title)

            self.progress = QProgressBar()
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
            self.progress.setFormat("0 / 0")
            self.progress.setTextVisible(True)
            progress_layout.addWidget(self.progress)

            metric_layout = QGridLayout()
            metric_layout.setHorizontalSpacing(10)
            self.total_metric = Metric("总记录")
            self.processed_metric = Metric("已处理")
            self.matched_metric = Metric("已匹配")
            self.unmatched_metric = Metric("未匹配")
            for index, metric in enumerate(
                (
                    self.total_metric,
                    self.processed_metric,
                    self.matched_metric,
                    self.unmatched_metric,
                )
            ):
                metric_layout.addWidget(metric, 0, index)
                metric_layout.setColumnStretch(index, 1)
            progress_layout.addLayout(metric_layout)

            self.timing = QLabel("用时 --  |  预计剩余 --")
            self.timing.setObjectName("timing")
            progress_layout.addWidget(self.timing)
            layout.addWidget(progress_section)

            self.status = QLabel("请选择三个文件后开始匹配。")
            self.status.setObjectName("status")
            self.status.setWordWrap(True)
            layout.addWidget(self.status)

            buttons = QHBoxLayout()
            buttons.addStretch(1)
            self.cancel_button = QPushButton("取消")
            self.cancel_button.setEnabled(False)
            self.cancel_button.clicked.connect(self.cancel_matching)
            self.match_button = QPushButton("开始匹配")
            self.match_button.setObjectName("primaryButton")
            self.match_button.setDefault(True)
            self.match_button.clicked.connect(self.start_matching)
            buttons.addWidget(self.cancel_button)
            buttons.addWidget(self.match_button)
            layout.addLayout(buttons)

            self.setStyleSheet(
                """
                QMainWindow { background: #F4F6F8; }
                QScrollArea { background: #F4F6F8; }
                QLabel#title { color: #17212B; font-size: 22px; font-weight: 600; }
                QLabel#runState { color: #42616D; background: #E4EDF0; padding: 6px 12px; border-radius: 4px; }
                QFrame#section { background: #FFFFFF; border: 1px solid #D3DDE2; border-radius: 8px; }
                QLabel#sectionTitle { color: #2E4650; font-size: 13px; font-weight: 600; }
                QLineEdit { min-height: 30px; padding: 2px 8px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
                QLineEdit:focus { border: 2px solid #17818B; }
                QCheckBox { color: #263940; min-height: 30px; }
                QToolButton { min-width: 32px; min-height: 32px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
                QToolButton:hover { background: #EAF3F4; }
                QProgressBar { min-height: 22px; border: 1px solid #B9C7CF; border-radius: 4px; text-align: center; background: #F4F6F8; }
                QProgressBar::chunk { background: #17818B; border-radius: 3px; }
                QFrame#metric { background: #F8FAFB; border: 1px solid #E0E6E9; border-radius: 4px; }
                QLabel#metricLabel { color: #687780; font-size: 12px; }
                QLabel#metricValue { color: #17212B; font-size: 18px; font-weight: 600; }
                QLabel#timing { color: #50616A; }
                QLabel#status { color: #42545E; min-height: 22px; }
                QPushButton { min-width: 88px; min-height: 36px; padding: 0 14px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
                QPushButton:hover { background: #EEF3F5; }
                QPushButton#primaryButton { color: #FFFFFF; background: #167A83; border-color: #167A83; }
                QPushButton#primaryButton:hover { background: #10636B; }
                QPushButton:disabled { color: #94A2A9; background: #EDF1F3; border-color: #D9E0E4; }
                """
            )

        def restore_last_paths(self) -> None:
            self.master_field.set_value(self.settings.value("paths/master", "", type=str))
            self.target_field.set_value(self.settings.value("paths/target", "", type=str))
            self.output_field.set_value(self.settings.value("paths/output", "", type=str))
            all_sheets = self.settings.value("scope/all_sheets", False, type=bool)
            selected_sheet = self.settings.value("scope/sheet", "", type=str)
            self.all_sheets_checkbox.setChecked(all_sheets)
            self.sheet_input.setText(selected_sheet)
            self.update_sheet_mode(all_sheets)

        def save_last_paths(self) -> None:
            self.settings.setValue("paths/master", self.master_field.path_input.text().strip())
            self.settings.setValue("paths/target", self.target_field.path_input.text().strip())
            self.settings.setValue("paths/output", self.output_field.path_input.text().strip())
            self.settings.setValue("scope/all_sheets", self.all_sheets_checkbox.isChecked())
            self.settings.setValue("scope/sheet", self.sheet_input.text().strip())
            self.settings.sync()

        def update_sheet_mode(self, all_sheets: bool) -> None:
            self.sheet_input.setEnabled(not all_sheets)
            if all_sheets:
                self.sheet_input.setToolTip("已选择遍历所有包含联行号列的工作表")
            else:
                self.sheet_input.setToolTip("请输入需要匹配的工作表名称")

        def start_matching(self) -> None:
            master_path = self.master_field.value()
            target_path = self.target_field.value()
            output_path = self.output_field.value()

            if master_path is None or not master_path.is_file():
                self.show_input_error("请选择有效的行名行号主数据文件。")
                return
            if target_path is None or not target_path.is_file():
                self.show_input_error("请选择有效的待匹配文件。")
                return
            if output_path is None:
                self.show_input_error("请选择结果文件位置。")
                return
            if master_path.resolve() == target_path.resolve():
                self.show_input_error("主数据文件和待匹配文件不能相同。")
                return

            target_sheet_names: set[str] | None = None
            if not self.all_sheets_checkbox.isChecked():
                selected_sheet = self.sheet_input.text().strip()
                if not selected_sheet:
                    self.show_input_error("未勾选遍历所有工作表时，请输入一个具体工作表名称。")
                    return
                target_sheet_names = {selected_sheet}

            if output_path.exists():
                answer = QMessageBox.question(
                    self,
                    "确认覆盖",
                    f"结果文件已存在，是否覆盖？\n{output_path}",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return

            self.save_last_paths()
            self.reset_progress()
            self.set_running(True)
            self.run_state.setText("运行中")
            self.status.setText("正在准备匹配任务...")
            self.started_at = time.monotonic()
            self.clock.start()

            self.worker = MatchWorker(
                master_path.resolve(),
                target_path.resolve(),
                output_path.resolve(),
                target_sheet_names,
            )
            self.worker.stage_changed.connect(self.stage_changed)
            self.worker.progressed.connect(self.progressed)
            self.worker.completed.connect(self.match_completed)
            self.worker.cancelled.connect(self.match_cancelled)
            self.worker.failed.connect(self.match_failed)
            self.worker.start()

        def reset_progress(self) -> None:
            self.last_processed = 0
            self.last_total = 0
            self.last_summary = {}
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
            self.progress.setFormat("0 / 0")
            for metric in (
                self.total_metric,
                self.processed_metric,
                self.matched_metric,
                self.unmatched_metric,
            ):
                metric.set_value(0)
            self.timing.setText("用时 00:00  |  预计剩余 --")

        def set_running(self, running: bool) -> None:
            self.master_field.set_fields_enabled(not running)
            self.target_field.set_fields_enabled(not running)
            self.output_field.set_fields_enabled(not running)
            self.all_sheets_checkbox.setEnabled(not running)
            self.sheet_input.setEnabled(not running and not self.all_sheets_checkbox.isChecked())
            self.match_button.setEnabled(not running)
            self.cancel_button.setEnabled(running)

        def stage_changed(self, message: str) -> None:
            self.status.setText(message)
            if self.last_total == 0:
                self.progress.setRange(0, 0)
                self.progress.setFormat("")

        def progressed(self, processed: int, total: int, summary: dict) -> None:
            self.last_processed = processed
            self.last_total = total
            self.last_summary = summary
            self.progress.setRange(0, max(total, 1))
            self.progress.setValue(processed)
            percentage = 100 if total and processed >= total else int(processed * 100 / total) if total else 0
            self.progress.setFormat(f"{processed:,} / {total:,} ({percentage}%)")
            self.total_metric.set_value(total)
            self.processed_metric.set_value(processed)
            self.matched_metric.set_value(summary.get("matched", 0))
            self.unmatched_metric.set_value(
                summary.get("not_found", 0)
                + summary.get("duplicate", 0)
                + summary.get("empty", 0)
            )
            self.refresh_timing()

        def refresh_timing(self) -> None:
            if self.started_at is None:
                return
            elapsed = time.monotonic() - self.started_at
            if self.last_processed > 0 and self.last_total > self.last_processed:
                remaining = elapsed * (self.last_total - self.last_processed) / self.last_processed
                self.timing.setText(
                    f"用时 {self.format_duration(elapsed)}  |  预计剩余 {self.format_duration(remaining)}"
                )
            elif self.last_total and self.last_processed >= self.last_total:
                self.timing.setText(f"用时 {self.format_duration(elapsed)}  |  正在保存结果文件")
            else:
                self.timing.setText(f"用时 {self.format_duration(elapsed)}  |  预计剩余 --")

        @staticmethod
        def format_duration(seconds: float) -> str:
            total_seconds = max(0, round(seconds))
            hours, remainder = divmod(total_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes:02d}:{seconds:02d}"

        def cancel_matching(self) -> None:
            if self.worker is None:
                return
            self.worker.requestInterruption()
            self.cancel_button.setEnabled(False)
            self.status.setText("正在取消，等待当前批次结束...")

        def finish_run(self, state: str) -> None:
            self.clock.stop()
            self.set_running(False)
            self.run_state.setText(state)
            self.worker = None

        def match_completed(self, summary: dict, output_path: str, prefix_count: int) -> None:
            self.progressed(self.last_total, self.last_total, summary)
            self.finish_run("已完成")
            elapsed = 0 if self.started_at is None else time.monotonic() - self.started_at
            message = (
                f"已匹配 {summary.get('matched', 0):,} 条，未匹配 "
                f"{summary.get('not_found', 0) + summary.get('duplicate', 0) + summary.get('empty', 0):,} 条。\n"
                f"使用 {prefix_count} 个主数据类别，耗时 {self.format_duration(elapsed)}。\n"
                f"结果文件：{output_path}"
            )
            self.status.setText(message)
            QMessageBox.information(self, "匹配完成", message)

        def match_cancelled(self) -> None:
            self.finish_run("已取消")
            self.status.setText("匹配已取消，未生成结果文件。")

        def match_failed(self, error: str) -> None:
            self.finish_run("处理失败")
            self.status.setText("匹配失败，请检查文件格式和占用状态。")
            QMessageBox.critical(self, "匹配失败", error)

        def show_input_error(self, message: str) -> None:
            self.status.setText(message)
            QMessageBox.warning(self, "无法开始", message)

        def closeEvent(self, event) -> None:
            if self.worker and self.worker.isRunning():
                QMessageBox.information(self, "任务运行中", "请先取消或等待匹配任务完成。")
                event.ignore()
                return
            self.save_last_paths()
            event.accept()


def main() -> int:
    write_runtime_event(f"启动请求：Python {sys.version.split()[0]}")
    if QApplication is None:
        write_runtime_event("失败：PyQt5 不可用")
        print("缺少 PyQt5。请先执行：python -m pip install -r requirements.txt", file=sys.stderr)
        return 1

    try:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        write_runtime_event("QApplication 已创建")
        window = MainWindow()
        write_runtime_event("主窗口已创建")
        app.aboutToQuit.connect(lambda: write_runtime_event("收到正常退出信号"))
        window.show()
        write_runtime_event("主窗口已显示")
        smoke_test_seconds = os.environ.get("BANK_MATCH_GUI_SMOKE_TEST_SECONDS")
        if smoke_test_seconds:
            try:
                QTimer.singleShot(max(1, int(smoke_test_seconds)) * 1000, app.quit)
                write_runtime_event("已启用启动自检自动退出")
            except ValueError:
                write_runtime_event("忽略无效的启动自检时长")
        exit_code = app.exec_()
        write_runtime_event(f"事件循环正常结束，退出码：{exit_code}")
        return exit_code
    except Exception:
        error_log = APPLICATION_DIRECTORY / "gui_startup_error.log"
        error_log.write_text(traceback.format_exc(), encoding="utf-8")
        write_runtime_event(f"Python 初始化异常，详情：{error_log.name}")
        print(f"界面启动失败，详细错误已写入：{error_log}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
