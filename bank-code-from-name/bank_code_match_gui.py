"""根据收款开户行匹配支付系统行号的 PyQt5 图形界面。"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

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

from match_bank_codes import MatchingCancelled, process_target_file, read_master_data


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
            self.stage_changed.emit("正在读取行名行号主数据...")
            master_sheets = read_master_data(self.master_path)
            if self.isInterruptionRequested():
                raise MatchingCancelled()

            self.stage_changed.emit(f"正在匹配 {len(master_sheets)} 个主数据类别...")
            summary = process_target_file(
                self.target_path,
                self.output_path,
                master_sheets,
                progress_callback=self.report_progress,
                cancel_callback=self.isInterruptionRequested,
                target_sheet_names=self.target_sheet_names,
            )
            self.completed.emit(summary, str(self.output_path), len(master_sheets))
        except MatchingCancelled:
            self.cancelled.emit()
        except Exception as error:
            self.failed.emit(str(error))

    def report_progress(self, processed: int, total: int, summary: dict[str, int]) -> None:
        self.progressed.emit(processed, total, summary)


class FileField(QWidget):
    path_changed = pyqtSignal()

    def __init__(self, *, file_filter: str, save_file: bool = False) -> None:
        super().__init__()
        self.file_filter = file_filter
        self.save_file = save_file
        self.path_input = QLineEdit()
        self.path_input.setClearButtonEnabled(True)
        self.path_input.setPlaceholderText("请选择文件")
        self.path_input.editingFinished.connect(self.path_changed.emit)

        self.browse_button = QToolButton()
        icon = QStyle.SP_DialogSaveButton if save_file else QStyle.SP_DirOpenIcon
        self.browse_button.setIcon(self.style().standardIcon(icon))
        self.browse_button.setToolTip("选择结果文件" if save_file else "选择文件")
        self.browse_button.clicked.connect(self.select_file)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        layout.addWidget(self.path_input, 1)
        layout.addWidget(self.browse_button)

    def value(self) -> Path | None:
        value = self.path_input.text().strip()
        return Path(value) if value else None

    def set_value(self, value: Path | str | None) -> None:
        self.path_input.setText("" if value is None else str(value))

    def set_enabled(self, enabled: bool) -> None:
        self.path_input.setEnabled(enabled)
        self.browse_button.setEnabled(enabled)

    def select_file(self) -> None:
        selected = self.value()
        directory = selected.parent if selected and selected.parent.is_dir() else Path.home()
        if self.save_file:
            file_name, _ = QFileDialog.getSaveFileName(
                self, "选择结果文件", str(directory), self.file_filter
            )
            if file_name:
                path = Path(file_name)
                self.set_value(path if path.suffix else path.with_suffix(".xlsx"))
                self.path_changed.emit()
            return

        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", str(directory), self.file_filter
        )
        if file_name:
            self.set_value(file_name)
            self.path_changed.emit()


class Metric(QFrame):
    def __init__(self, label: str) -> None:
        super().__init__()
        self.setObjectName("metric")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)
        title = QLabel(label)
        title.setObjectName("metricLabel")
        self.value = QLabel("0")
        self.value.setObjectName("metricValue")
        layout.addWidget(title)
        layout.addWidget(self.value)

    def set_value(self, value: int) -> None:
        self.value.setText(f"{value:,}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.settings = QSettings("PaySysTools", "BankCodeMatcher")
        self.worker: MatchWorker | None = None
        self.started_at: float | None = None
        self.last_processed = 0
        self.last_total = 0

        self.setWindowTitle("收款开户行匹配支付系统行号")
        self.setMinimumSize(720, 480)
        self.resize(920, 640)
        self.build_ui()
        self.restore_settings()

        self.clock = QTimer(self)
        self.clock.setInterval(500)
        self.clock.timeout.connect(self.refresh_timing)

    def build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        central_layout.addWidget(scroll_area)

        content = QWidget()
        content.setMinimumWidth(680)
        scroll_area.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        heading = QHBoxLayout()
        title = QLabel("收款开户行匹配支付系统行号")
        title.setObjectName("title")
        heading.addWidget(title)
        heading.addStretch(1)
        self.run_state = QLabel("等待开始")
        self.run_state.setObjectName("runState")
        heading.addWidget(self.run_state)
        layout.addLayout(heading)

        files = QFrame()
        files.setObjectName("section")
        files_layout = QVBoxLayout(files)
        files_layout.setContentsMargins(20, 18, 20, 18)
        files_layout.setSpacing(12)
        heading_label = QLabel("文件")
        heading_label.setObjectName("sectionTitle")
        files_layout.addWidget(heading_label)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)
        self.master_field = FileField(file_filter="Excel 工作簿 (*.xlsx)")
        self.target_field = FileField(file_filter="Excel 工作簿 (*.xlsx)")
        self.output_field = FileField(file_filter="Excel 工作簿 (*.xlsx)", save_file=True)
        form.addRow("行名行号主数据", self.master_field)
        form.addRow("待匹配供应商表", self.target_field)

        scope = QWidget()
        scope_layout = QHBoxLayout(scope)
        scope_layout.setContentsMargins(0, 0, 0, 0)
        scope_layout.setSpacing(10)
        self.all_sheets = QCheckBox("遍历所有工作表")
        self.all_sheets.setChecked(True)
        self.all_sheets.setToolTip("处理所有包含收款开户行列的工作表")
        self.sheet_name = QLineEdit()
        self.sheet_name.setPlaceholderText("例如：Sheet1")
        scope_layout.addWidget(self.all_sheets)
        scope_layout.addWidget(self.sheet_name, 1)
        form.addRow("匹配范围", scope)
        form.addRow("结果文件", self.output_field)
        files_layout.addLayout(form)
        layout.addWidget(files)

        progress_section = QFrame()
        progress_section.setObjectName("section")
        progress_layout = QVBoxLayout(progress_section)
        progress_layout.setContentsMargins(20, 18, 20, 18)
        progress_layout.setSpacing(12)
        progress_label = QLabel("执行进度")
        progress_label.setObjectName("sectionTitle")
        progress_layout.addWidget(progress_label)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setFormat("0 / 0")
        progress_layout.addWidget(self.progress)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(10)
        self.total_metric = Metric("总记录")
        self.processed_metric = Metric("已处理")
        self.matched_metric = Metric("已匹配")
        self.unmatched_metric = Metric("未匹配")
        for index, metric in enumerate(
            (self.total_metric, self.processed_metric, self.matched_metric, self.unmatched_metric)
        ):
            metrics.addWidget(metric, 0, index)
            metrics.setColumnStretch(index, 1)
        progress_layout.addLayout(metrics)
        self.timing = QLabel("用时 --  |  预计剩余 --")
        self.timing.setObjectName("timing")
        progress_layout.addWidget(self.timing)
        layout.addWidget(progress_section)

        self.status = QLabel("请选择行名行号主数据和供应商表。")
        self.status.setObjectName("status")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.cancel_button = QPushButton("取消")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_matching)
        self.start_button = QPushButton("开始匹配")
        self.start_button.setObjectName("primaryButton")
        self.start_button.setDefault(True)
        self.start_button.clicked.connect(self.start_matching)
        buttons.addWidget(self.cancel_button)
        buttons.addWidget(self.start_button)
        layout.addLayout(buttons)

        self.all_sheets.toggled.connect(self.update_scope_state)
        self.target_field.path_changed.connect(self.set_default_output)
        self.setStyleSheet(
            """
            QMainWindow { background: #F4F6F8; }
            QLabel#title { color: #17212B; font-size: 22px; font-weight: 600; }
            QLabel#runState { color: #42616D; background: #E4EDF0; padding: 6px 12px; border-radius: 4px; }
            QFrame#section { background: #FFFFFF; border: 1px solid #D3DDE2; border-radius: 8px; }
            QLabel#sectionTitle { color: #2E4650; font-size: 13px; font-weight: 600; }
            QLineEdit { min-height: 30px; padding: 2px 8px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
            QLineEdit:focus { border: 2px solid #17818B; }
            QCheckBox { color: #263940; min-height: 30px; }
            QToolButton { min-width: 32px; min-height: 32px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
            QToolButton:hover, QPushButton:hover { background: #EEF3F5; }
            QProgressBar { min-height: 22px; border: 1px solid #B9C7CF; border-radius: 4px; text-align: center; background: #F4F6F8; }
            QProgressBar::chunk { background: #17818B; border-radius: 3px; }
            QFrame#metric { background: #F8FAFB; border: 1px solid #E0E6E9; border-radius: 4px; }
            QLabel#metricLabel { color: #687780; font-size: 12px; }
            QLabel#metricValue { color: #17212B; font-size: 18px; font-weight: 600; }
            QLabel#timing, QLabel#status { color: #50616A; }
            QPushButton { min-width: 88px; min-height: 36px; padding: 0 14px; border: 1px solid #B9C7CF; border-radius: 4px; background: #FFFFFF; }
            QPushButton#primaryButton { color: #FFFFFF; background: #167A83; border-color: #167A83; }
            QPushButton#primaryButton:hover { background: #10636B; }
            QPushButton:disabled { color: #94A2A9; background: #EDF1F3; border-color: #D9E0E4; }
            """
        )

    def restore_settings(self) -> None:
        self.master_field.set_value(self.settings.value("master", "", type=str))
        self.target_field.set_value(self.settings.value("target", "", type=str))
        self.output_field.set_value(self.settings.value("output", "", type=str))
        self.all_sheets.setChecked(self.settings.value("all_sheets", True, type=bool))
        self.sheet_name.setText(self.settings.value("sheet_name", "", type=str))
        self.update_scope_state(self.all_sheets.isChecked())

    def save_settings(self) -> None:
        self.settings.setValue("master", self.master_field.path_input.text().strip())
        self.settings.setValue("target", self.target_field.path_input.text().strip())
        self.settings.setValue("output", self.output_field.path_input.text().strip())
        self.settings.setValue("all_sheets", self.all_sheets.isChecked())
        self.settings.setValue("sheet_name", self.sheet_name.text().strip())
        self.settings.sync()

    def update_scope_state(self, all_sheets: bool) -> None:
        self.sheet_name.setEnabled(not all_sheets)

    def set_default_output(self) -> None:
        target_path = self.target_field.value()
        if target_path and not self.output_field.value():
            self.output_field.set_value(target_path.with_name(f"{target_path.stem}_匹配结果.xlsx"))

    def start_matching(self) -> None:
        master_path = self.master_field.value()
        target_path = self.target_field.value()
        output_path = self.output_field.value()
        if master_path is None or not master_path.is_file():
            self.show_error("请选择有效的行名行号主数据文件。")
            return
        if master_path.suffix.lower() != ".xlsx":
            self.show_error("行名行号主数据不是 .xlsx 文件，请先在 Excel 中另存为 .xlsx 后再运行。")
            return
        if target_path is None or not target_path.is_file():
            self.show_error("请选择有效的待匹配供应商表。")
            return
        if target_path.suffix.lower() != ".xlsx":
            self.show_error("待匹配供应商表不是 .xlsx 文件，请先在 Excel 中另存为 .xlsx 后再运行。")
            return
        if output_path is None:
            output_path = target_path.with_name(f"{target_path.stem}_匹配结果.xlsx")
            self.output_field.set_value(output_path)
        if output_path.suffix.lower() != ".xlsx":
            self.show_error("结果文件必须使用 .xlsx 扩展名。")
            return
        if master_path.resolve() == target_path.resolve():
            self.show_error("主数据文件和待匹配文件不能相同。")
            return
        if output_path.resolve() in {master_path.resolve(), target_path.resolve()}:
            self.show_error("结果文件不能覆盖主数据文件或待匹配文件。")
            return

        target_sheets: set[str] | None = None
        if not self.all_sheets.isChecked():
            selected_sheet = self.sheet_name.text().strip()
            if not selected_sheet:
                self.show_error("请输入需要处理的工作表名称。")
                return
            target_sheets = {selected_sheet}
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

        self.save_settings()
        self.reset_progress()
        self.set_running(True)
        self.started_at = time.monotonic()
        self.clock.start()
        self.run_state.setText("运行中")
        self.status.setText("正在准备匹配任务...")
        self.worker = MatchWorker(
            master_path.resolve(), target_path.resolve(), output_path.resolve(), target_sheets
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
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setFormat("0 / 0")
        for metric in (self.total_metric, self.processed_metric, self.matched_metric, self.unmatched_metric):
            metric.set_value(0)
        self.timing.setText("用时 00:00  |  预计剩余 --")

    def set_running(self, running: bool) -> None:
        self.master_field.set_enabled(not running)
        self.target_field.set_enabled(not running)
        self.output_field.set_enabled(not running)
        self.all_sheets.setEnabled(not running)
        self.sheet_name.setEnabled(not running and not self.all_sheets.isChecked())
        self.start_button.setEnabled(not running)
        self.cancel_button.setEnabled(running)

    def stage_changed(self, message: str) -> None:
        self.status.setText(message)
        if self.last_total == 0:
            self.progress.setRange(0, 0)
            self.progress.setFormat("")

    def progressed(self, processed: int, total: int, summary: dict) -> None:
        self.last_processed = processed
        self.last_total = total
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(processed)
        percentage = int(processed * 100 / total) if total else 0
        self.progress.setFormat(f"{processed:,} / {total:,} ({percentage}%)")
        self.total_metric.set_value(total)
        self.processed_metric.set_value(processed)
        self.matched_metric.set_value(summary.get("matched", 0))
        self.unmatched_metric.set_value(
            summary.get("ambiguous", 0) + summary.get("not_found", 0) + summary.get("empty", 0)
        )
        self.refresh_timing()

    def refresh_timing(self) -> None:
        if self.started_at is None:
            return
        elapsed = time.monotonic() - self.started_at
        if self.last_processed and self.last_total > self.last_processed:
            remaining = elapsed * (self.last_total - self.last_processed) / self.last_processed
            self.timing.setText(f"用时 {self.format_duration(elapsed)}  |  预计剩余 {self.format_duration(remaining)}")
        elif self.last_total and self.last_processed >= self.last_total:
            self.timing.setText(f"用时 {self.format_duration(elapsed)}  |  正在保存结果文件")
        else:
            self.timing.setText(f"用时 {self.format_duration(elapsed)}  |  预计剩余 --")

    @staticmethod
    def format_duration(seconds: float) -> str:
        minutes, seconds = divmod(max(0, round(seconds)), 60)
        hours, minutes = divmod(minutes, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes:02d}:{seconds:02d}"

    def cancel_matching(self) -> None:
        if self.worker is not None:
            self.worker.requestInterruption()
            self.cancel_button.setEnabled(False)
            self.status.setText("正在取消，等待当前批次结束...")

    def finish_run(self, state: str) -> None:
        self.clock.stop()
        self.set_running(False)
        self.run_state.setText(state)
        self.worker = None

    def match_completed(self, summary: dict, output_path: str, sheet_count: int) -> None:
        self.progressed(self.last_total, self.last_total, summary)
        self.finish_run("已完成")
        elapsed = 0 if self.started_at is None else time.monotonic() - self.started_at
        message = (
            f"已匹配 {summary.get('matched', 0):,} 条，未匹配 "
            f"{summary.get('ambiguous', 0) + summary.get('not_found', 0) + summary.get('empty', 0):,} 条。\n"
            f"已比对 {sheet_count} 个主数据类别，耗时 {self.format_duration(elapsed)}。\n"
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

    def show_error(self, message: str) -> None:
        self.status.setText(message)
        QMessageBox.warning(self, "无法开始", message)

    def closeEvent(self, event) -> None:
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "任务运行中", "请先取消或等待匹配完成。")
            event.ignore()
            return
        self.save_settings()
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    smoke_test_seconds = os.environ.get("BANK_CODE_MATCH_GUI_SMOKE_TEST_SECONDS")
    if smoke_test_seconds:
        QTimer.singleShot(max(1, int(smoke_test_seconds)) * 1000, app.quit)
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
