import csv
import configparser
import ctypes
from ctypes import wintypes
import logging
from logging.handlers import RotatingFileHandler
import os
import sys

from PySide6.QtCore import QEvent, Qt, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

WINDOW_SIZE = (700, 300)
COLUMN_WIDTH_RATIO = (3, 2)  # 文字列:説明
ROW_HEIGHT = 20
ICON_RELATIVE_PATH = os.path.join("ico", "CopyList.ico")
INI_FILENAME = "copylist.ini"
INI_SECTION = "settings"
CSV_EMPTY_LABEL = "(未選択)"
LOG_FILENAME = "copylist.log"
LOG_MAX_BYTES = 512 * 1024
LOG_BACKUP_COUNT = 2
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOACTIVATE = 0x0010
SWP_NOOWNERZORDER = 0x0200
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2


class CopyListWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CopyList")
        self.resize(*WINDOW_SIZE)

        self._suspend_events = False  # 変更イベントの再入を抑止
        self.app_dir = self._get_app_dir()
        self.log_path = os.path.join(self.app_dir, LOG_FILENAME)
        self.logger = self._setup_logger()
        self._user32 = self._load_user32()
        self.ini_path = os.path.join(self.app_dir, INI_FILENAME)
        self.csv_path = ""
        self.csv_encoding = "utf-8-sig"  # 既定はUTF-8(BOM付き)
        self._settings = self._load_settings()
        self.logger.info("CopyList start app_dir=%s", self.app_dir)
        self._set_window_icon()

        self.csv_combo = QComboBox(self)
        self.csv_combo.setMinimumWidth(220)
        self.pause_copy_checkbox = QCheckBox("一時停止", self)
        self.always_on_top_checkbox = QCheckBox("最前面", self)

        self.table = QTableWidget(0, 2, self)
        self.table.setHorizontalHeaderLabels(["文字列", "説明"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(ROW_HEIGHT)
        self.table.verticalHeader().setMinimumSectionSize(ROW_HEIGHT)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)

        btn_up = QPushButton("↑", self)
        btn_down = QPushButton("↓", self)
        btn_up.setFixedSize(50, 44)
        btn_down.setFixedSize(50, 44)

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("CSV:", self))
        top_layout.addWidget(self.csv_combo)
        top_layout.addStretch(1)
        top_layout.addWidget(self.always_on_top_checkbox)
        top_layout.addWidget(self.pause_copy_checkbox)

        content_layout = QHBoxLayout()
        content_layout.addWidget(self.table, 1)
        btn_layout = QVBoxLayout()
        btn_layout.addWidget(btn_up)
        btn_layout.addWidget(btn_down)
        btn_layout.addStretch(1)
        content_layout.addLayout(btn_layout)

        main_layout = QVBoxLayout(self)
        main_layout.addLayout(top_layout)
        main_layout.addLayout(content_layout, 1)

        # 編集/選択/クリックのイベント
        self.table.itemChanged.connect(self.on_value_changed)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.cellClicked.connect(self.on_cell_clicked)
        self.csv_combo.currentIndexChanged.connect(self.on_csv_changed)
        self.pause_copy_checkbox.toggled.connect(self.on_pause_toggled)
        self.always_on_top_checkbox.toggled.connect(self.on_always_on_top_toggled)
        btn_up.clicked.connect(self.on_move_up)
        btn_down.clicked.connect(self.on_move_down)

        self._initialize_controls_from_settings()
        self._restore_window_position_from_settings()
        self.load_csv()  # 起動時にCSV読み込み
        self.ensure_trailing_empty()
        QTimer.singleShot(0, self._apply_default_column_ratio)

    def _setup_logger(self):
        logger = logging.getLogger("copylist")
        logger.setLevel(logging.DEBUG)
        logger.propagate = False

        target_path = os.path.abspath(self.log_path)
        for handler in logger.handlers:
            base_path = getattr(handler, "baseFilename", "")
            if base_path and os.path.abspath(base_path) == target_path:
                return logger

        try:
            handler = RotatingFileHandler(
                self.log_path,
                maxBytes=LOG_MAX_BYTES,
                backupCount=LOG_BACKUP_COUNT,
                encoding="utf-8",
            )
            handler.setLevel(logging.DEBUG)
            handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            logger.addHandler(handler)
        except OSError:
            logger.addHandler(logging.NullHandler())
        return logger

    def _load_user32(self):
        if sys.platform != "win32":
            return None
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            user32.SetWindowPos.argtypes = [
                wintypes.HWND,
                wintypes.HWND,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_uint,
            ]
            user32.SetWindowPos.restype = wintypes.BOOL
            return user32
        except Exception:
            self.logger.exception("user32 load failed")
            return None

    def _window_state_snapshot(self):
        geo = self.geometry()
        try:
            hwnd = int(self.winId())
        except Exception:
            hwnd = 0
        return (
            f"visible={self.isVisible()} hidden={self.isHidden()} active={self.isActiveWindow()} "
            f"min={self.isMinimized()} max={self.isMaximized()} "
            f"flags=0x{int(self.windowFlags()):x} hwnd={hwnd} "
            f"geo=({geo.x()},{geo.y()},{geo.width()},{geo.height()})"
        )

    def _log_window_state(self, label):
        self.logger.debug("%s | %s", label, self._window_state_snapshot())

    def _apply_default_column_ratio(self):
        total = COLUMN_WIDTH_RATIO[0] + COLUMN_WIDTH_RATIO[1]
        if total <= 0:
            return

        viewport_width = max(1, self.table.viewport().width())
        col0_width = int(viewport_width * COLUMN_WIDTH_RATIO[0] / total)
        col0_width = max(80, col0_width)
        col1_width = max(80, viewport_width - col0_width)
        self.table.setColumnWidth(0, col0_width)
        self.table.setColumnWidth(1, col1_width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_default_column_ratio()

    def _get_app_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(sys.argv[0]))

    def _find_icon_path(self):
        candidates = []
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(os.path.join(meipass, ICON_RELATIVE_PATH))

        candidates.append(os.path.join(self.app_dir, ICON_RELATIVE_PATH))
        candidates.append(os.path.join(self.app_dir, "_internal", ICON_RELATIVE_PATH))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), ICON_RELATIVE_PATH))

        for path in candidates:
            if os.path.isfile(path):
                return path
        return ""

    def _set_window_icon(self):
        icon_path = self._find_icon_path()
        if not icon_path:
            return

        icon = QIcon(icon_path)
        if icon.isNull():
            return

        self.setWindowIcon(icon)
        app = QApplication.instance()
        if app is not None:
            app.setWindowIcon(icon)

    def _load_settings(self):
        defaults = {
            "last_csv": "",
            "always_on_top": False,
            "pause_copy": False,
            "window_x": None,
            "window_y": None,
        }
        if not os.path.exists(self.ini_path):
            return defaults

        parser = configparser.ConfigParser()
        try:
            parser.read(self.ini_path, encoding="utf-8")
            defaults["last_csv"] = parser.get(INI_SECTION, "last_csv", fallback="").strip()
            defaults["always_on_top"] = parser.getboolean(INI_SECTION, "always_on_top", fallback=False)
            defaults["pause_copy"] = parser.getboolean(INI_SECTION, "pause_copy", fallback=False)
            defaults["window_x"] = self._parse_optional_int(parser.get(INI_SECTION, "window_x", fallback=""))
            defaults["window_y"] = self._parse_optional_int(parser.get(INI_SECTION, "window_y", fallback=""))
        except (configparser.Error, OSError, ValueError):
            return {
                "last_csv": "",
                "always_on_top": False,
                "pause_copy": False,
                "window_x": None,
                "window_y": None,
            }
        return defaults

    @staticmethod
    def _parse_optional_int(value):
        text = str(value).strip()
        if not text:
            return None
        try:
            return int(text)
        except ValueError:
            return None

    def _save_settings(self):
        parser = configparser.ConfigParser()
        parser[INI_SECTION] = {
            "last_csv": self._current_csv_name(),
            "always_on_top": "1" if self.always_on_top_checkbox.isChecked() else "0",
            "pause_copy": "1" if self.pause_copy_checkbox.isChecked() else "0",
            "window_x": str(self.x()),
            "window_y": str(self.y()),
        }
        try:
            with open(self.ini_path, "w", encoding="utf-8") as f:
                parser.write(f)
            self.logger.debug("settings saved ini=%s", self.ini_path)
        except OSError:
            self.logger.exception("settings save failed ini=%s", self.ini_path)
            pass

    def _restore_window_position_from_settings(self):
        x = self._settings.get("window_x")
        y = self._settings.get("window_y")
        if isinstance(x, int) and isinstance(y, int):
            self.move(x, y)
            self.logger.debug("window position restored x=%s y=%s", x, y)

    def _list_csv_files(self):
        files = []
        try:
            for name in os.listdir(self.app_dir):
                path = os.path.join(self.app_dir, name)
                if os.path.isfile(path) and name.lower().endswith(".csv"):
                    files.append(name)
        except OSError:
            return []
        return sorted(files, key=str.lower)

    def _current_csv_name(self):
        data = self.csv_combo.currentData()
        return data if isinstance(data, str) else ""

    def _set_csv_path_from_name(self, name):
        if name:
            self.csv_path = os.path.join(self.app_dir, name)
        else:
            self.csv_path = ""

    def _refresh_csv_combo(self, selected_name=""):
        files = self._list_csv_files()
        self.csv_combo.blockSignals(True)
        try:
            self.csv_combo.clear()
            self.csv_combo.addItem(CSV_EMPTY_LABEL, "")
            for file_name in files:
                self.csv_combo.addItem(file_name, file_name)

            index = self.csv_combo.findData(selected_name)
            self.csv_combo.setCurrentIndex(index if index >= 0 else 0)
        finally:
            self.csv_combo.blockSignals(False)

        self._set_csv_path_from_name(self._current_csv_name())

    def _initialize_controls_from_settings(self):
        last_csv = self._settings.get("last_csv", "")
        self._refresh_csv_combo(last_csv)

        self.pause_copy_checkbox.blockSignals(True)
        self.pause_copy_checkbox.setChecked(bool(self._settings.get("pause_copy", False)))
        self.pause_copy_checkbox.blockSignals(False)

        self.always_on_top_checkbox.blockSignals(True)
        self.always_on_top_checkbox.setChecked(bool(self._settings.get("always_on_top", False)))
        self.always_on_top_checkbox.blockSignals(False)
        self._apply_always_on_top(self.always_on_top_checkbox.isChecked())

    def load_csv(self):
        self._suspend_events = True
        self.table.setRowCount(0)

        rows = []
        if self.csv_path and os.path.exists(self.csv_path):
            rows, encoding = self._read_csv_with_fallback()
            self.csv_encoding = encoding

        while rows and self._row_is_empty_values(rows[-1]):
            rows.pop()

        for row_values in rows:
            row_values = (row_values + ["", ""])[:2]
            self._append_row(row_values[0], row_values[1])

        self._suspend_events = False

    def _append_row(self, col0, col1):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self._set_row_values(row, [col0, col1])

    def _read_csv_with_fallback(self):
        if not self.csv_path:
            return [], self.csv_encoding

        # 文字コードを順に試して読み込む
        for encoding in ("utf-8-sig", "cp932"):
            try:
                with open(self.csv_path, newline="", encoding=encoding) as f:
                    reader = csv.reader(f)
                    rows = [[c for c in row] for row in reader]
                return rows, encoding
            except UnicodeDecodeError:
                continue

        with open(self.csv_path, newline="", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            rows = [[c for c in row] for row in reader]
        return rows, "utf-8-sig"

    def save_csv(self):
        if not self.csv_path:
            return
        rows = self._collect_rows_for_save()
        os.makedirs(os.path.dirname(self.csv_path), exist_ok=True)
        try:
            with open(self.csv_path, "w", newline="", encoding=self.csv_encoding) as f:
                writer = csv.writer(f)
                writer.writerows(rows)
        except UnicodeEncodeError:
            # 保存時に失敗したらUTF-8に切り替えて再保存
            self.csv_encoding = "utf-8-sig"
            with open(self.csv_path, "w", newline="", encoding=self.csv_encoding) as f:
                writer = csv.writer(f)
                writer.writerows(rows)

    def _collect_rows_for_save(self):
        rows = []
        for row in range(self.table.rowCount()):
            rows.append([self._get_text(row, 0), self._get_text(row, 1)])
        # 末尾の空行は保存しない
        while rows and self._row_is_empty_values(rows[-1]):
            rows.pop()
        return rows

    def _get_text(self, row, col):
        item = self.table.item(row, col)
        return item.text() if item else ""

    def _set_text(self, row, col, text):
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, col, item)
        if item.text() != text:
            item.setText(text)

    def _set_row_values(self, row, values):
        self._set_text(row, 0, values[0] if len(values) > 0 else "")
        self._set_text(row, 1, values[1] if len(values) > 1 else "")

    def _row_is_empty(self, row):
        if row < 0 or row >= self.table.rowCount():
            return True
        return self._get_text(row, 0).strip() == "" and self._get_text(row, 1).strip() == ""

    @staticmethod
    def _row_is_empty_values(values):
        if not values:
            return True
        v1 = (values[0] if len(values) > 0 else "").strip()
        v2 = (values[1] if len(values) > 1 else "").strip()
        return v1 == "" and v2 == ""

    def ensure_trailing_empty(self):
        previous_suspend = self._suspend_events
        self._suspend_events = True
        try:
            count = self.table.rowCount()

            # 末尾の空行が複数ある場合は1つにまとめる
            while count > 1 and self._row_is_empty(count - 1) and self._row_is_empty(count - 2):
                self.table.removeRow(count - 1)
                count -= 1

            if count == 0 or not self._row_is_empty(count - 1):
                # 最終行は常に空行にする
                self._append_row("", "")
        finally:
            self._suspend_events = previous_suspend

    def on_value_changed(self, item):
        if self._suspend_events:
            return
        self.ensure_trailing_empty()
        self.save_csv()

    def on_selection_changed(self):
        self._copy_selected_text()

    def on_cell_clicked(self, row, col):
        # クリック時にもコピーする
        self._copy_selected_text()

    def on_csv_changed(self, index):
        self._set_csv_path_from_name(self._current_csv_name())
        self.logger.debug("csv changed index=%s path=%s", index, self.csv_path)
        self.load_csv()
        self.ensure_trailing_empty()
        self._save_settings()

    def _apply_always_on_top(self, checked):
        self._log_window_state(f"topmost apply start checked={checked}")
        if sys.platform == "win32" and self.isVisible():
            if self._set_windows_topmost(checked):
                self._log_window_state(f"topmost apply end(winapi) checked={checked}")
                if checked:
                    self.raise_()
                    self.activateWindow()
                return

        # setWindowFlag() は表示中ウィンドウを一旦隠すため、
        # 変更前の表示状態を保持して再表示する。
        was_visible = self.isVisible()
        self.setWindowFlag(Qt.WindowStaysOnTopHint, checked)
        if was_visible:
            self.show()
            if checked:
                self.raise_()
                self.activateWindow()
        self._log_window_state(f"topmost apply end(qt-fallback) checked={checked}")

    def _set_windows_topmost(self, checked):
        try:
            hwnd = int(self.winId())
        except Exception:
            self.logger.exception("winId acquisition failed")
            return False
        if hwnd == 0:
            self.logger.warning("winId is 0, skip SetWindowPos")
            return False

        flags = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOOWNERZORDER
        insert_after = HWND_TOPMOST if checked else HWND_NOTOPMOST
        try:
            if self._user32 is None:
                self.logger.warning("user32 is not available, skip SetWindowPos")
                return False
            ctypes.set_last_error(0)
            result = self._user32.SetWindowPos(hwnd, insert_after, 0, 0, 0, 0, flags)
            err = ctypes.get_last_error()
            self.logger.debug(
                "SetWindowPos checked=%s hwnd=%s insert_after=%s flags=0x%x result=%s last_error=%s",
                checked,
                hwnd,
                insert_after,
                flags,
                bool(result),
                err,
            )
        except Exception:
            self.logger.exception("SetWindowPos failed")
            return False
        return bool(result)

    def on_always_on_top_toggled(self, checked):
        self.logger.info("topmost toggled checked=%s", checked)
        self._apply_always_on_top(checked)
        self._save_settings()

    def on_pause_toggled(self, checked):
        self._save_settings()

    def _copy_selected_text(self):
        if self.pause_copy_checkbox.isChecked():
            return
        row = self._get_selected_row()
        if row < 0:
            return
        text = self._get_text(row, 0)
        if text:
            QApplication.clipboard().setText(text)

    def _get_selected_row(self):
        row = self.table.currentRow()
        return row if row >= 0 else -1

    def on_move_up(self):
        row = self._get_selected_row()
        if row <= 0:
            return
        if self._row_is_empty(row):
            return
        self._swap_rows(row, row - 1)
        self._select_row(row - 1)
        self.ensure_trailing_empty()
        self.save_csv()

    def on_move_down(self):
        row = self._get_selected_row()
        count = self.table.rowCount()
        if row < 0 or row >= count - 1:
            return
        if self._row_is_empty(row):
            return
        if count >= 1 and self._row_is_empty(count - 1) and row >= count - 2:
            return
        self._swap_rows(row, row + 1)
        self._select_row(row + 1)
        self.ensure_trailing_empty()
        self.save_csv()

    def _swap_rows(self, row_a, row_b):
        values_a = [self._get_text(row_a, 0), self._get_text(row_a, 1)]
        values_b = [self._get_text(row_b, 0), self._get_text(row_b, 1)]

        # 値を書き換える間は変更イベントを止める
        self._suspend_events = True
        try:
            self._set_row_values(row_a, values_b)
            self._set_row_values(row_b, values_a)
        finally:
            self._suspend_events = False

    def _select_row(self, row):
        if row < 0 or row >= self.table.rowCount():
            return
        self.table.setCurrentCell(row, 0)
        self.table.selectRow(row)
        item = self.table.item(row, 0)
        if item is not None:
            self.table.scrollToItem(item)

    def closeEvent(self, event):
        self._log_window_state("closeEvent")
        self._save_settings()
        self.logger.info("CopyList closing")
        super().closeEvent(event)

    def showEvent(self, event):
        super().showEvent(event)
        self._log_window_state("showEvent")

    def hideEvent(self, event):
        super().hideEvent(event)
        self._log_window_state("hideEvent")

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.Type.WindowStateChange:
            self._log_window_state("changeEvent(WindowStateChange)")

    def event(self, event):
        if event.type() == QEvent.Type.WindowActivate:
            self._log_window_state("event(WindowActivate)")
        elif event.type() == QEvent.Type.WindowDeactivate:
            self._log_window_state("event(WindowDeactivate)")
        return super().event(event)


def main():
    app = QApplication(sys.argv)
    window = CopyListWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
