import csv
import os
import sys

import wx
import wx.dataview as dv


class CopyListFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="CopyList", size=(700, 300))

        self._suspend_events = False  # 変更イベントの再入を抑止
        self.app_dir = self._get_app_dir()
        self.csv_path = os.path.join(self.app_dir, "copylist.csv")
        self.csv_encoding = "utf-8-sig"  # 既定はUTF-8(BOM付き)

        panel = wx.Panel(self)

        self.dvlc = dv.DataViewListCtrl(
            panel,
            style=dv.DV_ROW_LINES | dv.DV_VERT_RULES | dv.DV_SINGLE,
        )
        self.dvlc.AppendTextColumn(
            "文字列",
            width=400,
            mode=dv.DATAVIEW_CELL_EDITABLE,
        )
        self.dvlc.AppendTextColumn(
            "説明",
            width=100,
            mode=dv.DATAVIEW_CELL_EDITABLE,
        )

        btn_up = wx.Button(panel, label="↑", size=(44, 32))
        btn_down = wx.Button(panel, label="↓", size=(44, 32))

        main_sizer = wx.BoxSizer(wx.HORIZONTAL)
        main_sizer.Add(self.dvlc, 1, wx.EXPAND | wx.ALL, 8)

        btn_sizer = wx.BoxSizer(wx.VERTICAL)
        btn_sizer.Add(btn_up, 0, wx.BOTTOM, 6)
        btn_sizer.Add(btn_down, 0)
        main_sizer.Add(btn_sizer, 0, wx.TOP | wx.RIGHT, 8)

        panel.SetSizer(main_sizer)

        # 編集/選択/クリックのイベント
        self.dvlc.Bind(dv.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.on_value_changed)
        self.dvlc.Bind(dv.EVT_DATAVIEW_SELECTION_CHANGED, self.on_selection_changed)
        self.dvlc.Bind(wx.EVT_LEFT_UP, self.on_left_up)
        btn_up.Bind(wx.EVT_BUTTON, self.on_move_up)
        btn_down.Bind(wx.EVT_BUTTON, self.on_move_down)

        self.load_csv()  # 起動時にCSV読み込み
        self.ensure_trailing_empty()

    def _get_app_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(sys.argv[0]))

    def load_csv(self):
        self._suspend_events = True
        self.dvlc.DeleteAllItems()

        rows = []
        if os.path.exists(self.csv_path):
            rows, encoding = self._read_csv_with_fallback()
            self.csv_encoding = encoding

        while rows and self._row_is_empty_values(rows[-1]):
            rows.pop()

        for row in rows:
            row = (row + ["", ""])[:2]
            self.dvlc.AppendItem(row)

        self._suspend_events = False

    def _read_csv_with_fallback(self):
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
        for row in range(self.dvlc.GetItemCount()):
            rows.append(
                [
                    self.dvlc.GetTextValue(row, 0),
                    self.dvlc.GetTextValue(row, 1),
                ]
            )
        # 末尾の空行は保存しない
        while rows and self._row_is_empty_values(rows[-1]):
            rows.pop()
        return rows

    def _row_is_empty(self, row):
        if row < 0 or row >= self.dvlc.GetItemCount():
            return True
        v1 = self.dvlc.GetTextValue(row, 0).strip()
        v2 = self.dvlc.GetTextValue(row, 1).strip()
        return v1 == "" and v2 == ""

    @staticmethod
    def _row_is_empty_values(values):
        if not values:
            return True
        v1 = (values[0] if len(values) > 0 else "").strip()
        v2 = (values[1] if len(values) > 1 else "").strip()
        return v1 == "" and v2 == ""

    def ensure_trailing_empty(self):
        count = self.dvlc.GetItemCount()

        # 末尾の空行が複数ある場合は1つにまとめる
        while count > 1 and self._row_is_empty(count - 1) and self._row_is_empty(count - 2):
            self.dvlc.DeleteItem(count - 1)
            count -= 1

        if count == 0 or not self._row_is_empty(count - 1):
            # 最終行は常に空行にする
            self.dvlc.AppendItem(["", ""])

    def on_value_changed(self, event):
        if self._suspend_events:
            return
        self.ensure_trailing_empty()
        self.save_csv()

    def on_selection_changed(self, event):
        self._copy_selected_text()

    def on_left_up(self, event):
        # クリック時にもコピーする
        self._copy_selected_text()
        event.Skip()

    def _copy_selected_text(self):
        row = self._get_selected_row()
        if row < 0:
            return
        text = self.dvlc.GetTextValue(row, 0)
        if text:
            self.copy_to_clipboard(text)

    def _get_selected_row(self):
        item = self.dvlc.GetSelection()
        if not item.IsOk():
            return -1
        return self.dvlc.ItemToRow(item)

    def on_move_up(self, event):
        row = self._get_selected_row()
        if row <= 0:
            return
        if self._row_is_empty(row):
            return
        self._swap_rows(row, row - 1)
        self._select_row(row - 1)
        self.ensure_trailing_empty()
        self.save_csv()

    def on_move_down(self, event):
        row = self._get_selected_row()
        count = self.dvlc.GetItemCount()
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
        values_a = [
            self.dvlc.GetTextValue(row_a, 0),
            self.dvlc.GetTextValue(row_a, 1),
        ]
        values_b = [
            self.dvlc.GetTextValue(row_b, 0),
            self.dvlc.GetTextValue(row_b, 1),
        ]

        # 値を書き換える間は変更イベントを止める
        self._suspend_events = True
        try:
            self.dvlc.SetTextValue(values_b[0], row_a, 0)
            self.dvlc.SetTextValue(values_b[1], row_a, 1)
            self.dvlc.SetTextValue(values_a[0], row_b, 0)
            self.dvlc.SetTextValue(values_a[1], row_b, 1)
        finally:
            self._suspend_events = False

    def _select_row(self, row):
        item = self.dvlc.RowToItem(row)
        if item.IsOk():
            self.dvlc.Select(item)
            self.dvlc.EnsureVisible(item)

    @staticmethod
    def copy_to_clipboard(text):
        if wx.TheClipboard.Open():
            try:
                wx.TheClipboard.SetData(wx.TextDataObject(text))
            finally:
                wx.TheClipboard.Close()


class CopyListApp(wx.App):
    def OnInit(self):
        frame = CopyListFrame()
        frame.CenterOnScreen()
        frame.Show()
        return True


if __name__ == "__main__":
    app = CopyListApp(False)
    app.MainLoop()
