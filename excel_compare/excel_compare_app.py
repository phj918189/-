import tkinter as tk
from tkinter import filedialog, messagebox
from tksheet import Sheet
import pandas as pd
import openpyxl
import xlrd
from openpyxl import Workbook
import os


# ============================================
# 🔹 오래된 .xls → .xlsx 자동 변환 함수
# ============================================
def convert_xls_to_xlsx(xls_path):
    """오래된 .xls 파일을 openpyxl에서 읽을 수 있는 .xlsx로 변환"""
    wb_xls = xlrd.open_workbook(xls_path)
    sheet = wb_xls.sheet_by_index(0)

    wb_xlsx = Workbook()
    ws_xlsx = wb_xlsx.active

    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            ws_xlsx.cell(row=r + 1, column=c + 1).value = sheet.cell_value(r, c)

    new_path = os.path.splitext(xls_path)[0] + "_converted.xlsx"
    wb_xlsx.save(new_path)
    return new_path


# ============================================
# 🔹 엑셀 비교 GUI
# ============================================
class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 비교 뷰어 (정확한 B2 기준)")
        self.root.geometry("1700x850")

        self.df_left = None
        self.df_right = None
        self._syncing_scroll = False
        self._syncing_select = False

        # ─ 상단 버튼 ─
        top = tk.Frame(root)
        top.pack(pady=10)

        tk.Button(top, text="📂 첫 번째 파일 열기", command=self.load_left,
                  bg="#3C91E6", fg="white", width=20).grid(row=0, column=0, padx=8)
        tk.Button(top, text="📂 두 번째 파일 열기", command=self.load_right,
                  bg="#3C91E6", fg="white", width=20).grid(row=0, column=1, padx=8)
        tk.Button(top, text="🔍 비교 시작", command=self.compare_files,
                  bg="#4CAF50", fg="white", width=18).grid(row=0, column=2, padx=8)

        # ─ 시트 영역 ─
        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.left_sheet = Sheet(frame, width=830, height=720)
        self.right_sheet = Sheet(frame, width=830, height=720)
        self.left_sheet.grid(row=0, column=0, sticky="nsew", padx=5)
        self.right_sheet.grid(row=0, column=1, sticky="nsew", padx=5)

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        # ─ 동기화 설정 ─
        self._install_sync_scroll()
        self._install_sync_selection()

    # ============================================
    # 🔹 엑셀 데이터 B2부터 불러오기
    # ============================================
    def _load_b2_from_excel(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active

        # 엑셀 좌표 기준으로 B2부터 읽기
        data = []
        for row in ws.iter_rows(min_row=2, min_col=2, values_only=True):
            values = [cell if cell is not None else "" for cell in row]
            # 완전히 빈 행 제외
            if any(values):
                data.append(values)

        if not data:
            raise ValueError("B2 이후에 데이터가 없습니다.")

        df = pd.DataFrame(data)
        # 완전히 빈 열 제거
        df = df.dropna(how="all", axis=1)
        return df.reset_index(drop=True)

    # ============================================
    # 🔹 스크롤 동기화
    # ============================================
    def _install_sync_scroll(self):
        def sync_y(src, tgt):
            def wrapper(*args):
                src(*args)
                try:
                    start, _ = src()
                    if not self._syncing_scroll:
                        self._syncing_scroll = True
                        tgt("moveto", start)
                        self._syncing_scroll = False
                except Exception:
                    pass
            return wrapper

        def sync_x(src, tgt):
            def wrapper(*args):
                src(*args)
                try:
                    start, _ = src()
                    if not self._syncing_scroll:
                        self._syncing_scroll = True
                        tgt("moveto", start)
                        self._syncing_scroll = False
                except Exception:
                    pass
            return wrapper

        self._left_yview_orig = self.left_sheet.MT.yview
        self._right_yview_orig = self.right_sheet.MT.yview
        self._left_xview_orig = self.left_sheet.MT.xview
        self._right_xview_orig = self.right_sheet.MT.xview

        self.left_sheet.MT.yview = sync_y(self._left_yview_orig, self.right_sheet.MT.yview_moveto)
        self.right_sheet.MT.yview = sync_y(self._right_yview_orig, self.left_sheet.MT.yview_moveto)
        self.left_sheet.MT.xview = sync_x(self._left_xview_orig, self.right_sheet.MT.xview_moveto)
        self.right_sheet.MT.xview = sync_x(self._right_xview_orig, self.left_sheet.MT.xview_moveto)

    # ============================================
    # 🔹 셀 선택 동기화
    # ============================================
    def _install_sync_selection(self):
        def sync_from_to(source, target):
            if self._syncing_select:
                return
            selected = source.get_currently_selected()
            if not selected or len(selected) < 2:
                return
            r, c = selected[0], selected[1]
            self._syncing_select = True
            try:
                target.select_cell(r, c, redraw=True)
                # 수동으로 셀이 화면에 보이도록 스크롤 조정 (see() 사용 시 버전별 None 처리 문제 회피)
                try:
                    top, bottom = target.MT.yview()
                    total_rows = max(1, target.MT.get_total_rows())
                    if not (top * total_rows <= r <= bottom * total_rows):
                        vis_rows = max(1, int((bottom - top) * total_rows))
                        frac = max(0.0, min(1.0, (r - vis_rows/2) / total_rows))
                        target.MT.yview_moveto(frac)
                    lft, rgt = target.MT.xview()
                    total_cols = max(1, target.MT.get_total_columns())
                    if not (lft * total_cols <= c <= rgt * total_cols):
                        vis_cols = max(1, int((rgt - lft) * total_cols))
                        fracx = max(0.0, min(1.0, (c - vis_cols/2) / total_cols))
                        target.MT.xview_moveto(fracx)
                except Exception:
                    pass
            finally:
                self._syncing_select = False

        def on_left_select(event=None):
            sync_from_to(self.left_sheet, self.right_sheet)

        def on_right_select(event=None):
            sync_from_to(self.right_sheet, self.left_sheet)

        for ev in ("<ButtonRelease-1>", "<KeyRelease>", "<B1-Motion>"):
            self.left_sheet.MT.bind(ev, on_left_select, add=True)
            self.right_sheet.MT.bind(ev, on_right_select, add=True)

    # ============================================
    # 🔹 파일 불러오기 (.xls 자동 변환 + B2기준)
    # ============================================
    def load_left(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if not path:
            return

        if path.lower().endswith(".xls"):
            try:
                path = convert_xls_to_xlsx(path)
            except Exception as e:
                messagebox.showerror("오류", f"파일 변환 실패:\n{e}")
                return

        try:
            self.df_left = self._load_b2_from_excel(path)
        except Exception as e:
            messagebox.showerror("오류", f"왼쪽 파일 로드 실패:\n{e}")
            return

        self.left_sheet.set_sheet_data(self.df_left.astype(str).values.tolist())
        self.left_sheet.headers(self._generate_headers(self.df_left.shape[1]))
        self.left_sheet.enable_bindings((
            "single_select", "drag_select", "copy", "arrowkeys"
        ))
        # 자동 리사이즈 비활성: 원본 크기 유지

    def load_right(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if not path:
            return

        if path.lower().endswith(".xls"):
            try:
                path = convert_xls_to_xlsx(path)
            except Exception as e:
                messagebox.showerror("오류", f"파일 변환 실패:\n{e}")
                return

        try:
            self.df_right = self._load_b2_from_excel(path)
        except Exception as e:
            messagebox.showerror("오류", f"오른쪽 파일 로드 실패:\n{e}")
            return

        self.right_sheet.set_sheet_data(self.df_right.astype(str).values.tolist())
        self.right_sheet.headers(self._generate_headers(self.df_right.shape[1]))
        self.right_sheet.enable_bindings((
            "single_select", "drag_select", "copy", "arrowkeys"
        ))
        # 자동 리사이즈 비활성: 원본 크기 유지

    # ============================================
    # 🔹 열 헤더 자동 생성 (A,B,C...)
    # ============================================
    @staticmethod
    def _generate_headers(num_columns):
        headers = []
        for i in range(1, num_columns + 1):
            n, s = i, ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                s = chr(65 + rem) + s
            headers.append(s)
        return headers

    # ============================================
    # 🔹 비교 실행
    # ============================================
    def compare_files(self):
        if self.df_left is None or self.df_right is None:
            messagebox.showwarning("안내", "두 파일을 모두 불러온 후 비교를 진행하세요.")
            return

        rows = min(len(self.df_left), len(self.df_right))
        cols = min(len(self.df_left.columns), len(self.df_right.columns))
        diff_count = 0

        for i in range(rows):
            for j in range(cols):
                val_a = str(self.df_left.iat[i, j])
                val_b = str(self.df_right.iat[i, j])
                if val_a != val_b:
                    diff_count += 1
                    self.left_sheet.highlight_cells(row=i, column=j, bg="lightgreen")
                    self.right_sheet.highlight_cells(row=i, column=j, bg="lightgreen")

        messagebox.showinfo("비교 완료", f"값이 다른 셀 수: {diff_count}")


# ============================================
# 🔹 실행부
# ============================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparatorApp(root)
    root.mainloop()
