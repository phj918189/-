import tkinter as tk
from tkinter import filedialog, messagebox
from tksheet import Sheet, rounded_box_coords
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
        self.root.title("엑셀 비교 뷰어")
        self.root.geometry("1700x850")

        self.df_left = None
        self.df_right = None
        self.path_left = None
        self.path_right = None
        self._syncing_scroll = False
        self._syncing_select = False

        # ─ 상단 버튼 ─
        top = tk.Frame(root)
        top.pack(pady=10)

        tk.Button(top, text="📂 첫 번째 파일", command=self.load_left,
                  bg="#3C91E6", fg="white", width=20).grid(row=0, column=0, padx=8)
        tk.Button(top, text="📂 두 번째 파일", command=self.load_right,
                  bg="#3C91E6", fg="white", width=20).grid(row=0, column=1, padx=8)
        tk.Button(top, text="🔍 비교 시작", command=self.compare_files,
                  bg="#FFC107", fg="black", width=20).grid(row=0, column=2, padx=8)
        
        # ─ 셀 주소 표시 라벨 ─
        self.cell_address_label = tk.Label(top, text="셀 주소: ", 
                                           font=("맑은 고딕", 10, "bold"), 
                                           bg="#F0F0F0", padx=10, pady=5)
        self.cell_address_label.grid(row=0, column=3, padx=8)

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
    # 🔹 엑셀 데이터 B2부터 불러오기 (A열은 빈 열로 유지)
    # ============================================
    def _load_b2_from_excel(self, path):
        """
        A열과 1행은 빈 셀, B2부터 데이터 시작
        A열도 포함해서 읽되, A열은 빈 값으로 유지
        1행도 빈 행으로 추가
        """
        wb = openpyxl.load_workbook(path, data_only=True)

        # 항상 첫 시트가 아닌 "채취일치" 시트를 사용
        ws = wb["채취일치"] if "채취일치" in wb.sheetnames else wb.active
        
        # 병합된 셀 정보 수집
        # 병합된 셀의 주 셀(첫 번째 셀)에만 값을 저장하고, 나머지 병합된 위치는 빈 값으로 처리
        merged_cell_map = {}  # Excel (행, 열) -> 값 (주 셀만 저장)
        merged_cell_ranges = {}  # Excel (행, 열) -> 병합 범위 정보 (주 셀만 저장)
        merged_cell_ignore = set()  # 병합된 셀 범위 내의 하위 셀들 (무시해야 할 위치)
        
        for merged_range in ws.merged_cells.ranges:
            # 병합 범위: min_row, min_col, max_row, max_col
            min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
            
            # 주 셀(첫 번째 셀)의 값 가져오기
            master_cell = ws.cell(min_row, min_col)
            master_value = master_cell.value if master_cell.value is not None else ""
            
            # 주 셀에만 값 저장 (B2부터만 처리)
            if min_row >= 2 and min_col >= 2:  # B2부터만 처리
                merged_cell_map[(min_row, min_col)] = str(master_value)
                merged_cell_ranges[(min_row, min_col)] = (min_row, min_col, max_row, max_col)
            
            # 병합된 셀 범위 내의 하위 셀들(주 셀 제외)은 무시 목록에 추가
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if r >= 2 and c >= 2:  # B2부터만 처리
                        # 주 셀이 아닌 경우에만 무시 목록에 추가
                        if not (r == min_row and c == min_col):
                            merged_cell_ignore.add((r, c))
        
        # Excel의 실제 행/열 범위 확인 (원본 구조 유지를 위해)
        excel_max_row = ws.max_row  # Excel의 최대 행 번호
        excel_max_col = ws.max_column  # Excel의 최대 열 번호
        
        data = []
        # Excel의 모든 행을 순회 (2행부터, 빈 행도 포함)
        for excel_row in range(2, excel_max_row + 1):
            values = []
            # A열(첫 번째 열)은 항상 빈 문자열로 처리
            values.append("")
            
            # B열부터 Excel의 최대 열까지 순회 (빈 열도 포함)
            for excel_col in range(2, excel_max_col + 1):  # 2=B열부터
                # 병합된 셀 범위 내의 하위 셀인지 확인 (무시해야 할 위치)
                if (excel_row, excel_col) in merged_cell_ignore:
                    # 병합된 셀의 하위 셀은 빈 값으로 처리 (원본 구조 유지)
                    values.append("")
                elif (excel_row, excel_col) in merged_cell_map:
                    # 병합된 셀의 주 셀인 경우, 저장된 값 사용
                    values.append(merged_cell_map[(excel_row, excel_col)])
                else:
                    # 일반 셀 값 (셀이 없거나 None이면 빈 문자열)
                    cell = ws.cell(excel_row, excel_col)
                    values.append(str(cell.value) if cell.value is not None else "")
            
            # 모든 행 추가 (빈 행도 포함하여 원본 구조 유지)
            data.append(values)
        
        # 열 개수 계산 (A열 + B열부터 최대 열까지 = excel_max_col)
        max_cols = excel_max_col  # A열(1개) + B열부터 최대 열까지(excel_max_col-1개) = excel_max_col개
        
        # DataFrame 생성 (빈 행/열도 모두 포함하여 원본 구조 유지)
        if data:
            df = pd.DataFrame(data)
            # 모든 열 유지 (빈 열도 포함하여 원본 구조 유지)
            # 열 개수가 부족하면 빈 열 추가
            if df.shape[1] < max_cols:
                # 부족한 열만큼 빈 열 추가
                for _ in range(max_cols - df.shape[1]):
                    df[len(df.columns)] = ""
        else:
            # 데이터가 없어도 Excel의 열 범위는 유지
            df = pd.DataFrame([[""] * max_cols])
        
        # 1행 추가 (모두 빈 값, A열 포함)
        empty_first_row = [""] * max_cols
        df = pd.concat([pd.DataFrame([empty_first_row], columns=df.columns), df], ignore_index=True)
        
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
    # 🔹 행/열 인덱스를 Excel 주소로 변환 (예: row=1, col=1 → "B2")
    # ============================================
    def _row_col_to_excel_address(self, row, col):
        """
        tksheet의 row/col 인덱스(0-based)를 Excel 주소로 변환
        - row=0 → Excel 행 1 (빈 행)
        - row=1 → Excel 행 2 (데이터 시작)
        - col=0 → Excel A열
        - col=1 → Excel B열
        """
        # Excel 행 번호 (1-based): 화면 row + 1
        excel_row = row + 1
        
        # Excel 열 이름 변환
        excel_col_num = col + 1  # Excel 열 번호 (1=A, 2=B, ...)
        n, s = excel_col_num, ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        
        return f"{s}{excel_row}"

    # ============================================
    # 🔹 셀 선택 동기화 + 주소 표시
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
                target.see(r, c)
            finally:
                self._syncing_select = False

        def update_cell_address(sheet, side):
            """선택된 셀의 Excel 주소를 업데이트"""
            try:
                selected = sheet.get_currently_selected()
                if selected and len(selected) >= 2:
                    r, c = selected[0], selected[1]
                    address = self._row_col_to_excel_address(r, c)
                    self.cell_address_label.config(text=f"셀 주소: {address} ({side})")
                else:
                    self.cell_address_label.config(text="셀 주소: -")
            except Exception:
                pass

        def on_left_select(event=None):
            sync_from_to(self.left_sheet, self.right_sheet)
            update_cell_address(self.left_sheet, "왼쪽")

        def on_right_select(event=None):
            sync_from_to(self.right_sheet, self.left_sheet)
            update_cell_address(self.right_sheet, "오른쪽")

        for ev in ("<ButtonRelease-1>", "<KeyRelease>", "<B1-Motion>"):
            self.left_sheet.MT.bind(ev, on_left_select, add=True)
            self.right_sheet.MT.bind(ev, on_right_select, add=True)

        for ev in ("cell_select", "row_select", "column_select", "shift_cell_select", "drag_select"):
            self.left_sheet.extra_bindings(ev, lambda p: on_left_select())
            self.right_sheet.extra_bindings(ev, lambda p: on_right_select())

    # ============================================
    # 🔹 파일 불러오기 (.xls 자동 변환 + B2 기준)
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

        self.path_left = path
        try:
            self.df_left = self._load_b2_from_excel(path)
        except Exception as e:
            messagebox.showerror("오류", f"왼쪽 파일 로드 실패:\n{e}")
            return

        # 줄바꿈, 긴 문장 셀도 전부 표시
        self.left_sheet.set_sheet_data(self.df_left.astype(str).values.tolist())
        self.left_sheet.set_all_cell_sizes_to_text()

        self.left_sheet.headers(self._generate_headers(self.df_left.shape[1]))
        self.left_sheet.enable_bindings(("single_select", "drag_select", "copy", "arrowkeys"))

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

        self.path_right = path
        try:
            self.df_right = self._load_b2_from_excel(path)
        except Exception as e:
            messagebox.showerror("오류", f"오른쪽 파일 로드 실패:\n{e}")
            return

        self.right_sheet.set_sheet_data(self.df_right.astype(str).values.tolist())
        self.right_sheet.set_all_cell_sizes_to_text()
        self.right_sheet.headers(self._generate_headers(self.df_right.shape[1]))
        self.right_sheet.enable_bindings(("single_select", "drag_select", "copy", "arrowkeys"))

    # ============================================
    # 🔹 열 헤더 자동 생성 (A,B,C,D... A열은 빈 열이지만 표시)
    # ============================================
    @staticmethod
    def _generate_headers(num_columns, excel_start_col=1):
        """
        Excel 열 헤더 생성 (예: A, B, C, D, ...)
        
        A열은 빈 열이지만 화면에는 표시되어야 함
        - DataFrame[0] = Excel A열 → 헤더 "A" (빈 열)
        - DataFrame[1] = Excel B열 → 헤더 "B" (데이터 시작)
        
        Args:
            num_columns: 생성할 헤더 개수 (DataFrame의 열 개수)
            excel_start_col: Excel 열 번호 (1=A, 2=B, 3=C, ...)
                            A열부터 표시하므로 기본값은 1
        """
        headers = []
        # Excel 열 번호를 Excel 열 이름으로 변환 (A=1, B=2, C=3, ...)
        for excel_col_num in range(excel_start_col, excel_start_col + num_columns):
            n, s = excel_col_num, ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                s = chr(65 + rem) + s
            headers.append(s)
        return headers

    # ============================================
    # 🔹 내부 비교 (노란색 표시)
    # ============================================
    def compare_files(self):
        if self.df_left is None or self.df_right is None:
            messagebox.showwarning("안내", "두 파일을 모두 불러온 후 비교를 진행하세요.")
            return

        rows = min(len(self.df_left), len(self.df_right))
        cols = min(len(self.df_left.columns), len(self.df_right.columns))
        diff_count = 0

        # 먼저 기존 강조 제거
        self.left_sheet.dehighlight_all()
        self.right_sheet.dehighlight_all()

        # A열과 1행은 비교하지 않음 (헤더 영역)
        for i in range(1, rows):
            for j in range(1, cols):
                # val_a = str(self.df_left.iat[i, j])
                # val_b = str(self.df_right.iat[i, j])
                # if val_a != val_b:

                # 공백,개행, 숫자포멧 불일치로 인한 "가짜" 차이 제거
                def norm(v):
                    s = "" if v is None else str(v)
                    s = s.replace("\r\n", "\n").replace("\r", "\n").strip()
                    try:
                        return str(float(s))
                    except:
                        return s

                val_a = norm(self.df_left.iat[i, j])
                val_b = norm(self.df_right.iat[i, j])
                if val_a != val_b:
                    diff_count += 1
                    self.left_sheet.highlight_cells(row=i, column=j, bg="green")  
                    self.right_sheet.highlight_cells(row=i, column=j, bg="green")

        messagebox.showinfo("비교 완료", f"값이 다른 셀 수: {diff_count}")


# ============================================
# 🔹 실행부
# ============================================
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparatorApp(root)
    root.mainloop()
