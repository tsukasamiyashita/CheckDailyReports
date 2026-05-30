# tsukasamiyashita/checkdailyreports/CheckDailyReports-94e7e606e86357cba1e8adce22bf563302ea0859/app.py
import os
import re
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import xlrd

class DailyReportCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CheckDailyReports-v1.0.0")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)

        # スタイル設定
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # 配色定義
        self.bg_color = "#f3f4f6"
        self.primary_color = "#2563eb"
        self.accent_color = "#dc2626"
        
        self.root.configure(bg=self.bg_color)
        
        # 状態保持変数
        self.target_dir = tk.StringVar()
        self.is_processing = False
        self.cancel_requested = False

        self._build_ui()

    def _build_ui(self):
        # メインコンテナ
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ヘッダーエリア
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = tk.Label(
            header_frame, 
            text="日報入力ミス自動チェックツール", 
            font=("Helvetica", 16, "bold"),
            bg=self.bg_color,
            fg="#1f2937"
        )
        title_label.pack(side=tk.LEFT)
        
        ver_label = tk.Label(
            header_frame,
            text="v1.0.0",
            font=("Helvetica", 10, "italic"),
            bg=self.bg_color,
            fg="#6b7280"
        )
        ver_label.pack(side=tk.LEFT, padx=10, pady=(5, 0))

        # フォルダ選択エリア
        folder_frame = ttk.LabelFrame(main_frame, text=" 1. 対象フォルダ選択 ", padding=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.target_dir, font=("Helvetica", 10))
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(folder_frame, text="参照...", command=self._browse_folder)
        browse_btn.pack(side=tk.RIGHT)

        # 動作ボタンエリア
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = tk.Button(
            btn_frame, 
            text="チェック開始", 
            bg=self.primary_color, 
            fg="white", 
            font=("Helvetica", 11, "bold"),
            activebackground="#1d4ed8",
            activeforeground="white",
            relief=tk.FLAT,
            padx=15,
            pady=6,
            command=self._start_check_thread
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = tk.Button(
            btn_frame, 
            text="中止", 
            bg="#9ca3af", 
            fg="white", 
            font=("Helvetica", 11, "bold"),
            activebackground="#78716c",
            activeforeground="white",
            relief=tk.FLAT,
            state=tk.DISABLED,
            padx=15,
            pady=6,
            command=self._stop_check
        )
        self.stop_btn.pack(side=tk.LEFT)

        # 進捗バーと進捗テキスト
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 2))
        
        self.status_label = tk.Label(
            self.progress_frame, 
            text="フォルダを選択して「チェック開始」をクリックしてください。", 
            font=("Helvetica", 9),
            bg=self.bg_color,
            fg="#4b5563",
            anchor="w"
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

        # 結果表示エリア（リスト）
        list_frame = ttk.LabelFrame(main_frame, text=" 2. チェック結果一覧 (エラー検出箇所) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # スクロールバー
        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)

        # Treeview（カラム設定）
        cols = ("file_path", "sheet_name", "cell_pos", "error_type", "detail")
        self.tree = ttk.Treeview(
            list_frame, 
            columns=cols, 
            show="headings", 
            yscrollcommand=scroll_y.set, 
            xscrollcommand=scroll_x.set
        )
        
        scroll_y.config(command=self.tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.config(command=self.tree.xview)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # カラムヘッダー設定
        self.tree.heading("file_path", text="ファイル名 / パス")
        self.tree.heading("sheet_name", text="シート名")
        self.tree.heading("cell_pos", text="セル位置")
        self.tree.heading("error_type", text="エラー区分")
        self.tree.heading("detail", text="エラー詳細内容")

        # カラム幅調整
        self.tree.column("file_path", width=250, anchor="w")
        self.tree.column("sheet_name", width=100, anchor="center")
        self.tree.column("cell_pos", width=80, anchor="center")
        self.tree.column("error_type", width=120, anchor="center")
        self.tree.column("detail", width=300, anchor="w")

        # ダブルクリックイベント（詳細表示）
        self.tree.bind("<Double-1>", self._show_detail_popup)

    def _browse_folder(self):
        selected = filedialog.askdirectory()
        if selected:
            self.target_dir.set(os.path.abspath(selected))

    def _stop_check(self):
        if self.is_processing:
            self.cancel_requested = True
            self.status_label.config(text="中止処理中...")
            self.stop_btn.config(state=tk.DISABLED)

    def _start_check_thread(self):
        target = self.target_dir.get().strip()
        if not target:
            messagebox.showwarning("警告", "対象フォルダが選択されていません。")
            return
        if not os.path.exists(target):
            messagebox.showerror("エラー", "選択されたフォルダが存在しません。")
            return

        self.is_processing = True
        self.cancel_requested = False
        self.start_btn.config(state=tk.DISABLED, bg="#9ca3af")
        self.stop_btn.config(state=tk.NORMAL, bg=self.accent_color)
        
        # リストクリア
        for item in self.tree.get_children():
            self.tree.delete(item)

        # スレッド起動
        thread = threading.Thread(target=self._run_checker, args=(target,), daemon=True)
        thread.start()

    def _run_checker(self, target_dir):
        # 対象ファイル収集 (xls, xlsx, xlsm)
        valid_extensions = (".xls", ".xlsx", ".xlsm")
        files_to_check = []
        for root_path, _, files in os.walk(target_dir):
            for file in files:
                if file.lower().endswith(valid_extensions) and not file.startswith("~$"):
                    files_to_check.append(os.path.join(root_path, file))

        total_files = len(files_to_check)
        if total_files == 0:
            self.root.after(0, self._finish_checker, 0, "対象のExcelファイルが見つかりませんでした。")
            return

        errors_found = 0
        for i, filepath in enumerate(files_to_check):
            if self.cancel_requested:
                break

            # 進捗更新
            progress_percent = int(((i + 1) / total_files) * 100)
            relative_name = os.path.relpath(filepath, target_dir)
            self.root.after(
                0, 
                self._update_progress, 
                progress_percent, 
                f"検証中 ({i+1}/{total_files}): {relative_name}"
            )

            # チェック実施
            errors = self._check_excel_file(filepath)
            if errors:
                errors_found += len(errors)
                self.root.after(0, self._add_errors_to_list, filepath, errors)

        msg = "チェックを中止しました。" if self.cancel_requested else f"チェック完了。全 {total_files} ファイル中 {errors_found} 件のエラーを検出しました。"
        self.root.after(0, self._finish_checker, errors_found, msg)

    def _update_progress(self, percent, text):
        self.progress_bar["value"] = percent
        self.status_label.config(text=text)

    def _add_errors_to_list(self, filepath, errors):
        for err in errors:
            self.tree.insert("", tk.END, values=(
                os.path.basename(filepath),
                err.get("sheet", "不明"),
                err.get("cell", "N/A"),
                err.get("type", "警告"),
                err.get("detail", "")
            ), tags=(filepath,))

    def _finish_checker(self, errors_found, message):
        self.is_processing = False
        self.progress_bar["value"] = 100 if not self.cancel_requested else 0
        self.status_label.config(text=message)
        self.start_btn.config(state=tk.NORMAL, bg=self.primary_color)
        self.stop_btn.config(state=tk.DISABLED, bg="#9ca3af")
        
        if self.cancel_requested:
            messagebox.showinfo("中止", "処理を中断しました。")
        elif errors_found > 0:
            messagebox.showwarning("チェック完了", f"入力ミスが {errors_found} 件検出されました。リストを確認してください。")
        else:
            messagebox.showinfo("チェック完了", "入力ミスは検出されませんでした。")

    def _show_detail_popup(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        values = self.tree.item(selected_item[0], "values")
        if len(values) < 5:
            return
            
        popup = tk.Toplevel(self.root)
        popup.title("エラー詳細情報")
        popup.geometry("550x350")
        popup.transient(self.root)
        popup.grab_set()

        frm = ttk.Frame(popup, padding=15)
        frm.pack(fill=tk.BOTH, expand=True)

        labels = [
            ("ファイル名:", values[0]),
            ("シート名:", values[1]),
            ("セル位置:", values[2]),
            ("エラー種別:", values[3])
        ]

        for i, (lbl, val) in enumerate(labels):
            tk.Label(frm, text=lbl, font=("Helvetica", 10, "bold"), anchor="w").grid(row=i, column=0, sticky="nw", pady=5)
            tk.Label(frm, text=val, font=("Helvetica", 10), justify="left", anchor="w", wraplength=400).grid(row=i, column=1, sticky="nw", pady=5, padx=10)

        tk.Label(frm, text="詳細説明:", font=("Helvetica", 10, "bold"), anchor="w").grid(row=4, column=0, sticky="nw", pady=5)
        
        detail_txt = tk.Text(frm, font=("Helvetica", 10), height=6, wrap=tk.WORD, bg="#f9fafb", relief=tk.SOLID, bd=1)
        detail_txt.insert(tk.END, values[4])
        detail_txt.config(state=tk.DISABLED)
        detail_txt.grid(row=4, column=1, sticky="nsew", pady=5, padx=10)

        frm.rowconfigure(4, weight=1)
        frm.columnconfigure(1, weight=1)

        close_btn = ttk.Button(frm, text="閉じる", command=popup.destroy)
        close_btn.grid(row=5, column=0, columnspan=2, pady=(15, 0))

    # ==========================================
    # ヘルパーメソッド
    # ==========================================
    def _should_check_sheet(self, sheet_name):
        # 設定やマスタ、コード定義用シートなどは検証対象外にする
        exclude_keywords = ["設定", "コード", "マスタ", "master", "list", "リスト", "summary", "集計"]
        name_lower = sheet_name.lower()
        return not any(k in name_lower for k in exclude_keywords)

    def _get_clean_value(self, val):
        if val is None:
            return ""
        if isinstance(val, float):
            # 浮動小数値かつ実質整数の場合はintキャストして文字列化 (20000000.0 -> "20000000")
            if val.is_integer():
                return str(int(val))
            return str(val)
        return str(val).strip()

    def _find_code_columns(self, get_val_func, nrows, ncols):
        """見出し行を特定し、管理No, 型枠, 作業コードの列番号を返す共通関数"""
        col_mng_no = 1  # デフォルト B列
        col_katawaku = 9  # デフォルト J列
        col_sagyo = 11  # デフォルト L列
        
        mng_keywords = ["管理no", "管理番号", "管理№", "工事番号", "工事no", "工事№", "管no", "管理ナンバー"]
        katawaku_keywords = ["型枠コード", "型枠cd", "型枠id", "型枠番号", "型枠"]
        sagyo_keywords = ["作業コード", "作業cd", "作業id", "作業番号", "作業"]
        
        found_header = False
        # パス1: 同一行に管理Noとコード群が揃っている本物の見出し行を探索
        for r in range(min(nrows, 100)):
            temp_mng = -1
            temp_kat = -1
            temp_sag = -1
            for c in range(ncols):
                val_str = get_val_func(r, c).lower().replace(" ", "").replace("　", "")
                if not val_str:
                    continue
                    
                if temp_mng == -1 and any(val_str == k or val_str.startswith(k) for k in mng_keywords + ["管理"]):
                    temp_mng = c
                if temp_kat == -1 and any(val_str == k or val_str.startswith(k) for k in katawaku_keywords):
                    temp_kat = c
                if temp_sag == -1 and any(val_str == k or val_str.startswith(k) for k in sagyo_keywords):
                    temp_sag = c
                    
            if temp_mng != -1 and (temp_kat != -1 or temp_sag != -1):
                col_mng_no = temp_mng
                if temp_kat != -1: col_katawaku = temp_kat
                if temp_sag != -1: col_sagyo = temp_sag
                found_header = True
                break
                
        # パス2: 揃っていない場合、管理Noが単独で存在する行を見出し行とする
        if not found_header:
            for r in range(min(nrows, 100)):
                for c in range(ncols):
                    val_str = get_val_func(r, c).lower().replace(" ", "").replace("　", "")
                    if any(val_str == k or val_str.startswith(k) for k in mng_keywords + ["管理"]):
                        col_mng_no = c
                        found_header = True
                        break
                if found_header:
                    break

        return col_mng_no, col_katawaku, col_sagyo

    # ==========================================
    # エクセル解析エンジン (絶対保存・上書きしない)
    # ==========================================
    def _check_excel_file(self, filepath):
        errors = []
        ext = os.path.splitext(filepath)[1].lower()

        try:
            if ext == ".xls":
                errors.extend(self._parse_xls(filepath))
            elif ext in (".xlsx", ".xlsm"):
                errors.extend(self._parse_xlsx_xlsm(filepath))
        except Exception as e:
            errors.append({
                "sheet": "ファイル全体",
                "cell": "N/A",
                "type": "読込失敗",
                "detail": f"Excelファイルを開くことができませんでした。破損またはパスワード保護の可能性があります。 (エラー: {str(e)})"
            })
        return errors

    def _parse_xls(self, filepath):
        errors = []
        wb = xlrd.open_workbook(filepath, formatting_info=False)
        for sheet_index in range(wb.nsheets):
            sheet = wb.sheet_by_index(sheet_index)
            sheet_name = sheet.name
            
            if sheet.nrows == 0 or sheet.ncols == 0:
                continue

            # スキャン対象シートかどうかの検証
            if not self._should_check_sheet(sheet_name):
                continue

            self._scan_sheet_data_xls(sheet, sheet_name, errors)
            self._check_code_integrity_xls(sheet, sheet_name, errors)
        return errors

    def _parse_xlsx_xlsm(self, filepath):
        errors = []
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            if sheet.max_row is None or sheet.max_row == 0:
                continue
                
            # スキャン対象シートかどうかの検証
            if not self._should_check_sheet(sheet_name):
                continue

            self._scan_sheet_data_xlsx(sheet, sheet_name, errors)
            self._check_code_integrity_xlsx(sheet, sheet_name, errors)
            
        wb.close()
        return errors

    # ------------------------------------------
    # 管理Noと型枠/作業コードの未入力整合性検証 (.xls 用)
    # ------------------------------------------
    def _check_code_integrity_xls(self, sheet, sheet_name, errors):
        nrows = sheet.nrows
        ncols = sheet.ncols
        
        # 見出し列の確実な特定
        col_mng_no, col_katawaku, col_sagyo = self._find_code_columns(
            lambda r, c: self._get_clean_value(sheet.cell_value(r, c)), 
            nrows, 
            ncols
        )

        for r in range(nrows):
            curr_mng_col = col_mng_no
            mng_val = ""
            if curr_mng_col < ncols:
                mng_val = self._get_clean_value(sheet.cell_value(r, curr_mng_col))
                
            is_valid_mng = False
            if mng_val and mng_val != "None" and mng_val != "":
                if mng_val.isdigit():
                    val_num = int(mng_val)
                    if not (40000 <= val_num <= 50000): # 日付シリアル値除外
                        is_valid_mng = True
                elif len(mng_val) >= 4:
                    is_valid_mng = True
                        
            # 自動探索フォールバック
            if not is_valid_mng and ncols > 1:
                max_search_col = min(9, ncols)
                for temp_c in range(1, max_search_col):
                    temp_val = self._get_clean_value(sheet.cell_value(r, temp_c))
                    if temp_val and temp_val.isdigit():
                        val_num = int(temp_val)
                        if len(temp_val) >= 4 and not (40000 <= val_num <= 50000):
                            mng_val = temp_val
                            curr_mng_col = temp_c
                            is_valid_mng = True
                            break
                            
            if is_valid_mng:
                kat_val = ""
                if col_katawaku < ncols:
                    kat_val = self._get_clean_value(sheet.cell_value(r, col_katawaku))
                        
                sag_val = ""
                if col_sagyo < ncols:
                    sag_val = self._get_clean_value(sheet.cell_value(r, col_sagyo))
                        
                is_kat_empty = (kat_val == "" or kat_val == "None" or kat_val == "0")
                is_sag_empty = (sag_val == "" or sag_val == "None" or sag_val == "0")
                
                # 型枠コード、あるいは作業コードの「両方とも」が未入力の場合にエラーとする
                if is_kat_empty and is_sag_empty:
                    cell_pos_str = f"{xlrd.formula.colname(curr_mng_col)}{r + 1}"
                        
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "コード未入力エラー",
                        "detail": f"管理No '{mng_val}' が指定されていますが、型枠コードと作業コードの両方が入力されていません。"
                    })

    # ------------------------------------------
    # 管理Noと型枠/作業コードの未入力整合性検証 (.xlsx 用)
    # ------------------------------------------
    def _check_code_integrity_xlsx(self, sheet, sheet_name, errors):
        rows = list(sheet.iter_rows(max_row=500, max_col=50, values_only=True))
        if not rows:
            return
            
        nrows = len(rows)
        ncols = max(len(row) for row in rows) if nrows > 0 else 0
        
        def _get_val(r, c):
            if r < len(rows) and c < len(rows[r]):
                return self._get_clean_value(rows[r][c])
            return ""

        # 見出し列の確実な特定
        col_mng_no, col_katawaku, col_sagyo = self._find_code_columns(_get_val, nrows, ncols)

        for r_idx, row in enumerate(rows):
            curr_mng_col = col_mng_no
            mng_val = ""
            if curr_mng_col < len(row):
                mng_val = self._get_clean_value(row[curr_mng_col])
                
            is_valid_mng = False
            if mng_val and mng_val != "None" and mng_val != "":
                if mng_val.isdigit():
                    val_num = int(mng_val)
                    if not (40000 <= val_num <= 50000):
                        is_valid_mng = True
                elif len(mng_val) >= 4:
                    is_valid_mng = True
                        
            # 自動探索フォールバック
            if not is_valid_mng and len(row) > 1:
                max_search_col = min(9, len(row))
                for temp_c in range(1, max_search_col):
                    temp_val = self._get_clean_value(row[temp_c])
                    if temp_val and temp_val.isdigit():
                        val_num = int(temp_val)
                        if len(temp_val) >= 4 and not (40000 <= val_num <= 50000):
                            mng_val = temp_val
                            curr_mng_col = temp_c
                            is_valid_mng = True
                            break
                            
            if is_valid_mng:
                kat_val = ""
                if col_katawaku < len(row):
                    kat_val = self._get_clean_value(row[col_katawaku])
                        
                sag_val = ""
                if col_sagyo < len(row):
                    sag_val = self._get_clean_value(row[col_sagyo])
                        
                is_kat_empty = (kat_val == "" or kat_val == "None" or kat_val == "0")
                is_sag_empty = (sag_val == "" or sag_val == "None" or sag_val == "0")
                
                # 型枠コード、あるいは作業コードの「両方とも」が未入力の場合にエラーとする
                if is_kat_empty and is_sag_empty:
                    cell_pos_str = f"{openpyxl.utils.get_column_letter(curr_mng_col + 1)}{r_idx + 1}"
                        
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "コード未入力エラー",
                        "detail": f"管理No '{mng_val}' が指定されていますが、型枠コードと作業コードの両方が入力されていません。"
                    })

    # ------------------------------------------
    # 汎用スマートスキャンルールエンジン
    # ------------------------------------------
    def _scan_sheet_data_xls(self, sheet, sheet_name, errors):
        nrows = sheet.nrows
        ncols = sheet.ncols
        
        max_r = min(nrows, 200)
        max_c = min(ncols, 50)

        for r in range(max_r):
            for c in range(max_c):
                cell_val = sheet.cell_value(r, c)
                if not isinstance(cell_val, str):
                    continue
                
                clean_val = cell_val.replace(" ", "").replace("　", "")
                
                if any(k in clean_val for k in ["氏名", "担当者", "名前", "記述者", "作成者"]) and "印" not in clean_val:
                    self._validate_neighbor_xls(sheet, sheet_name, r, c, "氏名", errors)

                if any(k in clean_val for k in ["日付", "年月日", "作成日", "報告日"]):
                    self._validate_neighbor_xls(sheet, sheet_name, r, c, "日付", errors)

                if any(k == clean_val for k in ["時間", "工数", "作業時間", "h", "H"]):
                    self._validate_table_column_xls(sheet, sheet_name, r, c, errors)

    def _scan_sheet_data_xlsx(self, sheet, sheet_name, errors):
        for r_idx, row in enumerate(sheet.iter_rows(max_row=200, max_col=50, values_only=True), start=1):
            for c_idx, cell_val in enumerate(row, start=1):
                if cell_val is None or not isinstance(cell_val, str):
                    continue
                
                clean_val = cell_val.replace(" ", "").replace("　", "")
                
                if any(k in clean_val for k in ["氏名", "担当者", "名前", "記述者", "作成者"]) and "印" not in clean_val:
                    self._validate_neighbor_xlsx(sheet, sheet_name, r_idx, c_idx, "氏名", errors)

                if any(k in clean_val for k in ["日付", "年月日", "作成日", "報告日"]):
                    self._validate_neighbor_xlsx(sheet, sheet_name, r_idx, c_idx, "日付", errors)

                if any(k == clean_val for k in ["時間", "工数", "作業時間", "h", "H"]):
                    self._validate_table_column_xlsx(sheet, sheet_name, r_idx, c_idx, errors)

    def _validate_neighbor_xls(self, sheet, sheet_name, r, c, item_type, errors):
        candidates = []
        if c + 1 < sheet.ncols:
            candidates.append((r, c + 1, "右隣"))
        if r + 1 < sheet.nrows:
            candidates.append((r + 1, c, "下"))

        validated = False
        for target_r, target_c, direction in candidates:
            val = sheet.cell_value(target_r, target_c)
            if val is not None and str(val).strip() != "":
                val_str = str(val).strip()
                cell_pos_str = f"{xlrd.formula.colname(target_c)}{target_r + 1}"
                
                if item_type == "日付":
                    if sheet.cell_type(target_r, target_c) == xlrd.XL_CELL_DATE:
                        pass
                    else:
                        if not self._is_valid_date_string(val_str):
                            errors.append({
                                "sheet": sheet_name,
                                "cell": cell_pos_str,
                                "type": "日付フォーマットエラー",
                                "detail": f"「{item_type}」セルの{direction}に有効な日付が入力されていません。入力値: '{val_str}'"
                            })
                validated = True
                break
                
        if not validated and candidates:
            col_letter = xlrd.formula.colname(c)
            errors.append({
                "sheet": sheet_name,
                "cell": f"{col_letter}{r + 1}",
                "type": "未入力項目",
                "detail": f"「{item_type}」セルの周辺（右隣または下）に有効な値がありません。入力が漏れている可能性があります。"
            })

    def _validate_table_column_xls(self, sheet, sheet_name, header_r, header_c, errors):
        consecutive_empty = 0
        r = header_r + 1
        
        while r < sheet.nrows and consecutive_empty < 3:
            val = sheet.cell_value(r, header_c)
            cell_pos_str = f"{xlrd.formula.colname(header_c)}{r + 1}"
            
            other_data_filled = False
            for col_idx in range(sheet.ncols):
                if col_idx != header_c:
                    other_val = sheet.cell_value(r, col_idx)
                    if other_val is not None and str(other_val).strip() != "":
                        other_data_filled = True
                        break
            
            if val is None or str(val).strip() == "":
                if other_data_filled:
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "工数未入力",
                        "detail": "業務記述行ですが、作業時間(工数)が空欄になっています。"
                    })
                consecutive_empty += 1
            else:
                consecutive_empty = 0
                try:
                    num_val = float(val)
                    if num_val < 0:
                        errors.append({
                            "sheet": sheet_name,
                            "cell": cell_pos_str,
                            "type": "工数負数エラー",
                            "detail": f"作業時間に負の値が入力されています。入力値: {num_val}"
                        })
                    elif num_val > 24:
                        errors.append({
                            "sheet": sheet_name,
                            "cell": cell_pos_str,
                            "type": "工数過大エラー",
                            "detail": f"1日の作業工数として過大な時間(24h超)が入力されています。入力値: {num_val}"
                        })
                except ValueError:
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "数値フォーマットエラー",
                        "detail": f"時間入力欄に数字以外の値が入力されています。入力値: '{val}'"
                    })
            r += 1

    def _validate_neighbor_xlsx(self, sheet, sheet_name, r, c, item_type, errors):
        candidates = []
        candidates.append((r, c + 1, "右隣"))
        candidates.append((r + 1, c, "下"))

        validated = False
        for target_r, target_c, direction in candidates:
            val = sheet.cell(row=target_r, column=target_c).value
            if val is not None and str(val).strip() != "":
                val_str = str(val).strip()
                cell_pos_str = f"{openpyxl.utils.get_column_letter(target_c)}{target_r}"
                
                if item_type == "日付":
                    if isinstance(val, (datetime, tk.Variable)):
                        pass
                    else:
                        if not self._is_valid_date_string(val_str):
                            errors.append({
                                "sheet": sheet_name,
                                "cell": cell_pos_str,
                                "type": "日付フォーマットエラー",
                                "detail": f"「{item_type}」セルの{direction}に有効な日付が入力されていません。入力値: '{val_str}'"
                            })
                validated = True
                break

        if not validated:
            col_letter = openpyxl.utils.get_column_letter(c)
            errors.append({
                "sheet": sheet_name,
                "cell": f"{col_letter}{r}",
                "type": "未入力項目",
                "detail": f"「{item_type}」セルの周辺（右隣または下）に有効なデータが入力されていません。"
            })

    def _validate_table_column_xlsx(self, sheet, sheet_name, header_r, header_c, errors):
        consecutive_empty = 0
        r = header_r + 1
        
        while consecutive_empty < 3:
            val = sheet.cell(row=r, column=header_c).value
            cell_pos_str = f"{openpyxl.utils.get_column_letter(header_c)}{r}"
            
            other_data_filled = False
            for col_idx in range(1, 20):
                if col_idx != header_c:
                    other_val = sheet.cell(row=r, column=col_idx).value
                    if other_val is not None and str(other_val).strip() != "":
                        other_data_filled = True
                        break
            
            if val is None or str(val).strip() == "":
                if other_data_filled:
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "工数未入力",
                        "detail": "業務記述行ですが、作業時間(工数)が空欄になっています。"
                    })
                consecutive_empty += 1
            else:
                consecutive_empty = 0
                try:
                    num_val = float(val)
                    if num_val < 0:
                        errors.append({
                            "sheet": sheet_name,
                            "cell": cell_pos_str,
                            "type": "工数負数エラー",
                            "detail": f"作業時間に負の値が入力されています。入力値: {num_val}"
                        })
                    elif num_val > 24:
                        errors.append({
                            "sheet": sheet_name,
                            "cell": cell_pos_str,
                            "type": "工数過大エラー",
                            "detail": f"1日の作業工数として過大な時間(24h超)が入力されています。入力値: {num_val}"
                        })
                except ValueError:
                    errors.append({
                        "sheet": sheet_name,
                        "cell": cell_pos_str,
                        "type": "数値フォーマットエラー",
                        "detail": f"時間入力欄に数字以外の値が入力されています。入力値: '{val}'"
                    })
            r += 1

    def _is_valid_date_string(self, text):
        patterns = [
            r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$",
            r"^\d{1,2}[-/]\d{1,2}$",
            r"^\d{4}年\d{1,2}月\d{1,2}日$",
            r"^\d{1,2}月\d{1,2}日$"
        ]
        text_clean = text.strip()
        for p in patterns:
            if re.match(p, text_clean):
                return True
        return False

if __name__ == "__main__":
    root = tk.Tk()
    app = DailyReportCheckerApp(root)
    root.mainloop()