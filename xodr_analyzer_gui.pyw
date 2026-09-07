from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


class XodrAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XODR Analyzer v2")
        self.root.geometry("560x330")
        self.root.resizable(False, False)

        self.selected_path = tk.StringVar()

        title = tk.Label(
            root,
            text="OpenDRIVE XODR 자동 분석기 v2",
            font=("Malgun Gothic", 16, "bold"),
        )
        title.pack(pady=(25, 10))

        description = tk.Label(
            root,
            text=(
                "XODR에서 Excel, CSV, 그래프와 품질검사 결과를 "
                "자동 생성합니다."
            ),
            font=("Malgun Gothic", 10),
        )
        description.pack(pady=5)

        path_frame = tk.Frame(root)
        path_frame.pack(pady=15)

        path_entry = tk.Entry(
            path_frame,
            textvariable=self.selected_path,
            width=55,
        )
        path_entry.pack(side=tk.LEFT, padx=(0, 8))

        browse_button = tk.Button(
            path_frame,
            text="파일 선택",
            command=self.select_file,
            width=10,
        )
        browse_button.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(
            root,
            mode="indeterminate",
            length=460,
        )
        self.progress.pack(pady=15)

        self.status_label = tk.Label(
            root,
            text="XODR 파일을 선택하세요.",
            font=("Malgun Gothic", 10),
        )
        self.status_label.pack(pady=5)

        self.run_button = tk.Button(
            root,
            text="분석 시작",
            command=self.run_analysis,
            width=24,
            height=2,
            font=("Malgun Gothic", 11),
        )
        self.run_button.pack(pady=15)

    def select_file(self):
        selected = filedialog.askopenfilename(
            title="XODR 파일 선택",
            filetypes=[
                ("OpenDRIVE files", "*.xodr"),
                ("All files", "*.*"),
            ],
        )

        if selected:
            self.selected_path.set(selected)
            self.status_label.config(
                text="파일을 선택했습니다. 분석 시작을 누르세요."
            )

    def run_analysis(self):
        selected = self.selected_path.get().strip()

        if not selected:
            messagebox.showwarning(
                "파일 없음",
                "먼저 XODR 파일을 선택하세요.",
            )
            return

        xodr_path = Path(selected)

        if not xodr_path.exists():
            messagebox.showerror(
                "파일 오류",
                "선택한 파일을 찾을 수 없습니다.",
            )
            return

        self.run_button.config(state=tk.DISABLED)
        self.progress.start(10)
        self.status_label.config(text="분석 중입니다...")
        self.root.update_idletasks()

        analyzer = Path(__file__).with_name("xodr_analyzer.py")
        command = [sys.executable, str(analyzer), str(xodr_path)]

        try:
            completed = subprocess.run(
                command,
                check=False,
                creationflags=(
                    subprocess.CREATE_NO_WINDOW
                    if os.name == "nt"
                    else 0
                ),
            )

            result_folder = (
                xodr_path.parent
                / f"{xodr_path.stem}_analysis_result"
            )

            if completed.returncode == 0:
                self.status_label.config(text="분석 완료")
                messagebox.showinfo(
                    "분석 완료",
                    "분석이 정상적으로 완료되었습니다.\n\n"
                    f"결과 폴더:\n{result_folder}",
                )

                try:
                    os.startfile(result_folder)
                except Exception:
                    pass
            else:
                self.status_label.config(text="분석 실패")
                error_log = result_folder / "error.log"
                messagebox.showerror(
                    "분석 실패",
                    "분석 중 오류가 발생했습니다.\n\n"
                    f"오류 로그:\n{error_log}",
                )

        except Exception as error:
            self.status_label.config(text="실행 오류")
            messagebox.showerror("실행 오류", str(error))

        finally:
            self.progress.stop()
            self.run_button.config(state=tk.NORMAL)


root = tk.Tk()
app = XodrAnalyzerApp(root)
root.mainloop()
