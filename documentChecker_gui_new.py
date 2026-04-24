from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import documentChecker


class CheckerGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("documentChecker GUI")
        self.root.geometry("760x280")
        self.root.minsize(700, 250)

        self.target_folder_var = tk.StringVar()
        self.output_file_var = tk.StringVar(value="review_results.xlsx")
        self.cover_keyword_var = tk.StringVar()
        self.status_var = tk.StringVar(value="待機中")
        self.cancel_event = threading.Event()
        self.was_cancelled = False

        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="対象フォルダ").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.target_folder_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        tk.Button(frame, text="参照", command=self._select_target_folder).grid(row=0, column=2, sticky="ew")

        tk.Label(frame, text="出力ファイル (xlsx)").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.output_file_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        tk.Button(frame, text="保存先", command=self._select_output_file).grid(row=1, column=2, sticky="ew", pady=(8, 0))

        tk.Label(frame, text="表紙キーワード (任意)").grid(row=2, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.cover_keyword_var).grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=(8, 0))

        self.run_button = tk.Button(frame, text="実行", command=self._run_check)
        self.run_button.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(14, 0))

        self.cancel_button = tk.Button(frame, text="中止", command=self._cancel_check, state=tk.DISABLED)
        self.cancel_button.grid(row=3, column=2, sticky="ew", pady=(14, 0))

        tk.Label(frame, textvariable=self.status_var, anchor="w", fg="#0b4f6c").grid(
            row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0)
        )

        frame.columnconfigure(1, weight=1)

    def _select_target_folder(self) -> None:
        path = filedialog.askdirectory(title="チェック対象フォルダを選択")
        if path:
            self.target_folder_var.set(path)
            if not self.output_file_var.get().strip():
                self.output_file_var.set(str(Path(path) / "review_results.xlsx"))

    def _select_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="出力ファイルを選択",
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
            initialfile="review_results.xlsx",
        )
        if path:
            self.output_file_var.set(path)

    def _run_check(self) -> None:
        target_folder = self.target_folder_var.get().strip()
        output_file = self.output_file_var.get().strip()
        cover_keyword = self.cover_keyword_var.get().strip()

        if not target_folder:
            messagebox.showerror("入力エラー", "対象フォルダを指定してください。")
            return

        if not Path(target_folder).exists():
            messagebox.showerror("入力エラー", "対象フォルダが存在しません。")
            return

        if not output_file:
            messagebox.showerror("入力エラー", "出力ファイル名を指定してください。")
            return

        argv = [target_folder, "-0", output_file]
        if cover_keyword:
            argv.extend(["--cover-keyword", cover_keyword])

        self.run_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self.cancel_event.clear()
        self.was_cancelled = False
        self.status_var.set("0/0/0/0（完了/失敗/処理中/母数）")

        worker = threading.Thread(target=self._run_worker, args=(argv,), daemon=True)
        worker.start()

    def _cancel_check(self) -> None:
        self.cancel_event.set()
        self.status_var.set("中止要求中... 0/0/0/0（完了/失敗/処理中/母数）")

    def _on_progress(self, payload: dict[str, object]) -> None:
        self.root.after(0, lambda: self._apply_progress(payload))

    def _apply_progress(self, payload: dict[str, object]) -> None:
        completed = int(payload.get("completed", 0))
        failed = int(payload.get("failed", 0))
        processing = int(payload.get("processing", 0))
        total = int(payload.get("total", 0))
        self.status_var.set(f"{completed}/{failed}/{processing}/{total}（完了/失敗/処理中/母数）")

        if payload.get("phase") == "done" and bool(payload.get("cancelled", False)):
            self.was_cancelled = True

    def _run_worker(self, argv: list[str]) -> None:
        try:
            documentChecker.main(argv, progress_callback=self._on_progress, cancel_requested=self.cancel_event.is_set)
        except SystemExit as exc:
            msg = str(exc) if str(exc) else "処理を終了しました。"
            self.root.after(0, lambda: self._finish(False, f"終了: {msg}"))
        except Exception:
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._finish(False, f"エラーが発生しました。\n{tb}"))
        else:
            if self.was_cancelled or self.cancel_event.is_set():
                self.root.after(0, lambda: self._finish(True, "中止しました。途中までの結果を出力している場合があります。"))
            else:
                self.root.after(0, lambda: self._finish(True, "完了しました。出力ファイルを確認してください。"))

    def _finish(self, success: bool, message: str) -> None:
        self.run_button.configure(state=tk.NORMAL)
        self.cancel_button.configure(state=tk.DISABLED)
        self.status_var.set(f"待機中 / {self.status_var.get()}")
        if success:
            messagebox.showinfo("documentChecker", message)
        else:
            messagebox.showerror("documentChecker", message)


def main() -> None:
    root = tk.Tk()
    CheckerGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
