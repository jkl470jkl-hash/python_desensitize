import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

from core_redact import process_docx, process_txt, run_self_test, OUTPUT_DIR

SUPPORTED_EXT = {".txt", ".docx"}


class FileItem:
    def __init__(self, path: str):
        self.path = path
        self.name = os.path.basename(path)
        self.status = "待处理"
        self.stats = {}
        self.message = ""


class RedactGUI:
    def __init__(self, root: TkinterDnD.Tk):
        self.root = root
        self.root.title("中文合同脱敏工具（本地离线）")
        self.root.geometry("900x600")
        self.files: list[FileItem] = []
        self.output_dir = os.path.join(os.path.dirname(__file__), "output")

        self._build_widgets()

    def _build_widgets(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top_frame, text="输出目录:").pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=self.output_dir)
        self.output_entry = ttk.Entry(top_frame, textvariable=self.output_var, width=60)
        self.output_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="选择", command=self.choose_output_dir).pack(side=tk.LEFT)
        ttk.Button(top_frame, text="打开", command=self.open_output_dir).pack(side=tk.LEFT, padx=5)

        mid_frame = ttk.Frame(self.root)
        mid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(mid_frame, columns=("path", "status"), show="headings")
        self.tree.heading("path", text="文件路径")
        self.tree.heading("status", text="状态")
        self.tree.column("path", width=550)
        self.tree.column("status", width=100)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(mid_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        drop_label = ttk.Label(self.root, text="将 .txt / .docx 文件拖拽到列表，或使用下方按钮添加", anchor="center")
        drop_label.pack(fill=tk.X, padx=10)
        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind("<Drop>", self.on_drop)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_frame, text="添加文件", command=self.add_files_dialog).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_files).pack(side=tk.LEFT, padx=5)
        self.start_btn = ttk.Button(btn_frame, text="开始脱敏", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="自测", command=self.show_self_test).pack(side=tk.LEFT, padx=5)

        log_frame = ttk.LabelFrame(self.root, text="日志")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def choose_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)
            self.output_dir = path

    def open_output_dir(self):
        path = self.output_var.get()
        if not path:
            return
        if not os.path.exists(path):
            os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)
        except AttributeError:
            messagebox.showinfo("提示", f"请手动前往目录: {path}")

    def add_files_dialog(self):
        paths = filedialog.askopenfilenames(filetypes=[("合同文件", "*.txt *.docx")])
        if paths:
            self.add_files(list(paths))

    def on_drop(self, event):
        paths = self.root.splitlist(event.data)
        self.add_files(list(paths))

    def add_files(self, paths: list[str]):
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".doc":
                self.log(f"跳过 {os.path.basename(path)}，请先另存为 .docx 再拖入。")
                continue
            if ext not in SUPPORTED_EXT:
                self.log(f"不支持的文件格式：{path}")
                continue
            item = FileItem(path)
            self.files.append(item)
            self.tree.insert("", tk.END, iid=id(item), values=(self._shorten_path(path), item.status))

    def _shorten_path(self, path: str, max_len: int = 70):
        return path if len(path) <= max_len else f"...{path[-max_len:]}"

    def clear_files(self):
        self.files.clear()
        for child in self.tree.get_children():
            self.tree.delete(child)

    def show_self_test(self):
        run_self_test()
        messagebox.showinfo("自测", "控制台已输出自测示例，请在终端查看。")

    def start_processing(self):
        if not self.files:
            messagebox.showinfo("提示", "请先添加要脱敏的文件")
            return
        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._process_all, daemon=True).start()

    def _process_all(self):
        for item in self.files:
            self.update_status(item, "处理中")
            try:
                ext = os.path.splitext(item.path)[1].lower()
                if ext == ".docx":
                    folder, result = process_docx(item.path, output_root=self.output_dir)
                elif ext == ".txt":
                    folder, result = process_txt(item.path, output_root=self.output_dir)
                else:
                    raise ValueError("不支持的文件类型")
                item.stats = result.stats
                self.update_status(item, "已完成")
                self.log(f"完成：{item.name} → {folder}")
                if result.stats:
                    stats_str = ", ".join([f"{k}:{v}" for k, v in result.stats.items()])
                    self.log(f"发现敏感项：{stats_str}")
            except Exception as exc:  # noqa: BLE001
                item.message = str(exc)
                self.update_status(item, "出错")
                self.log(f"处理失败：{item.name}，原因：{exc}")
        self.start_btn.config(state=tk.NORMAL)

    def update_status(self, item: FileItem, status: str):
        item.status = status
        self.tree.item(id(item), values=(self._shorten_path(item.path), status))

    def log(self, text: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


def main():
    root = TkinterDnD.Tk()
    app = RedactGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
