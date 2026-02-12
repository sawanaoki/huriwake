import os
import sys
import json
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess

# 実行環境の判定（EXE化対応）
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
ICON_FILE = os.path.join(BASE_DIR, "app_icon.ico")

PRESET_EXTENSIONS = [
    ".pdf", ".jpg", ".png", ".txt", ".docx", ".xlsx", ".pptx", ".zip", ".csv", ".mp4"
]

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {"mappings": {}}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"mappings": {}}

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")

def get_unique_path(path):
    base, extension = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}_{i}{extension}"):
        i += 1
    return f"{base}_{i}{extension}"

class ConflictDialog:
    def __init__(self, parent, filename):
        self.result = "cancel"
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("ファイル名が重複しています")
        self.dialog.geometry("450x180")
        self.dialog.resizable(False, False)
        self.dialog.attributes("-topmost", True)
        
        label = ttk.Label(self.dialog, text=f"移動先に同名のファイルが存在します:\n{filename}", wraplength=400, justify="center")
        label.pack(pady=20)

        btn_frame = ttk.Frame(self.dialog)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="上書き", command=self.overwrite).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="名前を変更して保存", command=self.rename).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=self.cancel).pack(side=tk.LEFT, padx=5)

        self.dialog.focus_set()
        self.dialog.grab_set()
        self.dialog.wait_window()

    def overwrite(self): self.result = "overwrite"; self.dialog.destroy()
    def rename(self): self.result = "rename"; self.dialog.destroy()
    def cancel(self): self.result = "cancel"; self.dialog.destroy()

class ResultDialog:
    def __init__(self, parent, results):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("振り分け実行結果")
        self.dialog.geometry("600x450")
        
        frame = ttk.Frame(self.dialog, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(frame, text="処理が完了しました", font=("Yu Gothic", 12, "bold"))
        header.pack(pady=(0, 10), anchor="w")

        # Treeviewでリザルトを表示
        cols = ("File", "Status", "Destination")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        tree.heading("File", text="ファイル名")
        tree.heading("Status", text="状態")
        tree.heading("Destination", text="移動先")
        tree.column("File", width=200)
        tree.column("Status", width=80, anchor="center")
        tree.column("Destination", width=250)
        tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)

        for res in results:
            tree.insert("", tk.END, values=(res['name'], res['status'], res['dest']))

        ttk.Button(self.dialog, text="閉じる", command=self.dialog.destroy).pack(pady=10)
        self.dialog.grab_set()
        self.dialog.wait_window()

def distribute_file_internal(file_path, mappings, parent_win, results):
    """単一ファイルの移動ロジック（内部用）"""
    if not os.path.isfile(file_path):
        return

    filename = os.path.basename(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in mappings:
        dest_dir = mappings[ext]
        if not os.path.exists(dest_dir):
            try:
                os.makedirs(dest_dir)
            except Exception as e:
                results.append({"name": filename, "status": "エラー", "dest": f"フォルダ作成失敗: {e}"})
                return

        dest_path = os.path.join(dest_dir, filename)
        
        if os.path.exists(dest_path):
            dialog = ConflictDialog(parent_win, filename)
            choice = dialog.result
            
            if choice == "cancel":
                results.append({"name": filename, "status": "スキップ", "dest": "キャンセルされました"})
                return
            elif choice == "rename":
                dest_path = get_unique_path(dest_path)
        
        try:
            shutil.move(file_path, dest_path)
            results.append({"name": filename, "status": "移動完了", "dest": dest_dir})
        except Exception as e:
            results.append({"name": filename, "status": "エラー", "dest": str(e)})
    else:
        # 設定されていない拡張子は意図的に無視するが、リザルトには載せない（ユーザー要望）
        pass

def process_paths(paths):
    config = load_config()
    mappings = config.get("mappings", {})
    results = []
    
    root = tk.Tk()
    root.withdraw()
    
    # WindowsのDPIスケーリング対応
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    for path in paths:
        if os.path.isdir(path):
            # フォルダ内のファイルをスキャン（直下のみ）
            for item in os.listdir(path):
                full_item_path = os.path.join(path, item)
                if os.path.isfile(full_item_path):
                    distribute_file_internal(full_item_path, mappings, root, results)
        elif os.path.isfile(path):
            distribute_file_internal(path, mappings, root, results)

    if results:
        ResultDialog(root, results)
    else:
        messagebox.showinfo("通知", "移動対象のファイルは見つかりませんでした。\n（設定済みの拡張子を持つファイルのみが移動対象です）")
    
    root.destroy()

class SettingsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Huriwake 設定")
        self.root.geometry("650x500")
        
        if os.path.exists(ICON_FILE):
            try: self.root.iconbitmap(ICON_FILE)
            except Exception: pass
        
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception: pass

        self.config = load_config()
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(frame, text="ファイル振り分け設定", font=("Yu Gothic", 14, "bold"))
        header.grid(row=0, column=0, columnspan=4, pady=(0, 15), sticky="w")

        cols = ("Ext", "Path")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        self.tree.heading("Ext", text="拡張子")
        self.tree.heading("Path", text="移動先フォルダ")
        self.tree.column("Ext", width=100, anchor="center")
        self.tree.column("Path", width=450)
        self.tree.grid(row=1, column=0, columnspan=3, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=1, column=3, sticky="ns")

        input_frame = ttk.LabelFrame(frame, text="設定の追加・更新", padding="10")
        input_frame.grid(row=2, column=0, columnspan=4, pady=15, sticky="ew")

        ttk.Label(input_frame, text="拡張子:").grid(row=0, column=0, padx=5, sticky="w")
        self.ext_combo = ttk.Combobox(input_frame, values=PRESET_EXTENSIONS, width=10)
        self.ext_combo.grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(input_frame, text="フォルダ:").grid(row=0, column=2, padx=5, sticky="w")
        self.path_entry = ttk.Entry(input_frame, width=40)
        self.path_entry.grid(row=0, column=3, padx=5, sticky="ew")

        ttk.Button(input_frame, text="参照", command=self.browse_folder).grid(row=0, column=4, padx=5)
        ttk.Button(input_frame, text="登録", command=self.add_mapping).grid(row=0, column=5, padx=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=4, sticky="ew")

        ttk.Button(btn_frame, text="選択した設定を削除", command=self.delete_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="フォルダを開く", command=self.open_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="デスクトップにショートカットを作成", command=self.create_shortcut).pack(side=tk.RIGHT, padx=5)

        self.tree.bind("<Double-1>", lambda e: self.open_folder())
        self.refresh_list()
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

    def refresh_list(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.config = load_config()
        mappings = self.config.get("mappings", {})
        for ext in sorted(mappings.keys()): self.tree.insert("", tk.END, values=(ext, mappings[ext]))

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path.replace("/", "\\"))

    def add_mapping(self):
        ext = self.ext_combo.get().strip().lower()
        path = self.path_entry.get().strip()
        if ext and not ext.startswith("."): ext = "." + ext
        if ext and path:
            if "mappings" not in self.config: self.config["mappings"] = {}
            self.config["mappings"][ext] = path
            save_config(self.config)
            self.refresh_list()
            self.ext_combo.set("")
            self.path_entry.delete(0, tk.END)
            messagebox.showinfo("完了", f"拡張子 '{ext}' の設定を保存しました。")
        else: messagebox.showwarning("警告", "拡張子とフォルダの両方を入力してください。")

    def delete_mapping(self):
        selected = self.tree.selection()
        if not selected: return
        if messagebox.askyesno("確認", "選択した設定を削除しますか？"):
            for item in selected:
                ext = self.tree.item(item, "values")[0]
                if ext in self.config["mappings"]: del self.config["mappings"][ext]
            save_config(self.config)
            self.refresh_list()

    def open_folder(self):
        selected = self.tree.selection()
        if not selected: return
        path = self.tree.item(selected[0], "values")[1]
        if os.path.exists(path): os.startfile(path)
        else: messagebox.showerror("エラー", f"フォルダが見つかりません:\n{path}")

    def create_shortcut(self):
        desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
        shortcut_path = os.path.join(desktop, "Huriwake.lnk")
        icon_path = ICON_FILE if os.path.exists(ICON_FILE) else sys.executable
        target_path = sys.executable
        arguments = "" if getattr(sys, 'frozen', False) else f"`\"{os.path.abspath(__file__)}`\""
        
        ps_script = f"""
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
        $Shortcut.TargetPath = "{target_path}"
        $Shortcut.Arguments = "{arguments}"
        $Shortcut.WorkingDirectory = "{BASE_DIR}"
        $Shortcut.IconLocation = "{icon_path}"
        $Shortcut.Save()
        """
        try:
            subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True)
            messagebox.showinfo("成功", "デスクトップにショートカットを作成しました。")
        except Exception as e: messagebox.showerror("エラー", f"ショートカットの作成に失敗しました: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        process_paths(sys.argv[1:])
    else:
        root = tk.Tk()
        app = SettingsApp(root)
        root.mainloop()
