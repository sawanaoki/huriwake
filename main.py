import os
import sys
import json
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import fnmatch

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

RULE_TYPES = {
    "拡張子": "extension",
    "ファイル名（部分一致）": "filename",
    "フォルダ名（部分一致）": "foldername",
}
RULE_TYPE_LABELS = list(RULE_TYPES.keys())
RULE_TYPE_VALUES = list(RULE_TYPES.values())

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {"rules": []}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 旧形式（mappings）との後方互換
        if "mappings" in data and "rules" not in data:
            rules = []
            for ext, dest in data["mappings"].items():
                rules.append({"type": "extension", "pattern": ext, "dest": dest})
            data = {"rules": rules}
            save_config(data)
        if "rules" not in data:
            data["rules"] = []
        return data
    except Exception:
        return {"rules": []}

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

def type_value_to_label(value):
    for label, val in RULE_TYPES.items():
        if val == value:
            return label
    return value

def match_rule(rule, file_path):
    """ファイルパスがルールに一致するか判定"""
    rule_type = rule.get("type", "extension")
    pattern = rule.get("pattern", "").lower()
    filename = os.path.basename(file_path).lower()
    parent_folder = os.path.basename(os.path.dirname(file_path)).lower()

    if rule_type == "extension":
        ext = os.path.splitext(file_path)[1].lower()
        return ext == pattern
    elif rule_type == "filename":
        # 部分一致: パターンがファイル名に含まれていればマッチ
        return pattern in filename
    elif rule_type == "foldername":
        # 部分一致: パターンが親フォルダ名に含まれていればマッチ
        return pattern in parent_folder
    return False

def find_matching_rule(rules, file_path):
    """ファイルに最初にマッチするルールを返す（優先度: リスト順）"""
    for rule in rules:
        if match_rule(rule, file_path):
            return rule
    return None

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
        self.dialog.geometry("700x450")

        frame = ttk.Frame(self.dialog, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(frame, text="処理が完了しました", font=("Yu Gothic", 12, "bold"))
        header.pack(pady=(0, 10), anchor="w")

        cols = ("File", "Rule", "Status", "Destination")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        tree.heading("File", text="ファイル名")
        tree.heading("Rule", text="マッチしたルール")
        tree.heading("Status", text="状態")
        tree.heading("Destination", text="移動先")
        tree.column("File", width=180)
        tree.column("Rule", width=150)
        tree.column("Status", width=80, anchor="center")
        tree.column("Destination", width=220)
        tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)

        for res in results:
            tree.insert("", tk.END, values=(res['name'], res.get('rule', ''), res['status'], res['dest']))

        ttk.Button(self.dialog, text="閉じる", command=self.dialog.destroy).pack(pady=10)
        self.dialog.grab_set()
        self.dialog.wait_window()

def distribute_file_internal(file_path, rules, parent_win, results):
    """単一ファイルの移動ロジック（内部用）"""
    if not os.path.isfile(file_path):
        return

    filename = os.path.basename(file_path)
    rule = find_matching_rule(rules, file_path)

    if rule is None:
        # マッチするルールなし → 無視
        return

    rule_label = f"{type_value_to_label(rule['type'])}: {rule['pattern']}"
    dest_dir = rule["dest"]

    if not os.path.exists(dest_dir):
        try:
            os.makedirs(dest_dir)
        except Exception as e:
            results.append({"name": filename, "rule": rule_label, "status": "エラー", "dest": f"フォルダ作成失敗: {e}"})
            return

    dest_path = os.path.join(dest_dir, filename)

    if os.path.exists(dest_path):
        dialog = ConflictDialog(parent_win, filename)
        choice = dialog.result

        if choice == "cancel":
            results.append({"name": filename, "rule": rule_label, "status": "スキップ", "dest": "キャンセルされました"})
            return
        elif choice == "rename":
            dest_path = get_unique_path(dest_path)

    try:
        shutil.move(file_path, dest_path)
        results.append({"name": filename, "rule": rule_label, "status": "移動完了", "dest": dest_dir})
    except Exception as e:
        results.append({"name": filename, "rule": rule_label, "status": "エラー", "dest": str(e)})

def process_paths(paths):
    config = load_config()
    rules = config.get("rules", [])
    results = []

    root = tk.Tk()
    root.withdraw()

    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    for path in paths:
        if os.path.isdir(path):
            for item in os.listdir(path):
                full_item_path = os.path.join(path, item)
                if os.path.isfile(full_item_path):
                    distribute_file_internal(full_item_path, rules, root, results)
        elif os.path.isfile(path):
            distribute_file_internal(path, rules, root, results)

    if results:
        ResultDialog(root, results)
    else:
        messagebox.showinfo("通知", "移動対象のファイルは見つかりませんでした。\n（設定済みのルールに一致するファイルのみが移動対象です）")

    root.destroy()

class SettingsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Huriwake 設定")
        self.root.geometry("750x550")

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
        header.grid(row=0, column=0, columnspan=4, pady=(0, 10), sticky="w")

        desc = ttk.Label(frame, text="ルールは上から順に評価されます。最初にマッチしたルールが適用されます。", foreground="gray")
        desc.grid(row=1, column=0, columnspan=4, pady=(0, 10), sticky="w")

        # ルールリスト
        cols = ("Type", "Pattern", "Path")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        self.tree.heading("Type", text="種別")
        self.tree.heading("Pattern", text="パターン")
        self.tree.heading("Path", text="移動先フォルダ")
        self.tree.column("Type", width=150, anchor="center")
        self.tree.column("Pattern", width=150, anchor="center")
        self.tree.column("Path", width=350)
        self.tree.grid(row=2, column=0, columnspan=3, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=2, column=3, sticky="ns")

        # 入力フォーム
        input_frame = ttk.LabelFrame(frame, text="ルールの追加", padding="10")
        input_frame.grid(row=3, column=0, columnspan=4, pady=10, sticky="ew")

        ttk.Label(input_frame, text="種別:").grid(row=0, column=0, padx=5, sticky="w")
        self.type_combo = ttk.Combobox(input_frame, values=RULE_TYPE_LABELS, width=18, state="readonly")
        self.type_combo.grid(row=0, column=1, padx=5, sticky="w")
        self.type_combo.current(0)
        self.type_combo.bind("<<ComboboxSelected>>", self.on_type_changed)

        ttk.Label(input_frame, text="パターン:").grid(row=0, column=2, padx=5, sticky="w")
        self.pattern_combo = ttk.Combobox(input_frame, values=PRESET_EXTENSIONS, width=15)
        self.pattern_combo.grid(row=0, column=3, padx=5, sticky="w")

        ttk.Label(input_frame, text="移動先:").grid(row=1, column=0, padx=5, pady=(5, 0), sticky="w")
        self.path_entry = ttk.Entry(input_frame, width=50)
        self.path_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=(5, 0), sticky="ew")
        ttk.Button(input_frame, text="参照", command=self.browse_folder).grid(row=1, column=4, padx=5, pady=(5, 0))
        ttk.Button(input_frame, text="登録", command=self.add_rule).grid(row=1, column=5, padx=5, pady=(5, 0))

        input_frame.columnconfigure(3, weight=1)

        # ヒントラベル
        self.hint_label = ttk.Label(frame, text="💡 拡張子: 例「.pdf」 ／ ファイル名: 例「議事録」 ／ フォルダ名: 例「Downloads」", foreground="gray")
        self.hint_label.grid(row=4, column=0, columnspan=4, pady=(0, 5), sticky="w")

        # 下部ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=4, sticky="ew")

        ttk.Button(btn_frame, text="選択したルールを削除", command=self.delete_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="フォルダを開く", command=self.open_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="▲ 上へ", command=self.move_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="▼ 下へ", command=self.move_down).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="デスクトップにショートカットを作成", command=self.create_shortcut).pack(side=tk.RIGHT, padx=5)

        self.tree.bind("<Double-1>", lambda e: self.open_folder())
        self.refresh_list()
        frame.rowconfigure(2, weight=1)
        frame.columnconfigure(1, weight=1)

    def on_type_changed(self, event=None):
        selected = self.type_combo.get()
        if selected == "拡張子":
            self.pattern_combo.configure(values=PRESET_EXTENSIONS, state="normal")
        else:
            self.pattern_combo.configure(values=[], state="normal")
            self.pattern_combo.set("")

    def refresh_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.config = load_config()
        rules = self.config.get("rules", [])
        for rule in rules:
            type_label = type_value_to_label(rule.get("type", "extension"))
            self.tree.insert("", tk.END, values=(type_label, rule.get("pattern", ""), rule.get("dest", "")))

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path.replace("/", "\\"))

    def add_rule(self):
        type_label = self.type_combo.get()
        pattern = self.pattern_combo.get().strip().lower()
        dest = self.path_entry.get().strip()

        if not type_label or not pattern or not dest:
            messagebox.showwarning("警告", "すべての項目を入力してください。")
            return

        rule_type = RULE_TYPES.get(type_label, "extension")

        if rule_type == "extension" and not pattern.startswith("."):
            pattern = "." + pattern

        new_rule = {"type": rule_type, "pattern": pattern, "dest": dest}

        if "rules" not in self.config:
            self.config["rules"] = []
        self.config["rules"].append(new_rule)
        save_config(self.config)
        self.refresh_list()
        self.pattern_combo.set("")
        self.path_entry.delete(0, tk.END)
        messagebox.showinfo("完了", f"ルールを登録しました: {type_label} = '{pattern}'")

    def delete_rule(self):
        selected = self.tree.selection()
        if not selected:
            return
        if messagebox.askyesno("確認", "選択したルールを削除しますか？"):
            indices = sorted([self.tree.index(item) for item in selected], reverse=True)
            for idx in indices:
                if idx < len(self.config["rules"]):
                    del self.config["rules"][idx]
            save_config(self.config)
            self.refresh_list()

    def move_up(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = self.tree.index(selected[0])
        if idx > 0:
            self.config["rules"][idx], self.config["rules"][idx - 1] = self.config["rules"][idx - 1], self.config["rules"][idx]
            save_config(self.config)
            self.refresh_list()
            children = self.tree.get_children()
            if idx - 1 < len(children):
                self.tree.selection_set(children[idx - 1])

    def move_down(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = self.tree.index(selected[0])
        if idx < len(self.config["rules"]) - 1:
            self.config["rules"][idx], self.config["rules"][idx + 1] = self.config["rules"][idx + 1], self.config["rules"][idx]
            save_config(self.config)
            self.refresh_list()
            children = self.tree.get_children()
            if idx + 1 < len(children):
                self.tree.selection_set(children[idx + 1])

    def open_folder(self):
        selected = self.tree.selection()
        if not selected:
            return
        path = self.tree.item(selected[0], "values")[2]
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{path}")

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
        except Exception as e:
            messagebox.showerror("エラー", f"ショートカットの作成に失敗しました: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        process_paths(sys.argv[1:])
    else:
        root = tk.Tk()
        app = SettingsApp(root)
        root.mainloop()
