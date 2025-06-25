import os
import pathlib
import win32com.client
import configparser
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import sys

def get_base_path():
    if getattr(sys, 'frozen', False):
        return pathlib.Path(sys._MEIPASS)
    else:
        return pathlib.Path(os.path.dirname(os.path.abspath(__file__)))

base_path = get_base_path()
config_path = base_path / "config" / "config.ini"
certs_path = base_path / "config" / "certs"

def import_certificate_with_certutil(cert_path: str, password: str):
    """certutil.exe を使用して証明書をインポートします."""
    try:
        command = [
            "Certutil",
            "-p",
            password,
            "-user",
            "-importpfx",
            cert_path
        ]
        result = subprocess.run(command, capture_output=True, text=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW) # コンソールウィンドウを非表示にする
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        return False, f"証明書のインポートに失敗しました (certutil)。エラーコード: {e.returncode}\n{e.stderr}"
    except FileNotFoundError:
        return False, "Error: certutil.exe が見つかりません。"
    except Exception as e:
        return False, f"予期しないエラーが発生しました: {e}"

def install_certificates(selected_sections):
    """選択された複数の証明書をインストールします."""
    global certs_path, config

    for section in selected_sections:
        cert_num = config[section].get('cert_num')
        password = config[section].get('password')

        if cert_num and password:
            cert_file = f"client-{cert_num}.p12"
            cert_full_path = str(certs_path / cert_file)

            install_success, install_message = import_certificate_with_certutil(cert_full_path, password)
            if install_success:
                 messagebox.showinfo("成功", f"セクション '{section}' の証明書が正常にインストールされました。")
            else:
                 messagebox.showerror("インストール失敗", f"セクション '{section}' のインストールに失敗しました。\n\n詳細:\n{install_message}")

        else:
            messagebox.showerror("エラー", f"セクション '{section}' に cert_num または password が見つかりません。")

def get_selected_certificates():
    """リストボックスで選択された証明書のセクション名を返します."""
    selected_indices = cert_list.curselection()
    return [available_certs_displayed[i] for i in selected_indices]

def install_selected():
    selected_certs = get_selected_certificates()
    if not selected_certs:
        messagebox.showerror("エラー", "インストールする証明書を選択してください。")
        return
    
    # 確認ダイアログの表示
    confirm_message = "以下の証明書をインストールしますか？\n\n- " + "\n- ".join(selected_certs)
    if messagebox.askyesno("インストール確認", confirm_message):
        install_certificates(selected_certs)

def select_all():
    """リストボックスのすべての項目を選択します."""
    cert_list.select_set(0, tk.END)

def deselect_all():
    """リストボックスのすべての選択を解除します."""
    cert_list.select_clear(0, tk.END)

def update_cert_list(filter_text=None):
    cert_list.delete(0, tk.END)
    available_certs_displayed.clear()
    show_hidden = show_hidden_var.get()
    for section in config.sections():
        if config.has_option(section, 'cert_num'):
            cert_num = config[section].get('cert_num', '不明')
            hidden = config[section].getint('hidden', 0)
            display = False
            if show_hidden or hidden == 0:
                if filter_text is None or filter_text.lower() in section.lower():
                    display = True
            if display:
                cert_list.insert(tk.END, f"{section}")
                available_certs_displayed.append(section)

def filter_by_db():
    selected_db = db_filter_var.get()
    update_cert_list(selected_db if selected_db != "ALL" else None)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("証明書インストーラー")
    root.geometry("650x600") # ウィンドウサイズを少し調整

    if not config_path.exists():
        messagebox.showerror("エラー", f"設定ファイル '{config_path}' が見つかりません。")
        root.destroy()
        sys.exit(1)

    config = configparser.ConfigParser()
    try:
        config.read(config_path, encoding='utf-8')
    except UnicodeDecodeError:
        config.read(config_path, encoding='cp932')
    except Exception as e:
        messagebox.showerror("エラー", f"設定ファイル '{config_path}' の読み込み中にエラーが発生しました: {e}")
        root.destroy()
        sys.exit(1)

    available_certs_displayed = []
    db_options = ["ALL"]
    for section in config.sections():
        if config.has_option(section, 'cert_num'):
            for db_type in ["DB1", "DB2", "DB3", "DB4"]:
                if db_type.lower() in section.lower() and db_type not in db_options:
                    db_options.append(db_type)
    
    # "ALL"以外をソート
    sorted_dbs = sorted([opt for opt in db_options if opt != "ALL"])
    db_options = ["ALL"] + sorted_dbs


    show_hidden_var = tk.BooleanVar()
    show_hidden_var.set(False)

    filter_label = ttk.Label(root, text="絞り込み設定:")
    filter_label.pack(pady=5, anchor='w', padx=10)

    filter_frame = ttk.Frame(root)
    filter_frame.pack(pady=5, fill='x', padx=10)

    db_filter_var = tk.StringVar(value="ALL")
    for db_type in db_options:
        radio_button = ttk.Radiobutton(filter_frame, text=db_type, variable=db_filter_var, value=db_type, command=filter_by_db)
        radio_button.pack(side=tk.LEFT, padx=5)

    # --- 変更点 1: commandをfilter_by_dbに変更 ---
    show_hidden_check = ttk.Checkbutton(root, text="現在稼働していない事業所を表示", variable=show_hidden_var, command=filter_by_db)
    show_hidden_check.pack(pady=5, anchor='w', padx=10)

    cert_label = ttk.Label(root, text="インストールする証明書を選択してください:")
    cert_label.pack(pady=(10, 0), anchor='w', padx=10)

    cert_list_frame = ttk.Frame(root)
    cert_list_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

    cert_list_scrollbar = ttk.Scrollbar(cert_list_frame, orient=tk.VERTICAL)
    cert_list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # --- 変更点 2: fontオプションを追加 ---
    cert_list = tk.Listbox(cert_list_frame, height=15, width=80, yscrollcommand=cert_list_scrollbar.set, selectmode=tk.MULTIPLE, font=("Meiryo UI", 12))
    cert_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    cert_list_scrollbar.config(command=cert_list.yview)

    bottom_frame = ttk.Frame(root)
    bottom_frame.pack(side=tk.BOTTOM, padx=10, pady=10, fill=tk.X)

    button_frame = ttk.Frame(bottom_frame)
    button_frame.pack(expand=True, fill='x')

    select_all_button = ttk.Button(button_frame, text="全選択", command=select_all)
    select_all_button.pack(side=tk.LEFT)

    deselect_all_button = ttk.Button(button_frame, text="全選択解除", command=deselect_all)
    deselect_all_button.pack(side=tk.LEFT, padx=5)

    install_button = ttk.Button(button_frame, text="インストール", command=install_selected, padding=(10, 5))
    install_button.pack(side=tk.RIGHT)

    update_cert_list()

    root.mainloop()