import tkinter as tk
from tkinter import messagebox, ttk
import imaplib
import email
from email.header import decode_header
import threading
import json
import os
from email.utils import parseaddr

# IMAPサーバーの設定
IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 993

# グローバル変数
email_data = []
filtered_email_data = []
mail = None
logged_in_user = ""
NUM_EMAILS_TO_LOAD = 100
accounts_file = "accounts.json"
load_thread = None
stop_event = threading.Event()

# タブ用のメールリスト
tab_labels = {
    'ALL': [],
    'Classroom': [],
    'Moodle': [],
    'Canvas': [],
    'CST-VOICE': [],
    'マイナビ': [],
    'その他': []
}

# 各タブのフィルタ条件
tab_filters = {
    'Classroom': 'no-reply@classroom.google.com',
    'Moodle': 'noreply@moodle.ce.cst.nihon-u.ac.jp',
    'Canvas': 'notifications@instructure.com',
    'CST-VOICE': 'cst.voice.mail@nihon-u.ac.jp',
    'マイナビ': ['s-sk-tokyo-career2-cp@mynavi.jp', 'job-s26@mynavi.jp']
}

def save_account(username, password, app_password):
    accounts = load_accounts()
    accounts[username] = {'password': password, 'app_password': app_password}
    with open(accounts_file, 'w') as f:
        json.dump(accounts, f)

def load_accounts():
    if os.path.exists(accounts_file):
        with open(accounts_file, 'r') as f:
            return json.load(f)
    return {}

def connect_to_mailbox(username, password):
    global mail
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(username, password)
        return mail
    except imaplib.IMAP4.error as e:
        messagebox.showerror("Login Error", f"Failed to login: {str(e)}")
        return None

def fetch_emails(mail, start, count):
    try:
        mail.select('inbox')
        _, msg_ids = mail.search(None, 'ALL')
        msg_ids = msg_ids[0].split()
        msg_ids.reverse()  # メールIDを逆順にする
        for num in msg_ids[start:start+count]:
            if stop_event.is_set():
                break
            _, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            # メールの形式や添付ファイルの有無を確認してラベル付けする
            content_type = msg.get_content_type()
            charset = msg.get_content_charset() or 'utf-8'  # 文字コードの指定
            has_attachments = msg.get_content_maintype() == 'multipart' and msg.get_filename()
            has_images = any(part.get_content_type().startswith('image/') for part in msg.walk())

            # 件名のデコード
            subject = decode_header(msg['Subject'])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(charset or 'utf-8', 'ignore')

            # 発信元のデコードと分解
            from_name, from_addr = parseaddr(msg['From'])
            from_addr = decode_header(from_addr)[0][0]
            if isinstance(from_addr, bytes):
                from_addr = from_addr.decode(charset or 'utf-8', 'ignore')

            # タブへのラベル付け
            if any(addr in from_addr for addr in tab_filters['マイナビ']):
                label = 'マイナビ'
            elif tab_filters['Classroom'] in from_addr:
                label = 'Classroom'
            elif tab_filters['Moodle'] in from_addr:
                label = 'Moodle'
            elif tab_filters['Canvas'] in from_addr:
                label = 'Canvas'
            elif tab_filters['CST-VOICE'] in from_addr:
                label = 'CST-VOICE'
            else:
                label = 'その他'

            # メールデータの保存
            email_item = (num, msg, content_type, charset, has_attachments, has_images, subject, from_addr, label)
            tab_labels['ALL'].append(email_item)
            tab_labels[label].append(email_item)
            yield email_item
    except Exception as e:
        messagebox.showerror("Fetch Error", f"Failed to fetch emails: {str(e)}")

def load_emails(start_index=0):
    for email_item in fetch_emails(mail, start_index, NUM_EMAILS_TO_LOAD):
        if stop_event.is_set():
            break

    # 現在のフィルタに従ってメールを再表示する
    on_filter_change()

    # 「さらに表示」ボタンを表示
    if len(tab_labels['ALL']) > 0 and len(tab_labels['ALL']) % NUM_EMAILS_TO_LOAD == 0:
        more_button.pack(pady=10)

def start_loading_emails():
    global load_thread
    stop_loading_emails()  # 既存の読み込みスレッドがある場合は停止する
    for key in tab_labels:
        tab_labels[key].clear()
    email_listbox.delete(0, tk.END)
    stop_event.clear()
    load_thread = threading.Thread(target=load_emails)
    load_thread.start()

def stop_loading_emails():
    global load_thread
    stop_event.set()
    if load_thread:
        load_thread.join()
        load_thread = None

def login(event=None):
    global logged_in_user
    username = username_entry.get()
    password = password_entry.get()

    logged_in_user = username
    show_app_password_screen()

def app_password_login(event=None):
    global mail
    app_password = app_password_entry.get()

    if connect_to_mailbox(logged_in_user, app_password):
        show_mailbox_screen()
        start_loading_emails()
        save_account(logged_in_user, password_entry.get(), app_password)

def show_email_content(event):
    selection = email_listbox.curselection()
    if selection:
        index = selection[0]
        if index < len(filtered_email_data):
            _, msg, content_type, charset, has_attachments, has_images, _, from_addr, _ = filtered_email_data[index]

            # メールの詳細情報を取得
            subject = decode_header(msg['Subject'])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(charset or 'utf-8', 'ignore')

            parsed_date = email.utils.parsedate_to_datetime(msg['Date'])
            formatted_date = parsed_date.strftime("%Y-%m-%d %H:%M:%S")

            # 上部に表示する情報を設定
            info_text = f"From: {from_addr}\n"
            info_text += f"Subject: {subject}\n"
            info_text += f"Received: {formatted_date}\n"
            info_text += "-" * 30 + "\n"

            # メールの内容を取得
            content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain" and not content:
                        content = part.get_payload(decode=True).decode(charset or 'utf-8', 'ignore')
                    elif part.get_content_type() == "text/html" and not content:
                        content = part.get_payload(decode=True).decode(charset or 'utf-8', 'ignore')
            else:
                content = msg.get_payload(decode=True).decode(charset or 'utf-8', 'ignore')

            # ラベル付け情報とメール内容を表示
            content_text.config(state=tk.NORMAL)
            content_text.delete('1.0', tk.END)
            content_text.insert(tk.END, info_text)
            if content_type == 'text/html':
                content_text.set_html(content)
            else:
                content_text.insert(tk.END, content)
            content_text.config(state=tk.DISABLED)

def load_more_emails():
    more_button.pack_forget()
    start_index = len(tab_labels['ALL'])
    threading.Thread(target=load_emails, args=(start_index,)).start()

def on_filter_change(*args):
    selected_filter = filter_var.get()

    # フィルタリングされたメールデータを取得
    filtered_email_data.clear()
    email_listbox.delete(0, tk.END)

    if selected_filter in tab_labels:
        filtered_email_data.extend(tab_labels[selected_filter])
    
    for idx, email_item in enumerate(filtered_email_data):
        email_listbox.insert(tk.END, f"{idx + 1}. {email_item[6]}")  # 件名を表示

    # メールの内容ボックスもリセット
    content_text.config(state=tk.NORMAL)
    content_text.delete('1.0', tk.END)
    content_text.config(state=tk.DISABLED)

def on_scroll(event):
    if event.delta > 0:  # 上スクロールの場合のみ許可
        email_listbox.yview_scroll(int(-1*(event.delta/120)), "units")

def show_login_screen():
    root.geometry("400x300")  # ログイン画面のサイズを指定
    login_frame.pack(fill='both', expand=True)
    app_password_frame.pack_forget()
    mailbox_frame.pack_forget()
    load_saved_accounts()

def show_app_password_screen():
    login_frame.pack_forget()
    root.geometry("400x300")  # アプリパスワード画面のサイズを指定
    app_password_frame.pack(fill='both', expand=True)

def show_mailbox_screen():
    app_password_frame.pack_forget()
    root.attributes("-fullscreen", False)  # フルスクリーンを解除
    root.state('zoomed')  # ウィンドウ状態のフルスクリーンにする
    mailbox_frame.pack(fill='both', expand=True)
    user_label.config(text=f"Logged in as: {logged_in_user}")

def load_saved_accounts():
    accounts = load_accounts()
    for username in accounts:
        saved_accounts_listbox.insert(tk.END, username)

def select_account(event):
    selection = saved_accounts_listbox.curselection()
    if selection:
        index = selection[0]
        username = saved_accounts_listbox.get(index)
        accounts = load_accounts()
        if username in accounts:
            credentials = accounts[username]
            username_entry.delete(0, tk.END)
            username_entry.insert(0, username)
            password_entry.delete(0, tk.END)
            password_entry.insert(0, credentials['password'])
            app_password_entry.delete(0, tk.END)
            app_password_entry.insert(0, credentials['app_password'])
            login()

def on_closing():
    stop_event.set()
    if load_thread:
        load_thread.join()
    if mail:
        mail.logout()
    root.destroy()

# GUIの作成
root = tk.Tk()
root.title("Email Client")

# ログイン画面
login_frame = tk.Frame(root)
username_label = tk.Label(login_frame, text="Email Address:")
username_label.pack(pady=5)
username_entry = tk.Entry(login_frame, width=30)
username_entry.pack()

password_label = tk.Label(login_frame, text="Password:")
password_label.pack(pady=5)
password_entry = tk.Entry(login_frame, width=30, show="*")
password_entry.pack()

# Enterキーでログインを実行
username_entry.bind('<Return>', login)
password_entry.bind('<Return>', login)

login_button = tk.Button(login_frame, text="Login", command=login)
login_button.pack(pady=10)

# 保存されたアカウントを表示するリストボックス
saved_accounts_listbox = tk.Listbox(login_frame)
saved_accounts_listbox.pack(fill=tk.BOTH, expand=True, pady=10)
saved_accounts_listbox.bind('<<ListboxSelect>>', select_account)

# アプリパスワード画面
app_password_frame = tk.Frame(root)
app_password_label = tk.Label(app_password_frame, text="App Password:")
app_password_label.pack(pady=5)
app_password_entry = tk.Entry(app_password_frame, width=30, show="*")
app_password_entry.pack()

# Enterキーでアプリパスワードログインを実行
app_password_entry.bind('<Return>', app_password_login)

app_password_button = tk.Button(app_password_frame, text="Login", command=app_password_login)
app_password_button.pack(pady=10)

# メールボックス画面
mailbox_frame = tk.Frame(root)
user_label = tk.Label(mailbox_frame, text="Logged in as:")
user_label.pack(anchor='ne', pady=5, padx=5)

# フィルタ用のプルダウンメニュー
filter_var = tk.StringVar(value='ALL')
filter_menu = ttk.Combobox(mailbox_frame, textvariable=filter_var, values=list(tab_labels.keys()), state='readonly')
filter_menu.pack(pady=5, anchor='nw', padx=5)
filter_var.trace_add('write', on_filter_change)

# メールリストと内容表示を左右に分ける
left_frame = tk.Frame(mailbox_frame)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
right_frame = tk.Frame(mailbox_frame)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# メールリスト
email_listbox = tk.Listbox(left_frame, width=50, height=20)
email_listbox.pack(fill=tk.BOTH, expand=True)
email_listbox.bind('<<ListboxSelect>>', show_email_content)

# スクロールイベントのバインディングを追加
email_listbox.bind('<MouseWheel>', on_scroll)

# メール内容表示
content_text = HTMLScrolledText(right_frame, highlightthickness=0)
content_text.pack(fill=tk.BOTH, expand=True)

more_button = tk.Button(left_frame, text="さらに表示", command=load_more_emails)

# 初期状態でALLを選択
on_filter_change()

show_login_screen()

# アプリケーション終了時にプログラムタスクを終了するためのハンドラー
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
