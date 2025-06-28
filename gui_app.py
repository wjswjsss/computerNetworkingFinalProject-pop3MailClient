#!/usr/bin/env python3
import threading
import poplib
import ssl
import tempfile
import webbrowser
import smtplib
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
import difflib
import markdown
import os
import pathlib
import re
import html as html_lib

# -------------------
# Default QQ credentials (modifiable via login UI) for frequent testing
# -------------------
DEFAULT_USER = '123456@qq.com'      # ur account
DEFAULT_PWD  = 'I dont know'        # ur pwd, well uh I forgot to remove the pwd & account when initial git push, 
                                    # but dont try to use them cuz I have disabled them xd. 

# Environment overrides (if set)
USER_ENV     = os.getenv('POP3_USER')
PWD_ENV      = os.getenv('POP3_PWD')
SMTP_PWD_ENV = os.getenv('SMTP_PWD')

# Effective credentials initialized to defaults, may be overridden at login
USER       = USER_ENV or DEFAULT_USER
PWD        = PWD_ENV  or DEFAULT_PWD
SMTP_PWD   = SMTP_PWD_ENV

# Server settings
POP3_SERVER = os.getenv('POP3_HOST', 'pop.qq.com')
POP3_PORT   = int(os.getenv('POP3_PORT', '995'))
POP3_SSL    = True
SMTP_SERVER = os.getenv('SMTP_HOST', 'smtp.qq.com')
SMTP_PORT   = int(os.getenv('SMTP_PORT', '587'))
SMTP_USE_TLS= True
SMTP_TIMEOUT= 15

# Poll interval (ms)
POLL_INTERVAL = 5000  # check every 5 seconds

# Logging
logging.basicConfig(level=logging.INFO)

class POP3GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QQ POP3 Mail Client")
        self.geometry("1024x640")
        self.client = None
        self.client_lock = threading.Lock()
        self.headers_info = {}
        self.current_preview = None
        self.load_batch_size = 20
        self.next_load_start = None
        self._build_ui()
        self._connect_pop3()
        # initial load and then start polling
        threading.Thread(target=self._init_load, daemon=True).start()
        self.after(POLL_INTERVAL, self._poll_new_emails)

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('aqua')
        style.configure('Treeview', rowheight=24, font=('Arial', 10))
        style.configure('Treeview.Heading', font=('Arial', 11, 'bold'))
        style.configure('TButton', padding=6)
        style.configure('TEntry', padding=4)

        # Primary actions (top row)
        toolbar1 = ttk.Frame(self, padding=4)
        toolbar1.pack(fill='x')
        ttk.Button(toolbar1, text='Compose',    command=self._open_compose).pack(side='left', padx=4)
        ttk.Button(toolbar1, text='Load More',  command=self._load_more).pack(side='left', padx=4)
        ttk.Button(toolbar1, text='Load All',   command=self._load_all).pack(side='left', padx=4)
        self.preview_btn = ttk.Button(toolbar1, text='Preview', state='disabled', command=self._show_preview)
        self.preview_btn.pack(side='left', padx=4)
        self.open_btn   = ttk.Button(toolbar1, text='Open in Browser', state='disabled', command=self._open_in_browser)
        self.open_btn.pack(side='left', padx=4)

        # Search controls (second row)
        toolbar2 = ttk.Frame(self, padding=4)
        toolbar2.pack(fill='x')
        ttk.Label(toolbar2, text='Search:').pack(side='left', padx=(4,0))
        self.search_var = tk.StringVar()
        ttk.Entry(toolbar2, textvariable=self.search_var, width=30).pack(side='left', padx=4)
        ttk.Button(toolbar2, text='Go',       command=self._on_search).pack(side='left', padx=4)
        ttk.Button(toolbar2, text='Show All', command=self._show_all).pack(side='left', padx=4)
        ttk.Label(toolbar2, text='Date:').pack(side='left', padx=(20,0))
        self.date_var = tk.StringVar()
        ttk.Entry(toolbar2, textvariable=self.date_var, width=12).pack(side='left', padx=4)
        ttk.Button(toolbar2, text='Search Date', command=self._on_date_search).pack(side='left', padx=4)

        # Message list & scrollbars
        content = ttk.Frame(self)
        content.pack(fill='both', expand=True, side='left')
        cols = ('Num', 'From', 'Subject', 'Date')
        self.tree = ttk.Treeview(content, columns=cols, show='headings', selectmode='browse')
        for col, width in zip(cols, (60,220,420,150)):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor='w', stretch=(col=='Date'))
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        vsb = ttk.Scrollbar(content, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(content, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

    def _connect_pop3(self):
        try:
            if POP3_SSL:
                self.client = poplib.POP3_SSL(POP3_SERVER, POP3_PORT, timeout=10)
            else:
                self.client = poplib.POP3(POP3_SERVER, POP3_PORT, timeout=10)
                self.client.stls(context=ssl.create_default_context())
            self.client.user(USER)
            self.client.pass_(PWD)
        except Exception:
            logging.exception("POP3 connect failed")
            messagebox.showerror("Connection Failed", f"Could not connect to {POP3_SERVER}:{POP3_PORT}")
            self.destroy()

    def _init_load(self):
        # initial batch
        try:
            with self.client_lock:
                count, _ = self.client.stat()
            start = max(1, count - self.load_batch_size + 1)
            self._load_range(start, count)
            self.next_load_start = start - 1
        except Exception:
            logging.exception("Initial load error")

    def _load_range(self, start, end):
        for i in range(end, start-1, -1):
            if i in self.headers_info:
                continue
            try:
                with self.client_lock:
                    resp, lines, _ = self.client.top(i, 0)
                msg = BytesParser(policy=policy.default).parsebytes(b"\r\n".join(lines))
                self.headers_info[i] = msg
                self.tree.insert('', 'end', iid=str(i), values=(i, msg.get('from',''), msg.get('subject',''), msg.get('date','')))
            except Exception:
                logging.warning(f"Failed header {i}")

    def _load_more(self):
        if not self.next_load_start or self.next_load_start < 1:
            return
        end = self.next_load_start
        start = max(1, end - self.load_batch_size + 1)
        self._load_range(start, end)
        self.next_load_start = start - 1

    def _load_all(self):
        try:
            with self.client_lock:
                count, _ = self.client.stat()
                print(f"[debug] Total emails: {count}")
            self.tree.delete(*self.tree.get_children())
            self.headers_info.clear()
            self._load_range(1, count)
            self.next_load_start = 0
        except Exception:
            logging.exception("Load all failed")

    def _poll_new_emails(self):
        """Periodically check for new messages and display notification."""
        try:
            with self.client_lock:
                self.client.quit()
                self._connect_pop3()
            # highest existing uid
            count, _ = self.client.stat()
            existing = [int(i) for i in self.headers_info.keys()]
            max_existing = max(existing) if existing else 0
            if count > max_existing:
                # new messages arrived
                self._load_range(max_existing+1, count)
                self.next_load_start = max_existing
                messagebox.showinfo("New Email", f"{count - max_existing} new message(s) received.")
        except Exception:
            logging.exception("Poll error")
        finally:
            # reschedule
            self.after(POLL_INTERVAL, self._poll_new_emails)

    def _on_search(self):
        self.tree.delete(*self.tree.get_children())
        kws = [kw for kw in self.search_var.get().lower().split() if kw]
        matches = []
        for num, msg in self.headers_info.items():
            text = (msg.get('subject','') + ' ' + msg.get('from','')).lower()
            if all(kw in text for kw in kws):
                score = sum(difflib.SequenceMatcher(None,kw,text).ratio() for kw in kws)
                matches.append((score, num))
        for _, num in sorted(matches, reverse=True):
            m = self.headers_info[num]
            self.tree.insert('', 'end', iid=str(num), values=(num, m.get('from',''), m.get('subject',''), m.get('date','')))

    def _on_date_search(self):
        self.tree.delete(*self.tree.get_children())
        # normalize whitespace and case
        date_query = ' '.join(self.date_var.get().split()).lower()
        for num, msg in self.headers_info.items():
            date_str = msg.get('date','')
            norm_date = ' '.join(date_str.split()).lower()
            if date_query in norm_date:
                self.tree.insert('', 'end', iid=str(num),
                                 values=(num,
                                         msg.get('from',''),
                                         msg.get('subject',''),
                                         date_str))

    def _show_all(self):
        self.tree.delete(*self.tree.get_children())
        for num in sorted(self.headers_info.keys(), reverse=True):
            m = self.headers_info[num]
            self.tree.insert('', 'end', iid=str(num), values=(num, m.get('from',''), m.get('subject',''), m.get('date','')))

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        try:
            with self.client_lock:
                resp, lines, _ = self.client.retr(idx)
            msg = BytesParser(policy=policy.default).parsebytes(b"\r\n".join(lines))
            part = msg.get_body(preferencelist=('html','plain'))
            self.current_preview = part.get_content() if part else ''
            self.preview_btn.state(['!disabled'])
            self.open_btn.state(['!disabled'])
        except Exception:
            logging.warning(f"Failed retrieve {idx}")

    def _show_preview(self):
        raw = self.current_preview or ''
        text = re.sub(r'(?i)<br\s*/?>', '\n', raw)
        text = re.sub(r'(?i)</div>', '\n', text)
        text = re.sub(r'(?i)<div[^>]*>', '', text)
        text = re.sub(r'(?i)<blockquote[^>]*>', '\n> ', text)
        text = re.sub(r'(?i)</blockquote>', '\n', text)
        text = re.sub(r'(?i)</p>', '\n\n', text)
        text = re.sub(r'<[^>]+>', '', text)
        text = html_lib.unescape(text).strip()
        win = tk.Toplevel(self)
        win.title('Local Preview')
        win.geometry('600x400')
        txt = tk.Text(win, wrap='word')
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', text)
        txt.config(state='disabled')

    def _open_in_browser(self):
        html = f"<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>{self.current_preview or ''}</body></html>"
        with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html)
            tmp = f.name
        webbrowser.open(pathlib.Path(tmp).as_uri(), new=2)
        self.after(POLL_INTERVAL, lambda p=tmp: os.remove(p) if os.path.exists(p) else None)

    def _open_compose(self):
        win = tk.Toplevel(self)
        win.title('Compose')
        win.geometry('600x400')
        win.columnconfigure(1, weight=1)
        win.rowconfigure(2, weight=1)
        ttk.Label(win, text='To:').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        to_var = tk.StringVar()
        ttk.Entry(win, textvariable=to_var).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(win, text='Subject:').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        subj_var = tk.StringVar()
        ttk.Entry(win, textvariable=subj_var).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(win, text='Body:').grid(row=2, column=0, padx=5, pady=5, sticky='ne')
        body = tk.Text(win, wrap='word')
        body.grid(row=2, column=1, sticky='nsew', padx=5, pady=5)
        send_btn = ttk.Button(win, text='Send', command=lambda:self._send_email(to_var.get(), subj_var.get(), body.get('1.0','end')))
        send_btn.grid(row=3, column=1, sticky='e', padx=5, pady=5)

    def _send_email(self, to, subj, body_md):
        try:
            msg = EmailMessage()
            msg['From'] = USER
            msg['To'] = to
            msg['Subject'] = subj
            msg.set_content(body_md)
            msg.add_alternative(markdown.markdown(body_md), subtype='html')
            if SMTP_PORT == 465:
                smtp = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT)
            else:
                smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT)
                if SMTP_USE_TLS:
                    smtp.starttls()
                    smtp.ehlo()
            smtp.login(USER, SMTP_PWD or PWD)
            smtp.send_message(msg)
            smtp.quit()
            messagebox.showinfo('Success', 'Email sent')
        except Exception:
            logging.exception('Send failed')
            messagebox.showerror('Error', 'Send failed')

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Login')
        self.geometry('400x180')
        self.resizable(False, False)
        self.user_var = tk.StringVar(value=USER)
        self.pass_var = tk.StringVar(value=PWD)
        ttk.Label(self, text='Username:').grid(row=0, column=0, padx=10, pady=10, sticky='e')
        ttk.Entry(self, textvariable=self.user_var, width=30).grid(row=0, column=1)
        ttk.Label(self, text='Password:').grid(row=1, column=0, padx=10, pady=10, sticky='e')
        ttk.Entry(self, textvariable=self.pass_var, show='*', width=30).grid(row=1, column=1)
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text='Login', command=self.do_login).pack(side='left', padx=10)
        ttk.Button(btn_frame, text='Login QQ', command=self.default_login).pack(side='left', padx=10)

    def do_login(self):
        global USER, PWD
        USER = self.user_var.get() or USER
        PWD = self.pass_var.get() or PWD
        self.destroy()
        POP3GUI().mainloop()

    def default_login(self):
        self.destroy()
        POP3GUI().mainloop()

if __name__ == '__main__':
    LoginWindow().mainloop()
