import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import webbrowser, os, tempfile, time, windnd, re, winsound, threading, requests, json
import pandas as pd
from docx import Document

# --- ПІДКЛЮЧЕННЯ ДО GITHUB ---
VERSION = "7.8"
VERSION_URL = "https://raw.githubusercontent.com/alex-voron/LinkerPro/refs/heads/main/version.txt"
FILE_URL = "https://raw.githubusercontent.com/alex-voron/LinkerPro/refs/heads/main/linker_pro.pyw"

class LinkerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Linker Pro v{VERSION}")
        self.root.geometry("520x650")
        self.root.resizable(False, False)
        
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(self.script_dir, "config.json")
        
        self.lang_data = {
            'ukr': {
                'zone': "📥\n\nПЕРЕТЯГНІТЬ ФАЙЛ СЮДИ\n(TXT, XLSX, DOCX)\n\nабо НАТИСНІТЬ",
                'zone_drop': "🚀\n\nВІДПУСТІТЬ ФАЙЛ",
                'dedup': "Видаляти дублікати",
                'ping': "Перевірити Ping (LIVE/DEAD)",
                'clear': "Очистити результат",
                'last_res': "СТАТИСТИКА ОБРОБКИ",
                'found': "ЗНАЙДЕНО",
                'unique': "ВАЛІДНИХ",
                'errors': "ПОМИЛКИ",
                'copy_btn': "📋 КОПІЮВАТИ СПИСОК",
                'copied': "✅ СКОПІЙОВАНО",
                'update_found': "Доступна нова версія!",
                'update_no': "У вас остання версія",
                'update_err': "Помилка перевірки оновлень",
                'err_title': "Знайдено помилкові рядки:",
                'line': "Рядок",
                'total_html': "Всього"
            },
            'eng': { 
                'zone': "📥\n\nDROP FILE HERE\n\nor CLICK", 
                'zone_drop': "🚀\n\nRELEASE", 
                'dedup': "Remove duplicates", 
                'ping': "Check Status", 
                'clear': "Clear", 
                'last_res': "STATS", 
                'found': "FOUND", 
                'unique': "VALID", 
                'errors': "ERRORS", 
                'copy_btn': "📋 COPY", 
                'copied': "✅ COPIED", 
                'update_found': "Update available!", 
                'update_no': "Up to date", 
                'update_err': "Update check failed", 
                'err_title': "Invalid lines:", 
                'line': "Line",
                'total_html': "Total"
            },
            'rus': { 
                'zone': "📥\n\nПЕРЕТЯНИТЕ ФАЙЛ СЮДА\n\nили НАЖМИТЕ", 
                'zone_drop': "🚀\n\nОТПУСТИТЕ", 
                'dedup': "Удалять дубликаты", 
                'ping': "Проверить Ping", 
                'clear': "Очистить", 
                'last_res': "СТАТИСТИКА", 
                'found': "НАЙДЕНО", 
                'unique': "ВАЛИДНЫХ", 
                'errors': "ОШИБКИ", 
                'copy_btn': "📋 КОПИРОВАТЬ", 
                'copied': "✅ СКОПИРОВАНО", 
                'update_found': "Доступно обновление!", 
                'update_no': "Последняя версия", 
                'update_err': "Ошибка обновления", 
                'err_title': "Ошибочные строки:", 
                'line': "Строка",
                'total_html': "Всего"
            }
        }
        
        self.curr_lang = 'ukr'
        self.dark_mode = False
        self.colors = {
            'light': {'bg': '#f0f2f5', 'fg': '#1a1d21', 'card': '#ffffff', 'zone': '#ffffff', 'accent': '#0061ff', 'stat_bg': '#ffffff', 'border': '#ced4da'},
            'dark': {'bg': '#1a1d21', 'fg': '#f8f9fa', 'card': '#2d3339', 'zone': '#2d3339', 'accent': '#3391ff', 'stat_bg': '#2d3339', 'border': '#3e444a'}
        }
        
        self.setup_ui()
        self.init_from_config()
        
        # Реєстрація Drag-and-Drop
        self.root.bind("<Visibility>", lambda e: [windnd.hook_dropfiles(self.root, func=self.handle_drop), self.root.unbind("<Visibility>")])
        
        # Перевірка оновлень при старті
        threading.Thread(target=self.check_updates, args=(False,), daemon=True).start()

    def check_updates(self, manual=False):
        try:
            res = requests.get(VERSION_URL, timeout=5)
            online_v = res.text.strip()
            # Порівнюємо версії
            if online_v > VERSION:
                if messagebox.askyesno("Update", f"{self.lang_data[self.curr_lang]['update_found']} (v{online_v})\nDownload now?"):
                    webbrowser.open(FILE_URL)
            elif manual:
                messagebox.showinfo("Update", self.lang_data[self.curr_lang]['update_no'])
        except:
            if manual:
                messagebox.showwarning("Update", self.lang_data[self.curr_lang]['update_err'])

    def setup_ui(self):
        self.root.configure(bg=self.colors['light']['bg'])
        self.top_bar = tk.Frame(self.root, bg=self.colors['light']['bg'])
        self.top_bar.pack(fill="x", padx=25, pady=(15, 5))
        
        self.lang_frame = tk.Frame(self.top_bar, bg=self.colors['light']['bg'])
        self.lang_frame.pack(side="left")
        for l in ['ukr', 'eng', 'rus']:
            btn = tk.Button(self.lang_frame, text=l.upper(), font=("Segoe UI", 8, "bold"), command=lambda lang=l: self.set_language(lang), relief="flat", padx=8, pady=4)
            btn.pack(side="left", padx=2); self.apply_hover(btn, True)

        self.theme_btn = tk.Button(self.top_bar, text="🌙", font=("Segoe UI", 11), command=lambda: self.toggle_theme(True), relief="flat", width=4)
        self.theme_btn.pack(side="right", padx=2); self.apply_hover(self.theme_btn, True)

        self.upd_btn = tk.Button(self.top_bar, text="🔄", font=("Segoe UI", 11), command=lambda: self.check_updates(True), relief="flat", width=4)
        self.upd_btn.pack(side="right", padx=2); self.apply_hover(self.upd_btn, True)

        self.action_zone = tk.Label(self.root, text="", font=("Segoe UI Variable Text", 11, "bold"), bg=self.colors['light']['zone'], fg="#6c757d", relief="flat", highlightthickness=1, highlightbackground=self.colors['light']['border'], cursor="hand2", height=7)
        self.action_zone.pack(fill="x", padx=30, pady=15)
        self.action_zone.bind("<Button-1>", lambda e: self.browse_file())

        self.opt_frame = tk.Frame(self.root, bg=self.colors['light']['bg'])
        self.opt_frame.pack(fill="x", padx=35)
        self.dedup_var, self.ping_var = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
        self.check_dedup = tk.Checkbutton(self.opt_frame, text="", variable=self.dedup_var, bg=self.colors['light']['bg'], activebackground=self.colors['light']['bg'])
        self.check_dedup.pack(anchor="w")
        self.check_ping = tk.Checkbutton(self.opt_frame, text="", variable=self.ping_var, bg=self.colors['light']['bg'], activebackground=self.colors['light']['bg'])
        self.check_ping.pack(anchor="w")

        self.style = ttk.Style(); self.style.theme_use('clam')
        self.style.configure("Modern.Horizontal.TProgressbar", troughcolor='#ced4da', bordercolor='#ced4da', background='#0061ff', thickness=8)
        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate", style="Modern.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=35, pady=20)

        self.res_title = tk.Label(self.root, text="", font=("Segoe UI", 9, "bold"), bg=self.colors['light']['bg'], fg="#8a94a6"); self.res_title.pack(pady=5)
        self.stats_container = tk.Frame(self.root, bg=self.colors['light']['bg']); self.stats_container.pack(fill="x", padx=30)
        self.tile_found = self.create_tile(self.stats_container, 'found', "#1a1d21")
        self.tile_unique = self.create_tile(self.stats_container, 'unique', "#0061ff")
        self.tile_errors = self.create_tile(self.stats_container, 'errors', "#e03131")

        self.clear_btn = tk.Button(self.root, text="", command=self.clear_results, font=("Segoe UI", 8, "bold"), relief="flat", bg="#ced4da", fg="#495057", padx=25, pady=5)
        self.clear_btn.pack(pady=25); self.apply_hover(self.clear_btn, False)

    def create_tile(self, parent, key, color):
        f = tk.Frame(parent, bg=self.colors['light']['stat_bg'], highlightthickness=1, highlightbackground=self.colors['light']['border'], padx=10, pady=12)
        f.pack(side="left", expand=True, padx=6)
        v = tk.Label(f, text="0", font=("Segoe UI Variable Display", 18, "bold"), bg=self.colors['light']['stat_bg'], fg=color); v.pack()
        d = tk.Label(f, text="", font=("Segoe UI", 7, "bold"), bg=self.colors['light']['stat_bg'], fg="#8a94a6"); d.pack()
        f.desc_key = key
        return v

    def apply_hover(self, widget, is_top):
        def on_e(e): widget.config(bg=self.colors['dark' if self.dark_mode else 'light']['border'])
        def on_l(e): widget.config(bg=self.colors['dark' if self.dark_mode else 'light']['bg'] if is_top else ("#ced4da" if not self.dark_mode else self.colors['dark']['zone']))
        widget.bind("<Enter>", on_e); widget.bind("<Leave>", on_l)

    def set_language(self, lang):
        self.curr_lang = lang; d = self.lang_data[lang]
        self.action_zone.config(text=d['zone']); self.check_dedup.config(text=d['dedup']); self.check_ping.config(text=d['ping'])
        self.clear_btn.config(text=d['clear']); self.res_title.config(text=d['last_res'])
        for v in [self.tile_found, self.tile_unique, self.tile_errors]: v.master.winfo_children()[1].config(text=d[v.master.desc_key])
        self.save_config()

    def toggle_theme(self, save=True):
        self.dark_mode = not self.dark_mode
        c = self.colors['dark' if self.dark_mode else 'light']
        self.root.configure(bg=c['bg']); self.top_bar.config(bg=c['bg']); self.lang_frame.config(bg=c['bg'])
        self.action_zone.config(bg=c['zone'], fg=c['fg'] if self.dark_mode else "#6c757d", highlightbackground=c['border'])
        for ck in [self.check_dedup, self.check_ping]: ck.config(bg=c['bg'], fg=c['fg'], selectcolor=c['bg'] if not self.dark_mode else "#2d3339", activebackground=c['bg'])
        self.opt_frame.config(bg=c['bg']); self.stats_container.config(bg=c['bg']); self.res_title.config(bg=c['bg'])
        self.clear_btn.config(bg=c['zone'] if self.dark_mode else "#ced4da", fg=c['fg'] if self.dark_mode else "#495057")
        for v in [self.tile_found, self.tile_unique, self.tile_errors]:
            v.config(bg=c['stat_bg']); v.master.config(bg=c['stat_bg'], highlightbackground=c['border']); v.master.winfo_children()[1].config(bg=c['stat_bg'])
            v.config(fg=("#3391ff" if v == self.tile_unique else ("#ff6b6b" if v == self.tile_errors else "#f8f9fa")) if self.dark_mode else ("#0061ff" if v == self.tile_unique else ("#e03131" if v == self.tile_errors else "#1a1d21")))
        for btn in self.lang_frame.winfo_children(): btn.config(bg=c['bg'], fg=c['fg'])
        self.theme_btn.config(bg=c['bg'], fg=c['fg'], text="☀️" if self.dark_mode else "🌙"); self.upd_btn.config(bg=c['bg'], fg=c['fg'])
        self.style.configure("Modern.Horizontal.TProgressbar", troughcolor=c['border'], bordercolor=c['border'])
        if save: self.save_config()

    def init_from_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f); self.curr_lang = cfg.get('lang', 'ukr')
                    if cfg.get('dark_mode', False): self.toggle_theme(False)
                    self.set_language(self.curr_lang)
            except: self.set_language('ukr')
        else: self.set_language('ukr')

    def save_config(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f: json.dump({'lang': self.curr_lang, 'dark_mode': self.dark_mode}, f)
        except: pass

    def clear_results(self):
        for v in [self.tile_found, self.tile_unique, self.tile_errors]: v.config(text="0")
        self.progress['value'] = 0

    def handle_drop(self, files):
        paths = [f.decode('utf-8') if isinstance(f, bytes) else str(f) for f in files]
        if paths:
            self.action_zone.config(bg=self.colors['dark' if self.dark_mode else 'light']['border'], text=self.lang_data[self.curr_lang]['zone_drop'])
            threading.Thread(target=self.process_file, args=(paths[0],), daemon=True).start()

    def browse_file(self):
        p = filedialog.askopenfilename()
        if p: threading.Thread(target=self.process_file, args=(p,), daemon=True).start()

    def process_file(self, path):
        try:
            filename, ext = os.path.basename(path), os.path.splitext(path)[1].lower()
            if ext == '.txt':
                try: 
                    with open(path, "r", encoding="utf-8") as f: lines = f.readlines()
                except: 
                    with open(path, "r", encoding="cp1251") as f: lines = f.readlines()
            elif ext in ['.xlsx', '.xls']: lines = pd.read_excel(path, header=None).stack().tolist()
            elif ext == '.docx': lines = [p.text for p in Document(path).paragraphs]
            else: return
            valid, errors = [], []
            url_p = r'https?://[^\s/$.?#].[^\s<>"]+'
            for idx, raw in enumerate(lines, 1):
                clean = str(raw).strip()
                if not clean or clean.lower() == 'nan': continue
                match = re.search(url_p, clean)
                if match: valid.append(match.group(0))
                else: errors.append({'num': idx, 'content': clean})
            final = list(dict.fromkeys(valid)) if self.dedup_var.get() else valid
            results = []
            for i, link in enumerate(final):
                st = "—"
                if self.ping_var.get():
                    try: res = requests.head(link, timeout=3, allow_redirects=True); st = "LIVE" if res.status_code < 400 else f"ERR {res.status_code}"
                    except: st = "DEAD"
                results.append({'url': link, 'status': st})
                self.progress['value'] = (i+1)/len(final)*100; self.root.update_idletasks()
            self.tile_found.config(text=str(len(lines))); self.tile_unique.config(text=str(len(final))); self.tile_errors.config(text=str(len(errors)))
            out = os.path.join(tempfile.gettempdir(), f"linker_res_{int(time.time())}.html")
            with open(out, "w", encoding="utf-8") as f: f.write(self.make_html(results, errors, filename))
            webbrowser.open('file:///' + out.replace("\\", "/")); winsound.MessageBeep()
            self.action_zone.config(bg=self.colors['dark' if self.dark_mode else 'light']['zone'], text=self.lang_data[self.curr_lang]['zone'])
        except Exception as e: messagebox.showerror("Error", str(e))

    def make_html(self, links, errors, filename):
        d = self.lang_data[self.curr_lang]
        js = f"function copyAll(){{const t=Array.from(document.querySelectorAll('.link-url')).map((el,i)=>(i+1)+'. '+el.innerText).join('\\n');navigator.clipboard.writeText(t).then(()=>{{const b=document.getElementById('cbtn');b.innerText='{d['copied']}';setTimeout(()=>b.innerText='{d['copy_btn']}',2000);}});}}"
        html = f"<html><head><meta charset='UTF-8'><style>body{{font-family:'Segoe UI',sans-serif;margin:0;background:#f0f2f5;display:flex;flex-direction:column;align-items:center}}.card{{background:white;width:100%;max-width:850px;min-height:100vh;box-shadow:0 10px 30px rgba(0,0,0,0.05);padding-bottom:50px}}.header{{position:fixed;top:0;width:100%;max-width:850px;background:white;border-bottom:1px solid #eee;padding:20px 40px;box-sizing:border-box;z-index:9999;display:flex;justify-content:space-between;align-items:center}}.list-container{{padding:120px 40px 40px 40px}}.link-item{{padding:12px 0;border-bottom:1px solid #f8f9fa;display:flex;align-items:center}}.num{{font-weight:bold;color:#adb5bd;width:40px;font-size:14px}}.status{{font-size:10px;font-weight:bold;padding:4px 8px;border-radius:4px;margin-right:20px;width:50px;text-align:center}}.LIVE{{background:#e8f5e9;color:#2e7d32}}.DEAD,.ERR{{background:#fff5f5;color:#e03131}}.link-url{{color:#0061ff;text-decoration:none;word-break:break-all;font-size:15px}}.copy-btn{{background:#0061ff;color:white;border:none;padding:12px 25px;border-radius:8px;cursor:pointer;font-weight:bold}}.error-section{{margin-top:40px;padding:25px;background:#fff5f5;border-radius:12px}}.err-tag{{font-weight:bold;background:#e03131;color:white;padding:3px 8px;border-radius:4px;margin-right:15px;font-size:11px}}</style><script>{js}</script></head><body><div class='card'><div class='header'><div><h2 style='margin:0;font-size:1.2em;'>{filename}</h2><small style='color:#8a94a6;'>{d['total_html']}: {len(links)}</small></div><button id='cbtn' class='copy-btn' onclick='copyAll()'>{d['copy_btn']}</button></div><div class='list-container'>"
        for i, item in enumerate(links, 1):
            st_type = "LIVE" if "LIVE" in item['status'] else "DEAD"
            html += f"<div class='link-item'><span class='num'>{i}</span><span class='status {st_type}'>{item['status']}</span><a href='{item['url']}' class='link-url' target='_blank'>{item['url']}</a></div>"
        if errors:
            html += f"<div class='error-section'><h3 style='color:#e03131;margin-top:0;'>⚠️ {d['err_title']}</h3>"
            for err in errors: html += f"<div style='margin-bottom:8px;font-size:14px;'><span class='err-tag'>{d['line']} {err['num']}</span>{err['content']}</div>"
        return html + "</div></div></body></html>"

if __name__ == "__main__":
    root = tk.Tk(); app = LinkerApp(root); root.mainloop()