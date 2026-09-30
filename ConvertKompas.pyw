"""ConvertKompas: batch export KOMPAS v21 documents to earlier formats."""
import pathlib
import sys
import tempfile
import json
import os
import uuid
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / 'lib'))
import comtypes.client

comtypes.client.gen_dir = str(pathlib.Path(tempfile.gettempdir()) / 'ConvertKompas-com-cache')
pathlib.Path(comtypes.client.gen_dir).mkdir(parents=True, exist_ok=True)
KOMPAS_BIN = pathlib.Path(r'C:\Program Files\ASCON\KOMPAS-3D v21\Bin')
API5 = KOMPAS_BIN / 'kAPI5.tlb'
API7 = KOMPAS_BIN / 'kAPI7.tlb'
EXTENSIONS = {'.cdw', '.frw', '.m3d', '.a3d', '.spw', '.kdw'}
VERSIONS = {'16': 19, '16.1': 20, '17': 21, '17.1': 22,
            '18': 23, '18.1': 24, '19': 25, '20': 26}
REPORT = pathlib.Path(tempfile.gettempdir()) / 'ConvertKompas-last-run.log'
SETTINGS = pathlib.Path(os.environ.get('APPDATA') or pathlib.Path.home() / 'AppData' / 'Roaming') / 'ConvertKompas' / 'settings.json'


def load_settings(path=SETTINGS):
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data, path=SETTINGS):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def unique_destination(path, now=None):
    if not path.exists():
        return path
    stamp = (now or datetime.now()).strftime('%Y-%m-%d')
    base = '%s (%s)' % (path.stem, stamp)
    candidate = path.with_name(base + path.suffix)
    number = 2
    while candidate.exists():
        candidate = path.with_name('%s (%s)%s' % (base, number, path.suffix))
        number += 1
    return candidate


def version_matches(actual, expected):
    if actual == expected:
        return True
    # For v20, KOMPAS reports the native file-version word 0x1400001c.
    return expected == VERSIONS['20'] and (actual >> 24) == 20


def convert(source, version, log):
    if version not in VERSIONS:
        raise ValueError('Неизвестная версия результата: ' + version)
    save_mode = VERSIONS[version]
    if not API5.is_file() or not API7.is_file():
        raise RuntimeError('Не найдена установка КОМПАС-3D v21.')
    api5 = comtypes.client.GetModule(str(API5))
    api7 = comtypes.client.GetModule(str(API7))
    kompas = comtypes.client.CreateObject('KOMPAS.Application.5', interface=api5.KompasObject)
    try:
        app = kompas.ksGetApplication7().QueryInterface(api7.IApplication)
        count = 0
        failures = 0
        for src, dst in source:
            try:
                dst.parent.mkdir(parents=True, exist_ok=True)
                requested = dst
                dst = unique_destination(dst)
                if dst != requested:
                    log('Существующий файл сохранён; новый вариант: ' + str(dst))
                doc = app.Documents.Open(str(src), False, True)
                if doc is None:
                    raise RuntimeError('КОМПАС не открыл документ')
                try:
                    saved = doc.QueryInterface(api7.IKompasDocument1).SaveAsEx(str(dst), save_mode)
                finally:
                    doc.Close(0)
                if not saved or not dst.is_file() or dst.stat().st_size == 0:
                    raise RuntimeError('КОМПАС не создал файл')
                check = app.Documents.Open(str(dst), False, True)
                if check is None:
                    raise RuntimeError('Файл создан, но не открывается')
                try:
                    actual = check.QueryInterface(api7.IKompasDocument1).OpenVersion
                finally:
                    check.Close(0)
                if not version_matches(actual, save_mode):
                    raise RuntimeError('Файл создан, но версия отличается: ' + str(actual))
                count += 1
                log('Готово (v%s): %s' % (version, dst))
            except Exception as exc:
                failures += 1
                log('ОШИБКА ' + str(src) + ': ' + str(exc))
        return count, failures
    finally:
        kompas.Quit()


def jobs(input_path, output_path):
    if not input_path.strip():
        raise ValueError('Сначала выберите исходный файл или папку.')
    src = pathlib.Path(input_path).expanduser().resolve()
    out = pathlib.Path(output_path).expanduser().resolve()
    if not src.exists():
        raise ValueError('Исходная папка или файл не найдены.')
    if src.is_dir() and src == out:
        raise ValueError('Для результата нужна другая папка.')
    out.mkdir(parents=True, exist_ok=True)
    if src.is_file():
        if src.suffix.lower() not in EXTENSIONS:
            raise ValueError('Неподдерживаемое расширение файла.')
        files = [src]
        root = src.parent
    else:
        files = sorted(p for p in src.rglob('*') if p.is_file() and p.suffix.lower() in EXTENSIONS and not p.is_relative_to(out))
        root = src
    return [(p, out / p.relative_to(root)) for p in files]


def main():
    saved = load_settings()
    root = tk.Tk()
    root.title('ConvertKompas — КОМПАС v21 → ранние версии')
    root.geometry('800x550')
    root.minsize(650, 430)
    root.iconbitmap(str(HERE / 'ConvertKompas.ico'))
    background = '#F3F8FC'
    ink = '#173B57'
    root.configure(background=background)
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('App.TFrame', background=background)
    style.configure('App.TLabel', background=background, foreground=ink,
                    font=('Segoe UI', 10))
    style.configure('Title.TLabel', background=background, foreground='#0A3554',
                    font=('Segoe UI', 20, 'bold'))
    style.configure('Primary.TButton', background='#0879AE', foreground='white',
                    font=('Segoe UI', 11, 'bold'), padding=(16, 12),
                    borderwidth=0, relief='flat')
    style.map('Primary.TButton',
              background=[('disabled', '#9BAFBD'), ('pressed', '#07577D'), ('active', '#1091C9')],
              foreground=[('disabled', '#F4F7F9'), ('active', 'white')])
    style.configure('Secondary.TButton', background='#E2F1F9', foreground='#124C70',
                    font=('Segoe UI', 10, 'bold'), padding=(12, 7),
                    borderwidth=0, relief='flat')
    style.map('Secondary.TButton',
              background=[('pressed', '#BDDDEE'), ('active', '#CDE9F6')],
              foreground=[('active', '#0B3D5B')])
    style.configure('ActiveTab.TButton', background='#1091C9', foreground='white',
                    font=('Segoe UI', 11, 'bold'), padding=(18, 10),
                    borderwidth=0, relief='flat')
    style.map('ActiveTab.TButton', background=[('pressed', '#07577D'), ('active', '#20ABE2')],
              foreground=[('active', 'white')])
    style.configure('InactiveTab.TButton', background='#0879AE', foreground='white',
                    font=('Segoe UI', 11, 'bold'), padding=(18, 10),
                    borderwidth=0, relief='flat')
    style.map('InactiveTab.TButton', background=[('pressed', '#07577D'), ('active', '#20ABE2')],
              foreground=[('active', 'white')])
    input_var = tk.StringVar(value=saved.get('input', '') if isinstance(saved.get('input'), str) else '')
    output_var = tk.StringVar(value=saved.get('output', '') if isinstance(saved.get('output'), str) else '')
    selected_version = saved.get('version', '16.1')
    version_var = tk.StringVar(value=selected_version if selected_version in VERSIONS else '16.1')
    auto_output = {'enabled': bool(saved.get('output_auto', False)),
                   'path': output_var.get() if saved.get('output_auto', False) else None}
    internal_output_update = {'active': False}

    tab_bar = ttk.Frame(root, style='App.TFrame')
    tab_bar.pack(fill='x', padx=12, pady=(12, 0))
    tab_bar.columnconfigure(0, weight=1, uniform='equal_tabs')
    tab_bar.columnconfigure(1, weight=1, uniform='equal_tabs')
    content = ttk.Frame(root, style='App.TFrame')
    content.pack(fill='both', expand=True)
    content.rowconfigure(0, weight=1)
    content.columnconfigure(0, weight=1)
    conversion_tab = ttk.Frame(content, style='App.TFrame')
    about_tab = ttk.Frame(content, style='App.TFrame')
    conversion_tab.grid(row=0, column=0, sticky='nsew')
    about_tab.grid(row=0, column=0, sticky='nsew')

    def show_tab(name):
        (conversion_tab if name == 'conversion' else about_tab).tkraise()
        conversion_button.configure(style='ActiveTab.TButton' if name == 'conversion' else 'InactiveTab.TButton')
        about_button.configure(style='ActiveTab.TButton' if name == 'about' else 'InactiveTab.TButton')

    conversion_button = ttk.Button(tab_bar, text='Конвертация',
                                   command=lambda: show_tab('conversion'))
    about_button = ttk.Button(tab_bar, text='О программе',
                              command=lambda: show_tab('about'))
    conversion_button.grid(row=0, column=0, sticky='ew', padx=(0, 2))
    about_button.grid(row=0, column=1, sticky='ew', padx=(2, 0))
    show_tab('conversion')

    about_icon = tk.PhotoImage(file=str(HERE / 'ConvertKompas-preview.png'))
    icon_label = tk.Label(about_tab, image=about_icon, background=background)
    icon_label.image = about_icon
    icon_label.pack(pady=(28, 10))
    ttk.Label(about_tab, text='ConvertKompas', style='Title.TLabel').pack()
    ttk.Label(about_tab, text='Пакетное сохранение документов КОМПАС-3D v21 в ранних версиях',
              style='App.TLabel').pack(pady=(4, 20))
    ttk.Label(about_tab, text='Разработчик: Zabir', style='App.TLabel').pack()
    ttk.Label(about_tab, text='Поддержать USDT TRC20 - TPU4Kt3nBJ2WCkXu9FqzprP3AeHxhEtejo',
              style='App.TLabel').pack(pady=(14, 0))

    ttk.Label(conversion_tab, text='Исходная папка или файл КОМПАС v21', style='App.TLabel').pack(anchor='w', padx=12, pady=(12, 4))
    row = ttk.Frame(conversion_tab, style='App.TFrame'); row.pack(fill='x', padx=12)
    ttk.Entry(row, textvariable=input_var).pack(side='left', fill='x', expand=True, ipady=5)
    ttk.Button(row, text='Папка…', style='Secondary.TButton', command=lambda: input_var.set(filedialog.askdirectory() or input_var.get())).pack(side='left', padx=(8, 4))
    ttk.Button(row, text='Файл…', style='Secondary.TButton', command=lambda: input_var.set(filedialog.askopenfilename() or input_var.get())).pack(side='left')

    version_row = ttk.Frame(conversion_tab, style='App.TFrame')
    version_row.pack(fill='x', padx=12, pady=(14, 0))
    ttk.Label(version_row, text='Версия результата:', style='App.TLabel').pack(side='left')
    version_combo = ttk.Combobox(version_row, textvariable=version_var,
                                 values=tuple(VERSIONS), state='readonly', width=9)
    version_combo.pack(side='left', padx=(10, 0))

    ttk.Label(conversion_tab, text='Папка результата', style='App.TLabel').pack(anchor='w', padx=12, pady=(14, 4))
    row = ttk.Frame(conversion_tab, style='App.TFrame'); row.pack(fill='x', padx=12)
    ttk.Entry(row, textvariable=output_var).pack(side='left', fill='x', expand=True, ipady=5)
    ttk.Button(row, text='Обзор…', style='Secondary.TButton', command=lambda: output_var.set(filedialog.askdirectory() or output_var.get())).pack(side='left', padx=(8, 0))

    action_row = ttk.Frame(conversion_tab, style='App.TFrame')
    action_row.pack(fill='x', padx=12, pady=(12, 0))

    logbox = scrolledtext.ScrolledText(conversion_tab, state='disabled', wrap='word',
                                      background='white', foreground=ink,
                                      font=('Consolas', 10), relief='flat',
                                      highlightthickness=1, highlightbackground='#CADCE8')
    logbox.pack(fill='both', expand=True, padx=12, pady=(10, 12))

    def log(message):
        with REPORT.open('a', encoding='utf-8') as report:
            report.write(message + '\n')
        logbox.configure(state='normal')
        logbox.insert('end', message + '\n')
        logbox.see('end')
        logbox.configure(state='disabled')
        root.update_idletasks()

    save_timer = {'id': None}

    def persist_settings():
        save_timer['id'] = None
        try:
            save_settings({'input': input_var.get(),
                           'output': output_var.get(),
                           'version': version_var.get(),
                           'output_auto': auto_output['enabled']})
        except OSError as exc:
            log('Не удалось сохранить выбранные папки: ' + str(exc))

    def schedule_settings(*_):
        if save_timer['id'] is not None:
            root.after_cancel(save_timer['id'])
        save_timer['id'] = root.after(400, persist_settings)

    def set_auto_output(path):
        internal_output_update['active'] = True
        try:
            output_var.set(path)
        finally:
            internal_output_update['active'] = False
        auto_output['enabled'] = True
        auto_output['path'] = path

    def on_output_change(*_):
        if not internal_output_update['active']:
            auto_output['enabled'] = False
            auto_output['path'] = None
        schedule_settings()

    def run():
        try:
            src = pathlib.Path(input_var.get()).expanduser().resolve() if input_var.get().strip() else None
            selected = version_var.get()
            if selected not in VERSIONS:
                raise ValueError('Выберите версию результата из списка.')
            if src is not None and (not output_var.get().strip() or auto_output['enabled']):
                set_auto_output(str(src.parent / ('КОМПАС ' + selected)))
            persist_settings()
            REPORT.write_text('ConvertKompas ' + datetime.now().isoformat() + '\n', encoding='utf-8')
            log('Исходный путь: ' + input_var.get())
            log('Папка результата: ' + output_var.get())
            log('Версия результата: ' + selected)
            batch = jobs(input_var.get(), output_var.get())
            if not batch:
                raise ValueError('В исходной папке нет документов КОМПАС.')
            button.configure(state='disabled')
            log('Найдено файлов: ' + str(len(batch)))
            root.update()
            count, errors = convert(batch, selected, log)
            log('Завершено. Сохранено: %s; ошибок: %s.' % (count, errors))
            log('Отчёт: ' + str(REPORT))
            messagebox.showinfo('ConvertKompas', 'Сохранено: %s\nОшибок: %s' % (count, errors))
        except Exception as exc:
            log('ОШИБКА: ' + str(exc))
            messagebox.showerror('ConvertKompas', str(exc))
        finally:
            button.configure(state='normal')

    if input_var.get() or output_var.get():
        log('Ранее выбранные папки восстановлены. При необходимости измените их вручную.')
    else:
        log('Выберите исходный файл или папку. Папку результата можно не указывать: она создастся рядом автоматически.')
    button = ttk.Button(action_row, text='Начать конвертацию в КОМПАС v16.1',
                        style='Primary.TButton', command=run)
    button.pack(fill='x')

    def update_version(*_):
        button.configure(text='Начать конвертацию в КОМПАС v' + version_var.get())
        if auto_output['enabled'] and input_var.get().strip():
            source = pathlib.Path(input_var.get()).expanduser().resolve()
            set_auto_output(str(source.parent / ('КОМПАС ' + version_var.get())))
        schedule_settings()

    def on_close():
        if save_timer['id'] is not None:
            root.after_cancel(save_timer['id'])
        persist_settings()
        root.destroy()

    input_var.trace_add('write', schedule_settings)
    output_var.trace_add('write', on_output_change)
    version_var.trace_add('write', update_version)
    root.protocol('WM_DELETE_WINDOW', on_close)
    update_version()
    root.mainloop()


if __name__ == '__main__':
    if len(sys.argv) in (4, 5) and sys.argv[1] == '--batch':
        selected = sys.argv[4] if len(sys.argv) == 5 else '16.1'
        batch = jobs(sys.argv[2], sys.argv[3])
        converted, failed = convert(batch, selected, print)
        print('converted=%s failed=%s' % (converted, failed))
        sys.exit(0 if failed == 0 and converted > 0 else 1)
    main()
