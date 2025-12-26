import sys
import os
import json
import subprocess
import threading
import shutil
from pathlib import Path

try:
    import webview
except Exception:
    webview = None

try:
    from plyer import notification as plyer_notify
except Exception:
    plyer_notify = None

try:
    from win10toast_click import ToastNotifier
except Exception:
    ToastNotifier = None

CONFIG_FILE = Path.home() / '.zip-installer_config.json'


def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {"target_path": str(Path.home() / 'zip-installerInstalls')}


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


def ensure_dir(path):
    Path(path).mkdir(parents=True, exist_ok=True)


def is_elevated():
    if os.name != 'nt':
        return True
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def relaunch_as_admin(args):
    import ctypes
    params = ' '.join(['"%s"' % a for a in args])
    executable = sys.executable
    ctypes.windll.shell32.ShellExecuteW(None, 'runas', executable, params, None, 1)


class ProgressAPI:
    def __init__(self):
        self.progress = 0
        self.message = ''

    def update(self, p, msg=''):
        try:
            p = int(p)
        except Exception:
            p = 0
        self.progress = max(0, min(100, p))
        self.message = msg


def send_notifications(title, message, folder):
    # plyer notification (cross-platform)
    if plyer_notify:
        try:
            plyer_notify.notify(title=title, message=message)
        except Exception:
            pass

    # On Windows use win10toast_click to provide clickable toast
    if ToastNotifier and os.name == 'nt':
        try:
            toaster = ToastNotifier()
            def _open():
                subprocess.Popen(['explorer', str(folder)])
            toaster.show_toast(title, message, icon_path=None, duration=10, threaded=True, callback_on_click=_open)
        except Exception:
            pass


def extract_zip_with_progress(src, dest, api=None):
    import zipfile
    with zipfile.ZipFile(src, 'r') as zf:
        namelist = zf.namelist()
        total = len(namelist)
        for i, name in enumerate(namelist, 1):
            target_path = Path(dest) / name
            if name.endswith('/'):
                target_path.mkdir(parents=True, exist_ok=True)
            else:
                ensure_dir(target_path.parent)
                with zf.open(name) as source, open(target_path, 'wb') as out:
                    shutil.copyfileobj(source, out)
            if api:
                api.update(int(i / total * 100), f'Extracting {name}')


def extract_with_7z(src, dest, api=None):
    # Use 7z if available (7z.exe in PATH)
    cmd = ['7z', 'x', str(src), '-o' + str(dest), '-y']
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)
    except FileNotFoundError:
        raise RuntimeError('7z not found in PATH')

    for line in proc.stdout:
        line = line.strip()
        # attempt to parse percentage
        if '%' in line:
            try:
                p = int(line.split('%')[-2].split()[-1])
                if api:
                    api.update(p, line)
            except Exception:
                pass
    proc.wait()
    if proc.returncode != 0:
        raise RuntimeError('7z failed')


def extract_archive(src, dest, api=None):
    ext = Path(src).suffix.lower()
    if ext == '.zip':
        extract_zip_with_progress(src, dest, api)
    else:
        # try 7z for rar/7z/others
        extract_with_7z(src, dest, api)


def confirm_replace(target):
    # simple CLI fallback confirmation when no GUI
    if any(Path(target).iterdir()):
        resp = input(f'目标目录 {target} 非空，是否覆盖? (y/n): ')
        return resp.lower().startswith('y')
    return True


def run_install(archive_path):
    cfg = load_config()
    base = Path(archive_path).stem
    target_root = Path(cfg.get('target_path'))
    target_dir = target_root / base
    ensure_dir(target_root)
    if target_dir.exists() and any(target_dir.iterdir()):
        # ask user
        # If pywebview present we could show a dialog, but here do a simple console prompt
        if webview:
            # show a tiny webview confirm dialog
            def confirm_thread():
                webview.create_window('确认', html=f"<h3>目标目录已存在: {target_dir}</h3><p>是否覆盖？</p>", width=400, height=200)
            t = threading.Thread(target=confirm_thread, daemon=True)
            t.start()
            # fallback to CLI
            if not confirm_replace(target_dir):
                print('取消')
                return
        else:
            if not confirm_replace(target_dir):
                print('取消')
                return
        shutil.rmtree(target_dir)

    ensure_dir(target_dir)
    api = ProgressAPI()

    # start webview progress window if possible
    if webview:
        # load external HTML template to avoid f-string quoting issues
        tpl_path = Path(__file__).parent / './html/progress.html'
        if tpl_path.exists():
            html = tpl_path.read_text(encoding='utf-8')
        else:
            # fallback minimal HTML
            html = """
            <html><body><h3 id='msg'>准备解压</h3><div style='width:80%;background:#eee;border-radius:8px;padding:3px'><div id='bar' style='width:0%;height:24px;background:linear-gradient(90deg,#4caf50,#8bc34a);border-radius:6px'></div></div><script>function setProgress(p,msg){document.getElementById('bar').style.width = p+'%';document.getElementById('msg').innerText = msg + ' ('+p+'%)';}</script></body></html>
            """
        win = webview.create_window('解压进度', html=html, width=500, height=200)

        def background_extract():
            try:
                extract_archive(archive_path, target_dir, api)
                send_notifications('安装完成', f'{archive_path} 已解压到 {target_dir}', target_dir)
            except Exception as e:
                send_notifications('安装失败', str(e), target_dir)

        threading.Thread(target=background_extract, daemon=True).start()
        # poll to update UI
        def update_poll():
            while True:
                # use json.dumps to safely escape the message string for JS
                webview.evaluate_js(win, f"setProgress({api.progress}, {json.dumps(api.message)})")
                if api.progress >= 100:
                    break
                import time
                time.sleep(0.5)

        threading.Thread(target=update_poll, daemon=True).start()
        webview.start()
    else:
        # No GUI, do extraction and print progress
        try:
            extract_archive(archive_path, target_dir, api)
            send_notifications('安装完成', f'{archive_path} 已解压到 {target_dir}', target_dir)
        except Exception as e:
            send_notifications('安装失败', str(e), target_dir)


def show_settings():
    cfg = load_config()
    if webview:
        tpl_path = Path(__file__).parent / './html/settings.html'
        if tpl_path.exists():
            html = tpl_path.read_text(encoding='utf-8')
        else:
            # fallback minimal HTML
            html = """
            <html><body><h3>设置目标路径</h3><input id='path' /><button onclick="window.pywebview.api.save(document.getElementById('path').value)">保存</button></body></html>
            """

        class Api:
            def get_initial(self):
                return cfg.get('target_path')

            def save(self, p):
                cfg['target_path'] = p
                save_config(cfg)
                return True

            def choose(self):
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    path = filedialog.askdirectory()
                    root.destroy()
                    return path or ''
                except Exception:
                    return ''

        api = Api()
        webview.create_window('zip-installer 设置', html=html, js_api=api, width=700, height=600)
        webview.start()
    else:
        # CLI fallback: allow typing path or open dialog via tkinter
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            print(f'当前目标路径: {cfg.get("target_path")}')
            if input('按回车使用当前路径，或输入 c 调用目录选择: ').strip().lower() == 'c':
                p = filedialog.askdirectory()
                root.destroy()
                if p:
                    cfg['target_path'] = p
                    save_config(cfg)
            else:
                root.destroy()
        except Exception:
            p = input(f'目标路径 (回车使用 {cfg.get("target_path")}): ')
            if p.strip():
                cfg['target_path'] = p.strip()
                save_config(cfg)


def register_context_menu():
    # Register per-user context menu entries for .zip .rar .7z
    if os.name != 'nt':
        print('仅支持 Windows 注册右键菜单')
        return
    import winreg
    exe = sys.executable
    # if packaged, sys.argv[0] might be the exe
    script = Path(sys.argv[0]).resolve()
    cmd = f'"{exe}" "{script}" --install "%1"'
    exts = ['.zip', '.rar', '.7z']
    for e in exts:
        try:
            key_path = f'Software\\Classes\\{e}\\shell\\从此文件安装\\command'
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                winreg.SetValueEx(k, None, 0, winreg.REG_SZ, cmd)
            print(f'已为 {e} 注册右键菜单')
        except Exception as ex:
            print('注册失败', ex)


def main():
    if len(sys.argv) >= 2 and sys.argv[1] == '--install':
        if len(sys.argv) < 3:
            print('缺少文件路径')
            return
        archive = sys.argv[2]
        # check privileges for target path
        cfg = load_config()
        target_root = Path(cfg.get('target_path'))
        if not os.access(str(target_root), os.W_OK):
            if not is_elevated():
                # relaunch as admin
                relaunch_as_admin(sys.argv)
                return
        run_install(archive)
        return

    # no args -> settings and registration helper
    if '--register' in sys.argv:
        register_context_menu()
        return

    show_settings()


if __name__ == '__main__':
    main()
