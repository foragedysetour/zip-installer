import os
import sys
import json
import threading
import subprocess
import shutil
import re
from pathlib import Path

try:
	import tkinter as tk
	from tkinter import ttk, filedialog, messagebox
except Exception:
	tk = None

CONFIG_FILE = Path(__file__).with_name('config.json')


def load_config():
	if CONFIG_FILE.exists():
		try:
			return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
		except Exception:
			return {}
	return {}


def save_config(cfg):
	CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


def find_7z():
	from shutil import which
	exe = which('7z') or which('7z.exe')
	if exe:
		return exe
	# try common install path
	program_files = os.environ.get('ProgramFiles', r'C:\Program Files')
	candidate = Path(program_files) / '7-Zip' / '7z.exe'
	if candidate.exists():
		return str(candidate)
	return None


def ensure_base_path(cfg):
	base = cfg.get('base_path')
	if base and Path(base).exists():
		return base
	return None


def extract_with_7z(archive, dest_dir, progress_callback=None):
	exe = find_7z()
	if not exe:
		raise RuntimeError('7z.exe 未找到，请安装 7-Zip 并将其添加到 PATH。')

	dest_dir = str(dest_dir)
	cmd = [exe, 'x', str(archive), f'-o{dest_dir}', '-y']
	# run and parse percent
	proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=1, text=True, universal_newlines=True)
	percent = 0
	pattern = re.compile(r"(\d+)%")
	for line in proc.stdout:
		if progress_callback:
			m = pattern.search(line)
			if m:
				try:
					percent = int(m.group(1))
				except Exception:
					pass
				progress_callback(min(percent, 100), line.strip())
	proc.wait()
	if proc.returncode != 0:
		raise RuntimeError(f'7z 返回错误码 {proc.returncode}')


def open_in_explorer(path):
	try:
		subprocess.Popen(['explorer', str(path)])
	except Exception:
		pass


def notify_completion(folder):
	# try win10toast_click if available
	try:
		from win10toast_click import ToastNotifier
		toaster = ToastNotifier()
		def _on_click():
			open_in_explorer(folder)
		toaster.show_toast('解压完成', f'已解压到: {folder}', duration=10, threaded=True, callback_on_click=_on_click)
		return
	except Exception:
		pass

	# fallback: simple tkinter info dialog with Open button
	if tk:
		root = tk.Tk()
		root.withdraw()
		if messagebox.askyesno('解压完成', f'已解压到:\n{folder}\n\n是否打开目录?'):
			open_in_explorer(folder)
		root.destroy()
	else:
		print('已解压到:', folder)


class ProgressWindow:
	def __init__(self, title='解压中'):
		if not tk:
			self._dummy = True
			return
		self._dummy = False
		self.root = tk.Tk()
		self.root.title(title)
		self.root.geometry('420x120')
		self.label = ttk.Label(self.root, text='准备解压...')
		self.label.pack(pady=8)
		self.pb = ttk.Progressbar(self.root, length=380)
		self.pb.pack(padx=10)
		self.log = tk.Text(self.root, height=3, width=50, state='disabled')
		self.log.pack(padx=8, pady=8)

	def start(self):
		if self._dummy:
			return
		threading.Thread(target=self.root.mainloop, daemon=True).start()

	def update(self, percent, line=None):
		if self._dummy:
			print(percent, line)
			return
		def _upd():
			self.pb['value'] = percent
			self.label.config(text=f'进度: {percent}%')
			if line:
				self.log['state'] = 'normal'
				self.log.insert('end', line + '\n')
				self.log.see('end')
				self.log['state'] = 'disabled'
		self.root.after(0, _upd)

	def close(self):
		if self._dummy:
			return
		try:
			self.root.quit()
			self.root.destroy()
		except Exception:
			pass


def process_archive(archive_path, cfg):
	archive_path = Path(archive_path)
	if not archive_path.exists():
		raise FileNotFoundError(archive_path)

	base_path = cfg.get('base_path')
	if not base_path:
		raise RuntimeError('未配置保存路径，请在程序中设置 base_path')

	name = archive_path.stem
	dest = Path(base_path) / name
	if dest.exists() and any(dest.iterdir()):
		if tk:
			root = tk.Tk(); root.withdraw()
			res = messagebox.askyesno('目标已存在', f'目标文件夹 {dest} 已存在且包含文件。是否替换?')
			root.destroy()
			if not res:
				return 'cancelled'
			shutil.rmtree(dest)
		else:
			# no gui: cancel
			return 'cancelled'

	dest.mkdir(parents=True, exist_ok=True)

	pw = ProgressWindow(title=f'正在解压 {archive_path.name}')
	pw.start()

	def progress_cb(pct, line):
		pw.update(pct, line)

	exception = None
	try:
		extract_with_7z(archive_path, dest, progress_callback=progress_cb)
	except Exception as e:
		exception = e
	finally:
		pw.update(100, '完成')
		pw.close()

	if exception:
		raise exception

	notify_completion(str(dest))
	return str(dest)


def register_associations():
	# per-user registration via HKCU\Software\Classes
	try:
		import winreg
		script = str(Path(__file__).absolute())
		cmd = f'"{sys.executable}" "{script}" "%1"'
		exts = ['.zip', '.7z', '.rar']
		for ext in exts:
			with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f'Software\\Classes\\{ext}') as k:
				winreg.SetValue(k, '', winreg.REG_SZ, 'zip-installer.archive')
		with winreg.CreateKey(winreg.HKEY_CURRENT_USER, 'Software\\Classes\\zip-installer.archive') as k:
			winreg.SetValue(k, '', winreg.REG_SZ, 'Zip Installer Archive')
		with winreg.CreateKey(winreg.HKEY_CURRENT_USER, 'Software\\Classes\\zip-installer.archive\\shell\\open\\command') as k:
			winreg.SetValue(k, '', winreg.REG_SZ, cmd)
		return True
	except Exception as e:
		return False


def settings_gui():
	cfg = load_config()
	if not tk:
		print('需要 GUI 支持来配置路径')
		return
	root = tk.Tk()
	root.title('Zip Installer 设置')
	root.geometry('480x160')

	frm = ttk.Frame(root, padding=12)
	frm.pack(fill='both', expand=True)

	ttk.Label(frm, text='解压保存路径:').pack(anchor='w')
	path_var = tk.StringVar(value=cfg.get('base_path', ''))
	entry = ttk.Entry(frm, textvariable=path_var, width=60)
	entry.pack(fill='x', pady=6)

	def browse():
		p = filedialog.askdirectory()
		if p:
			path_var.set(p)

	ttk.Button(frm, text='浏览...', command=browse).pack(anchor='e')

	def save_and_close():
		cfg['base_path'] = path_var.get()
		save_config(cfg)
		messagebox.showinfo('已保存', '配置已保存。')
		root.destroy()

	def do_register():
		ok = register_associations()
		if ok:
			messagebox.showinfo('完成', '已为当前用户注册 .zip .7z .rar 的打开方式。')
		else:
			messagebox.showwarning('失败', '注册失败，请以管理员身份运行以写入系统注册表。')

	btns = ttk.Frame(frm)
	btns.pack(fill='x', pady=8)
	ttk.Button(btns, text='保存', command=save_and_close).pack(side='right')
	ttk.Button(btns, text='注册为默认打开程序', command=do_register).pack(side='left')

	try:
		root.mainloop()
	except KeyboardInterrupt:
		try:
			root.quit()
			root.destroy()
		except Exception:
			pass


def main():
	cfg = load_config()
	if len(sys.argv) > 1:
		# opened with a file
		archive = sys.argv[1]
		try:
			dest = process_archive(archive, cfg)
			if dest == 'cancelled':
				print('用户取消')
		except Exception as e:
			if tk:
				root = tk.Tk(); root.withdraw(); messagebox.showerror('错误', str(e)); root.destroy()
			else:
				print('错误:', e)
		return

	# no args: open settings gui
	try:
		settings_gui()
	except KeyboardInterrupt:
		# user pressed Ctrl+C in terminal; exit gracefully
		try:
			print('\n已中断，退出。')
		except Exception:
			pass


if __name__ == '__main__':
	main()

