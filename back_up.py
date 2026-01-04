import customtkinter as ctk
from tkinter import filedialog
import shutil
import json
import os
import threading
import sys
import paramiko 
from datetime import datetime, timedelta
from pathlib import Path
import win32com.client 
from plyer import notification # Librería para notificaciones

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS # type: ignore
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class BackupPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- GESTIÓN DE RUTAS ---
        raw_appdata = os.getenv('APPDATA')
        base_path = Path(raw_appdata) if raw_appdata else Path.home() / "AppData" / "Roaming"
        self.appdata_dir = base_path / "Backup Pro"
        self.appdata_dir.mkdir(parents=True, exist_ok=True)
        self.config_file = str(self.appdata_dir / "config.json")
        self.log_file = self.appdata_dir / "backup_log.txt"

        self.title("Backup Pro")
        self.geometry("550x500")

        # --- VARIABLES ---
        self.source_path = ctk.StringVar(value="No seleccionado")
        self.dest_path = ctk.StringVar(value=str(Path.home() / "Desktop"))
        self.is_dark_mode = ctk.BooleanVar(value=True)
        self.last_backup_date = ""
        self.remote_host, self.remote_user = ctk.StringVar(), ctk.StringVar()
        self.remote_pass, self.remote_path = ctk.StringVar(), ctk.StringVar(value="/home/")

        self.load_settings()
        ctk.set_appearance_mode("dark" if self.is_dark_mode.get() else "light")
        self.init_ui()
        self.update_icon()

        self.after(2000, self.check_and_run_daily_backup)

    def notify(self, title, message):
        """ Envía una notificación nativa a Windows """
        try:
            notification.notify(
                title=title,
                message=message,
                app_name="BackupPro",
                timeout=10,
                icon=resource_path("back_up_dark.ico")
            ) # type: ignore
        except: pass

    def write_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {message}\n")
        except: pass

    def clean_old_backups(self):
        dst = Path(self.dest_path.get())
        if not dst.exists(): return
        
        limit_date = datetime.now() - timedelta(days=90)
        self.write_log("Iniciando limpieza de archivos viejos (Retención 90 días)...")
        
        count = 0
        for file in dst.glob("backup_*.zip"):
            file_time = datetime.fromtimestamp(file.stat().st_mtime)
            if file_time < limit_date:
                try:
                    file.unlink()
                    count += 1
                except: pass
        if count > 0: self.write_log(f"Limpieza completada: {count} archivos eliminados.")

    def check_disk_space(self, source):
        total, used, free = shutil.disk_usage(self.dest_path.get())
        # Estimación simple: el tamaño de la carpeta origen
        source_size = sum(f.stat().st_size for f in Path(source).rglob('*') if f.is_file())
        
        # Dejamos un margen del 20% extra por seguridad
        return free > (source_size * 1.2)

    def init_ui(self):
        self.label_title = ctk.CTkLabel(self, text="Backup Pro", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_title.pack(pady=15)

        self.theme_switch = ctk.CTkSwitch(self, text="Modo Oscuro", variable=self.is_dark_mode, command=self.toggle_theme_event)
        self.theme_switch.pack(pady=5, padx=25, anchor="ne")

        self.btn_logs = ctk.CTkButton(self, text="Ver historial de Logs", fg_color="transparent", border_width=1, 
                                      height=24, command=lambda: os.startfile(self.appdata_dir))
        self.btn_logs.pack(pady=5)

        self.create_path_row("📁 Origen:", self.source_path, self.select_source)
        self.create_path_row("📍 Destino:", self.dest_path, self.select_dest)
        self.create_input_row("IP Host:", self.remote_host)
        self.create_input_row("Usuario:", self.remote_user)
        self.create_input_row("Password:", self.remote_pass, show="*")
        self.create_input_row("Ruta Remota:", self.remote_path)

        self.label_status = ctk.CTkLabel(self, text="Listo", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_status.pack(pady=10)

        self.btn_run = ctk.CTkButton(self, text="EJECUTAR AHORA", height=45, command=lambda: self.manual_backup(False))
        self.btn_run.pack(pady=10, padx=40, fill="x")

        self.manage_startup_shortcut()

    def run_backup_logic(self, is_auto=False):
        src, dst = self.source_path.get(), self.dest_path.get()
        if not os.path.exists(src): return

        # 1. Validar espacio
        if not self.check_disk_space(src):
            msg = "Error: Espacio insuficiente en disco de destino."
            self.write_log(msg)
            self.notify("Fallo de Backup", msg)
            self.label_status.configure(text="❌ Sin espacio", text_color="#e74c3c")
            return

        try:
            self.label_status.configure(text="⏳ Comprimiendo...", text_color="#f1c40f")
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
            out_base = str(Path(dst) / f"backup_{ts}")
            
            final_zip = shutil.make_archive(out_base, 'zip', src)
            
            # 2. Subida Remota
            sftp_ok = self.upload_to_sftp(final_zip)
            
            # 3. Notificación Final
            if sftp_ok:
                self.notify("Backup Exitoso", f"Copia local y SFTP guardadas correctamente.")
            else:
                self.notify("Backup Parcial", "Se guardó la copia local pero falló el servidor remoto.")

            # 4. Limpieza de archivos de > 90 días
            self.clean_old_backups()
            
            self.last_backup_date = datetime.now().strftime("%Y-%m-%d")
            self.save_settings()
            if is_auto: self.after(5000, self.destroy)
        except Exception as e:
            self.write_log(f"ERROR: {e}")
            self.notify("Error de Backup", "Ocurrió un error inesperado. Revisa los logs.")

    # ... (Resto de métodos: upload_to_sftp, load_settings, save_settings, etc. iguales al anterior) ...

    def upload_to_sftp(self, local_zip):
        transport = sftp = None
        try:
            h, u, p = self.remote_host.get(), self.remote_user.get(), self.remote_pass.get()
            rp = self.remote_path.get()
            if not h or not u: return False
            transport = paramiko.Transport((h, 22))
            transport.connect(username=u, password=p)
            sftp = paramiko.SFTPClient.from_transport(transport)
            remote_dir = rp if rp.endswith('/') else rp + '/'
            rem_file = remote_dir + os.path.basename(local_zip)
            sftp.put(local_zip, rem_file) # type: ignore
            return True
        except: return False
        finally:
            if sftp: sftp.close()
            if transport: transport.close()

    def update_icon(self):
        try:
            icon_file = "back_up.ico" if self.is_dark_mode.get() else "back_up_light.ico"
            icon_res = resource_path(icon_file)
            if os.path.exists(icon_res):
                self.after(200, lambda: self.wm_iconbitmap(icon_res))
        except: pass

    def toggle_theme_event(self):
        ctk.set_appearance_mode("dark" if self.is_dark_mode.get() else "light")
        self.update_icon()
        self.save_settings()

    def check_and_run_daily_backup(self):
        today = datetime.now().strftime("%Y-%m-%d")
        if self.source_path.get() != "No seleccionado" and self.last_backup_date != today:
            self.manual_backup(is_auto=True)

    def manual_backup(self, is_auto=False):
        threading.Thread(target=self.run_backup_logic, args=(is_auto,), daemon=True).start()

    def create_path_row(self, label, var, cmd):
        f = ctk.CTkFrame(self); f.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(f, text=label, width=100, anchor="w").pack(side="left", padx=10)
        ctk.CTkLabel(f, textvariable=var, text_color="gray", wraplength=300).pack(side="left", padx=5, expand=True)
        ctk.CTkButton(f, text="Cambiar", width=70, command=cmd).pack(side="right", padx=0)

    def create_input_row(self, label, var, show=None):
        f = ctk.CTkFrame(self); f.pack(pady=3, padx=20, fill="x")
        ctk.CTkLabel(f, text=label, width=100, anchor="w").pack(side="left", padx=10)
        e = ctk.CTkEntry(f, textvariable=var, show=show)
        e.pack(side="right", padx=0, fill="x", expand=True)
        e.bind("<FocusOut>", lambda ev: self.save_settings())

    def manage_startup_shortcut(self):
        try:
            appdata_str = os.getenv('APPDATA')
            base = Path(appdata_str) if appdata_str else Path.home() / "AppData" / "Roaming"
            startup_folder = base / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
            shortcut_path = str(startup_folder / "BackupPro.lnk")
            target = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.save()
        except: pass

    def load_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    d = json.load(f)
                    self.source_path.set(d.get("source", "No seleccionado"))
                    self.dest_path.set(d.get("dest", str(Path.home() / "Desktop")))
                    self.is_dark_mode.set(d.get("theme", "dark") == "dark")
                    self.last_backup_date = d.get("last_backup", "")
                    self.remote_host.set(d.get("r_host", ""))
                    self.remote_user.set(d.get("r_user", ""))
                    self.remote_pass.set(d.get("r_pass", ""))
                    self.remote_path.set(d.get("r_path", "/home/"))
            except: pass

    def save_settings(self):
        with open(self.config_file, "w") as f:
            json.dump({
                "source": self.source_path.get(), "dest": self.dest_path.get(),
                "theme": "dark" if self.is_dark_mode.get() else "light",
                "last_backup": self.last_backup_date,
                "r_host": self.remote_host.get(), "r_user": self.remote_user.get(),
                "r_pass": self.remote_pass.get(), "r_path": self.remote_path.get()
            }, f)

    def select_source(self):
        p = filedialog.askdirectory()
        if p: (self.source_path.set(p), self.save_settings()) # type: ignore

    def select_dest(self):
        p = filedialog.askdirectory()
        if p: (self.dest_path.set(p), self.save_settings()) # type: ignore

if __name__ == "__main__":
    app = BackupPro()
    app.mainloop()