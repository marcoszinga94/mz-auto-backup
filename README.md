# Backup Pro

> Aplicación de escritorio para crear backups locales en ZIP y subirlos a un servidor por SSH/SFTP.

![Python](https://img.shields.io/badge/Python-3-3776AB?logo=python&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/UI-CustomTkinter-1F6AA5)
![Paramiko](https://img.shields.io/badge/SSH-Paramiko-2E8B57)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows&logoColor=white)
![PyInstaller](https://img.shields.io/badge/build-PyInstaller-FFD43B)

## Características

- **Backup local comprimido**: empaqueta la carpeta de origen en un `.zip` con marca de tiempo.
- **Ejecución manual o diaria**: al abrir la app comprueba si ya se hizo el backup del día y lo lanza solo.
- **Copia remota opcional**: sube el ZIP por SFTP si se configuran host, usuario, contraseña y ruta remota.
- **Retención automática**: elimina los `backup_*.zip` de más de 90 días en el destino.
- **Comprobación de espacio**: verifica que el destino tenga al menos un 20 % más de espacio libre que el origen.
- **Notificaciones nativas** del sistema y registro histórico en `backup_log.txt`.
- **Ajustes persistentes** en la carpeta de datos del usuario: origen, destino, tema y datos remotos.
- **Tema claro/oscuro** e icono de ventana según el tema.
- **Acceso directo al inicio** de Windows creado automáticamente.

## Stack

| Capa | Tecnología |
| --- | --- |
| Lenguaje | Python 3 |
| Interfaz | CustomTkinter (Tkinter) |
| Compresión | `shutil.make_archive` |
| Transferencia remota | Paramiko (SSH/SFTP) |
| Integración Windows | `win32com`, `APPDATA`, accesos directos |
| Notificaciones | Plyer |
| Empaquetado | PyInstaller |

## Estructura del proyecto

```text
.
├── back_up.py            # aplicación completa: UI, backups y SFTP
├── BackupPro.spec        # configuración de PyInstaller
├── back_up_dark.ico      # icono para tema oscuro
└── back_up_light.ico     # icono para tema claro
```

## Puesta en marcha

### Requisitos

- Windows 10 u 11 (usa `APPDATA`, `win32com` e iconos `.ico`).
- Python 3 con Tkinter.
- Acceso de escritura a la carpeta de destino.
- Servidor SSH con SFTP solo si se quiere la copia remota.

### Instalación

No hay archivo de dependencias en el repositorio. Instalarlas manualmente:

```bash
pip install customtkinter paramiko pywin32 plyer
python back_up.py
```

Para generar un ejecutable sin consola:

```bash
pyinstaller BackupPro.spec
```

### Variables de entorno

No requiere variables de entorno. La configuración se guarda en `%APPDATA%\Backup Pro\config.json`.

## Scripts

| Comando | Descripción |
| --- | --- |
| `python back_up.py` | Abre la aplicación. |
| `pyinstaller BackupPro.spec` | Compila el ejecutable de Windows. |

## Despliegue

La distribución prevista es un ejecutable de Windows generado con PyInstaller. El repositorio no incluye instalador ni pipeline de publicación; la configuración y los logs quedan en la carpeta de datos del usuario, fuera del ejecutable.

## Estado del proyecto

Funcional, con toda la lógica en un único archivo. No hay tests ni empaquetado automatizado.

## Licencia

Sin licencia definida.

## Autor

Marcos Zingaretti — https://github.com/marcoszinga94
