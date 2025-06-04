@echo off

@REM start C:\Rpa_CX_Bots_Proxmox\prueba.py
@REM timeout /t 5
@REM start python C:\Rpa_CX_Bots_Proxmox\flask_izzi.py
@REM start python C:\Rpa_CX_Bots_Proxmox\main3.py



@echo off
@REM Obtiene la ubicación del script de lote actual
set "script_dir=%~dp0"

REM Cambia al directorio del script
cd "%script_dir%"
@REM echo "%script_dir%"
start .\prueba.py


@REM cd C:\Rpa_CX_Bots_Proxmox
@REM start .\prueba.py
timeout /t 5
start python .\flask_izzi.py
start python .\main3.py