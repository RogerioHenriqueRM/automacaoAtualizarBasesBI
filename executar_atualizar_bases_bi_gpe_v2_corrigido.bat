@echo off
title Atualizar Bases BI GPE

echo ==========================================
echo      ATUALIZANDO BASES DO BI GPE
echo ==========================================
echo.

"C:\Users\rogerio.matos\AppData\Local\Programs\Python\Python314\python.exe" ^
"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\atualizar_bases_bi_gpe_v2_sem_limite_50mb.py"

echo.
echo ==========================================
echo Processo finalizado.
echo Codigo de retorno: %ERRORLEVEL%
echo ==========================================
echo.

pause