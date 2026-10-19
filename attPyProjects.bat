@echo off

:: Caminho do diretório onde os arquivos devem ser atualizados
set DESTINO=C:\PyProjects\DocsGenOS

:: Caminho temporário para o clone (nome distinto para não sobrescrever a variável TEMP do sistema)
set TEMP_REPO=%TEMP%\repo_temp

:: URL do repositório remoto
set REPO=https://github.com/Samir-SAPSD/DocsGenOS.git

:: Apaga pasta temporária se já existir
rmdir /s /q "%TEMP_REPO%"

:: Clona o repositório na pasta temporária
git clone %REPO% "%TEMP_REPO%"

:: Copia os arquivos do repositório clonado para o destino, sobrescrevendo tudo
xcopy "%TEMP_REPO%\*" "%DESTINO%\" /E /H /Y /C

:: Limpa a pasta temporária
rmdir /s /q "%TEMP_REPO%"

echo Atualização concluída com sucesso!

:: Executa a configuração do ambiente para renovar os atalhos
call "%DESTINO%\configurar_ambiente.bat"
