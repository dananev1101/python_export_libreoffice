# Эскпорт диапазона в png

Для работы необходимо установить LibreOffice и запустить его в режиме сервера


"C:\Program Files\LibreOffice\program\soffice.exe" ^
  --headless ^
  --nologo ^
  --accept="socket,host=localhost,port=2099;urp;" ^
  -env:UserInstallation=file:///C:/lo_temp_24.8 ^
  --writer ^
  --norestore