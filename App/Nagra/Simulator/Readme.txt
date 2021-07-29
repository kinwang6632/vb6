1. SMS Command Gateway 用到的 Oracle Table

   (*)  com.cacmdresponselog
        com.casendcmdlog
        com.send_nagra
        com.recv_nagra

   請先確定這些 Schema 已建立, 可 Copy EMC COM 區 的 Schema.


2. 若要跑 CA 模擬軟體, 請先安裝 Perl 的 library. 提供的壓縮檔內有
   安裝檔  --> ActivePerl-5.6.0.623-MSWin32-x86-multi-thread.msi
   
   Perl library 安裝完成後, 把 CA 模擬軟體 解壓縮到另一個目錄(自訂都可以)
   Ex: C:\CA

3. SMS Command Gateway 共須要 4 個檔案

   (*) AppInfo.ini
   (*) midas.dll
   (*) Ngw.exe 
   (*) 模擬器.cfg

   解壓縮後放同一個目錄 Ex: C:\SMSCmd


4. 執行順序

   先執行 CA 模擬軟體 目錄下的 sms.bat  --> 會出現一 DOS 視窗
   再來執行 SMS Command Gateway  --> ngw.exe


5. SMS Command Gateway 設定請參考 SMS Command Gateway 操作手冊.doc;