rem ***** 91.11.22 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe emcct/emcct@mis @C:\Transform\Check\AddCodeTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad

rem 1. Load Table 與 Column 對照表: Code_Check1
sqlldr userid=emcct/emcct@mis control=Code_Check1.ctl  log=Code_Check1.log














