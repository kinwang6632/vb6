rem ***** 91.11.22 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe v30/v30@mis @C:\Transform\EMC_Transform\EMC_Check\AddCodeTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad

rem 1. Load Table 與 Column 對照表: Code_Map1
sqlldr userid=v30/v30@mis control=Code_Map1.ctl  log=Code_Map1.log














