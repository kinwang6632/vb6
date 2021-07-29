rem ***** 92.12.27 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe emctc/emctc@ty @C:\Transform\Map\AddTmpTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad


rem 1. Load CD009之新舊對照表: CD009_MAP
sqlldr userid=emctc/emctc@ty control=CD009_MAP.ctl  log=CD009_MAP.log


rem 2. Load CD012之新舊對照表: CD012_MAP
sqlldr userid=emctc/emctc@ty control=CD012_MAP.ctl  log=CD012_MAP.log


rem 3. Load CD013之新舊對照表: CD013_MAP
sqlldr userid=emctc/emctc@ty control=CD013_MAP.ctl  log=CD013_MAP.log


rem 4. Load CD015之新舊對照表: CD015_MAP
sqlldr userid=emctc/emctc@ty control=CD015_MAP.ctl  log=CD015_MAP.log


rem 5. Load CD016之新舊對照表: CD016_MAP
sqlldr userid=emctc/emctc@ty control=CD016_MAP.ctl  log=CD016_MAP.log


rem 6. Load CD020之新舊對照表: CD020_MAP
sqlldr userid=emctc/emctc@ty control=CD020_MAP.ctl  log=CD020_MAP.log


rem 7. Load CD021之新舊對照表: CD021_MAP
sqlldr userid=emctc/emctc@ty control=CD021_MAP.ctl  log=CD021_MAP.log


rem 8. Load CD026之新舊對照表: CD026_MAP
sqlldr userid=emctc/emctc@ty control=CD026_MAP.ctl  log=CD026_MAP.log


rem 9. Load CD051之新舊對照表: CD051_MAP
sqlldr userid=emctc/emctc@ty control=CD051_MAP.ctl  log=CD051_MAP.log


