rem ***** 93.02.17 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe emcct/emcct@mis @C:\Transform\Map\AddTmpTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad


rem 1. Load CD004之新舊對照表: CD004_MAP
sqlldr userid=emcct/emcct@mis control=CD004_MAP.ctl  log=CD004_MAP.log


rem 2. Load CD005之新舊對照表: CD005_MAP
sqlldr userid=emcct/emcct@mis control=CD005_MAP.ctl  log=CD005_MAP.log


rem 3. Load CD006之新舊對照表: CD006_MAP
sqlldr userid=emcct/emcct@mis control=CD006_MAP.ctl  log=CD006_MAP.log


rem 4. Load CD007之新舊對照表: CD007_MAP
sqlldr userid=emcct/emcct@mis control=CD007_MAP.ctl  log=CD007_MAP.log


rem 5. Load CD008之新舊對照表: CD008_MAP
sqlldr userid=emcct/emcct@mis control=CD008_MAP.ctl  log=CD008_MAP.log


rem 6. Load CD008A之新舊對照表: CD008A_MAP
sqlldr userid=emcct/emcct@mis control=CD008A_MAP.ctl  log=CD008A_MAP.log


rem 7. Load CD009之新舊對照表: CD009_MAP
sqlldr userid=emcct/emcct@mis control=CD009_MAP.ctl  log=CD009_MAP.log


rem 8. Load CD011之新舊對照表: CD011_MAP
sqlldr userid=emcct/emcct@mis control=CD011_MAP.ctl  log=CD011_MAP.log


rem 9. Load CD011B之新舊對照表: CD011B_MAP
sqlldr userid=emcct/emcct@mis control=CD011B_MAP.ctl  log=CD011B_MAP.log


rem 10. Load CD012之新舊對照表: CD012_MAP
sqlldr userid=emcct/emcct@mis control=CD012_MAP.ctl  log=CD012_MAP.log


rem 11. Load CD013之新舊對照表: CD013_MAP
sqlldr userid=emcct/emcct@mis control=CD013_MAP.ctl  log=CD013_MAP.log


rem 12. Load CD014之新舊對照表: CD014_MAP
sqlldr userid=emcct/emcct@mis control=CD014_MAP.ctl  log=CD014_MAP.log


rem 13. Load CD015之新舊對照表: CD015_MAP
sqlldr userid=emcct/emcct@mis control=CD015_MAP.ctl  log=CD015_MAP.log


rem 14. Load CD016之新舊對照表: CD016_MAP
sqlldr userid=emcct/emcct@mis control=CD016_MAP.ctl  log=CD016_MAP.log


rem 15. Load CD018之新舊對照表: CD018_MAP
sqlldr userid=emcct/emcct@mis control=CD018_MAP.ctl  log=CD018_MAP.log


rem 16. Load CD019之新舊對照表: CD019_MAP
sqlldr userid=emcct/emcct@mis control=CD019_MAP.ctl  log=CD019_MAP.log


rem 17. Load CD020之新舊對照表: CD020_MAP
sqlldr userid=emcct/emcct@mis control=CD020_MAP.ctl  log=CD020_MAP.log


rem 18. Load CD021之新舊對照表: CD021_MAP
sqlldr userid=emcct/emcct@mis control=CD021_MAP.ctl  log=CD021_MAP.log


rem 19. Load CD022之新舊對照表: CD022_MAP
sqlldr userid=emcct/emcct@mis control=CD022_MAP.ctl  log=CD022_MAP.log


rem 20. Load CD025之新舊對照表: CD025_MAP
sqlldr userid=emcct/emcct@mis control=CD025_MAP.ctl  log=CD025_MAP.log


rem 21. Load CD026之新舊對照表: CD026_MAP
sqlldr userid=emcct/emcct@mis control=CD026_MAP.ctl  log=CD026_MAP.log


rem 22. Load CD049之新舊對照表: CD049_MAP
sqlldr userid=emcct/emcct@mis control=CD049_MAP.ctl  log=CD049_MAP.log


rem 23. Load CD051之新舊對照表: CD051_MAP
sqlldr userid=emcct/emcct@mis control=CD051_MAP.ctl  log=CD051_MAP.log


rem 24. Load CD055之新舊對照表: CD055_MAP
sqlldr userid=emcct/emcct@mis control=CD055_MAP.ctl  log=CD055_MAP.log


rem 25. Load CD059之新舊對照表: CD059_MAP
sqlldr userid=emcct/emcct@mis control=CD059_MAP.ctl  log=CD059_MAP.log
