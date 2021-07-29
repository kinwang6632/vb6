Sqlplusw.exe emcct/emcct@mis @C:\Transform\NewCode\NewCheck\AddNewTempTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad

rem 0. Load  : Code_Check3
sqlldr userid=emcct/emcct@mis control=Code_Check3.ctl  log=Code_Check3.log

rem 1. Load CD004之新代碼資料: CD004_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD004_NEWTEMP.ctl  log=CD004_NEWTEMP.log

rem 2. Load CD005之新代碼資料: CD005_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD005_NEWTEMP.ctl  log=CD005_NEWTEMP.log

rem 3. Load CD006之新代碼資料: CD006_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD006_NEWTEMP.ctl  log=CD006_NEWTEMP.log

rem 4. Load CD007之新代碼資料: CD007_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD007_NEWTEMP.ctl  log=CD007_NEWTEMP.log

rem 5. Load CD008之新代碼資料: CD008_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD008_NEWTEMP.ctl  log=CD008_NEWTEMP.log

rem 6. Load CD008A之新代碼資料: CD008A_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD008A_NEWTEMP.ctl  log=CD008A_NEWTEMP.log

rem 7. Load CD009之新代碼資料: CD009_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD009_NEWTEMP.ctl  log=CD009_NEWTEMP.log

rem 8. Load CD011之新代碼資料: CD011_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD011_NEWTEMP.ctl  log=CD011_NEWTEMP.log

rem 9. Load CD011B之新代碼資料: CD011B_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD011B_NEWTEMP.ctl  log=CD011B_NEWTEMP.log

rem 10. Load CD012之新代碼資料: CD012_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD012_NEWTEMP.ctl  log=CD012_NEWTEMP.log

rem 11. Load CD013之新代碼資料: CD013_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD013_NEWTEMP.ctl  log=CD013_NEWTEMP.log

rem 12. Load CD014之新代碼資料: CD014_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD014_NEWTEMP.ctl  log=CD014_NEWTEMP.log

rem 13. Load CD015之新代碼資料: CD015_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD015_NEWTEMP.ctl  log=CD015_NEWTEMP.log

rem 14. Load CD016之新代碼資料: CD016_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD016_NEWTEMP.ctl  log=CD016_NEWTEMP.log

rem 15. Load CD018之新代碼資料: CD018_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD018_NEWTEMP.ctl  log=CD018_NEWTEMP.log

rem 16. Load CD019之新代碼資料: CD019_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD019_NEWTEMP.ctl  log=CD019_NEWTEMP.log

rem 17. Load CD020之新代碼資料: CD020_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD020_NEWTEMP.ctl  log=CD020_NEWTEMP.log

rem 18. Load CD021之新代碼資料: CD021_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD021_NEWTEMP.ctl  log=CD021_NEWTEMP.log

rem 19. Load CD022之新代碼資料: CD022_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD022_NEWTEMP.ctl  log=CD022_NEWTEMP.log

rem 20. Load CD025之新代碼資料: CD025_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD025_NEWTEMP.ctl  log=CD025_NEWTEMP.log

rem 21. Load CD026之新代碼資料: CD026_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD026_NEWTEMP.ctl  log=CD026_NEWTEMP.log

rem 22. Load CD049之新代碼資料: CD049_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD049_NEWTEMP.ctl  log=CD049_NEWTEMP.log

rem 23. Load CD051之新代碼資料: CD051_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD051_NEWTEMP.ctl  log=CD051_NEWTEMP.log

rem 24. Load CD055之新代碼資料: CD055_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD055_NEWTEMP.ctl  log=CD055_NEWTEMP.log

rem 25. Load CD059之新代碼資料: CD059_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD059_NEWTEMP.ctl  log=CD059_NEWTEMP.log