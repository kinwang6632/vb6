Sqlplusw.exe emcct/emcct@mis @C:\Transform\NewCode\NewCheck\AddNewTempTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad

rem 0. Load  : Code_Check3
sqlldr userid=emcct/emcct@mis control=Code_Check3.ctl  log=Code_Check3.log

rem 1. Load 新代碼 : CD009_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD009_NEWTEMP.ctl  log=CD009_NEWTEMP.log

rem 2. Load 新代碼 : CD012_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD012_NEWTEMP.ctl  log=CD012_NEWTEMP.log

rem 3. Load 新代碼 : CD013_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD013_NEWTEMP.ctl  log=CD013_NEWTEMP.log

rem 4. Load 新代碼 : CD015_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD015_NEWTEMP.ctl  log=CD015_NEWTEMP.log

rem 5. Load 新代碼 : CD016_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD016_NEWTEMP.ctl  log=CD016_NEWTEMP.log

rem 6. Load 新代碼 : CD020_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD020_NEWTEMP.ctl  log=CD020_NEWTEMP.log

rem 7. Load 新代碼 : CD021_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD021_NEWTEMP.ctl  log=CD021_NEWTEMP.log

rem 8. Load 新代碼 : CD026_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD026_NEWTEMP.ctl  log=CD026_NEWTEMP.log

rem 9. Load 新代碼 : CD051_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD051_NEWTEMP.ctl  log=CD051_NEWTEMP.log













