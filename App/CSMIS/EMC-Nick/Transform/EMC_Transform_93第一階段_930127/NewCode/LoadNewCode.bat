rem ***** 91.11.22 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe emctc/emctc@ty @C:\Transform\NewCode\TruncateCodeTable.sql


rem ***** 建立新代碼 Table 資料 *****

rem del *.log
rem del *.bad

rem 1. Load CD009之新代碼資料: CD009_NEW
sqlldr userid=emctc/emctc@ty control=CD009_NEW.ctl  log=CD009_NEW.log

rem 2. Load CD012之新代碼資料: CD012_NEW
sqlldr userid=emctc/emctc@ty control=CD012_NEW.ctl  log=CD012_NEW.log

rem 3. Load CD013之新代碼資料: CD013_NEW
sqlldr userid=emctc/emctc@ty control=CD013_NEW.ctl  log=CD013_NEW.log

rem 4. Load CD015之新代碼資料: CD015_NEW
sqlldr userid=emctc/emctc@ty control=CD015_NEW.ctl  log=CD015_NEW.log

rem 5. Load CD016之新代碼資料: CD016_NEW
sqlldr userid=emctc/emctc@ty control=CD016_NEW.ctl  log=CD016_NEW.log

rem 6. Load CD020之新代碼資料: CD020_NEW
sqlldr userid=emctc/emctc@ty control=CD020_NEW.ctl  log=CD020_NEW.log

rem 7. Load CD021之新代碼資料: CD021_NEW
sqlldr userid=emctc/emctc@ty control=CD021_NEW.ctl  log=CD021_NEW.log

rem 8. Load CD026之新代碼資料: CD026_NEW
sqlldr userid=emctc/emctc@ty control=CD026_NEW.ctl  log=CD026_NEW.log

rem 9. Load CD051之新代碼資料: CD051_NEW
sqlldr userid=emctc/emctc@ty control=CD051_NEW.ctl  log=CD051_NEW.log

