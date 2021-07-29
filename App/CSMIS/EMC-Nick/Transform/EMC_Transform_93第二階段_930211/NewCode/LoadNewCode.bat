rem ***** 93.02.02 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe emcct/emcct@mis @C:\Transform\NewCode\TruncateCodeTable.sql


rem ***** 建立新代碼 Table 資料 *****

rem del *.log
rem del *.bad

rem 1. Load CD004之新代碼資料: CD004_NEW
sqlldr userid=emcct/emcct@mis control=CD004_NEW.ctl  log=CD004_NEW.log

rem 2. Load CD005之新代碼資料: CD005_NEW
sqlldr userid=emcct/emcct@mis control=CD005_NEW.ctl  log=CD005_NEW.log

rem 3. Load CD006之新代碼資料: CD006_NEW
sqlldr userid=emcct/emcct@mis control=CD006_NEW.ctl  log=CD006_NEW.log

rem 4. Load CD007之新代碼資料: CD007_NEW
sqlldr userid=emcct/emcct@mis control=CD007_NEW.ctl  log=CD007_NEW.log

rem 5. Load CD008之新代碼資料: CD008_NEW
sqlldr userid=emcct/emcct@mis control=CD008_NEW.ctl  log=CD008_NEW.log

rem 6. Load CD008A之新代碼資料: CD008A_NEW
sqlldr userid=emcct/emcct@mis control=CD008A_NEW.ctl  log=CD008A_NEW.log

rem 7. Load CD008C之新代碼資料: CD008C_NEW
sqlldr userid=emcct/emcct@mis control=CD008C_NEW.ctl  log=CD008C_NEW.log

rem 8. Load CD009之新代碼資料: CD009_NEW
sqlldr userid=emcct/emcct@mis control=CD009_NEW.ctl  log=CD009_NEW.log

rem 9. Load CD011之新代碼資料: CD011_NEW
sqlldr userid=emcct/emcct@mis control=CD011_NEW.ctl  log=CD011_NEW.log

rem 10. Load CD011A之新代碼資料: CD011A_NEW
sqlldr userid=emcct/emcct@mis control=CD011A_NEW.ctl  log=CD011A_NEW.log

rem 11. Load CD011B之新代碼資料: CD011B_NEW
sqlldr userid=emcct/emcct@mis control=CD011B_NEW.ctl  log=CD011B_NEW.log

rem 12. Load CD012之新代碼資料: CD012_NEW
sqlldr userid=emcct/emcct@mis control=CD012_NEW.ctl  log=CD012_NEW.log

rem 13. Load CD013之新代碼資料: CD013_NEW
sqlldr userid=emcct/emcct@mis control=CD013_NEW.ctl  log=CD013_NEW.log

rem 14. Load CD014之新代碼資料: CD014_NEW
sqlldr userid=emcct/emcct@mis control=CD014_NEW.ctl  log=CD014_NEW.log

rem 15. Load CD015之新代碼資料: CD015_NEW
sqlldr userid=emcct/emcct@mis control=CD015_NEW.ctl  log=CD015_NEW.log

rem 16. Load CD016之新代碼資料: CD016_NEW
sqlldr userid=emcct/emcct@mis control=CD016_NEW.ctl  log=CD016_NEW.log

rem 17. Load CD018之新代碼資料: CD018_NEW
sqlldr userid=emcct/emcct@mis control=CD018_NEW.ctl  log=CD018_NEW.log

rem 18. Load CD019之新代碼資料: CD019_NEW
sqlldr userid=emcct/emcct@mis control=CD019_NEW.ctl  log=CD019_NEW.log

rem 19. Load CD019A之新代碼資料: CD019A_NEW
sqlldr userid=emcct/emcct@mis control=CD019A_NEW.ctl  log=CD019A_NEW.log

rem 20. Load CD019B之新代碼資料: CD019B_NEW
sqlldr userid=emcct/emcct@mis control=CD019B_NEW.ctl  log=CD019B_NEW.log

rem 21. Load CD019CD004之新代碼資料: CD019CD004_NEW
sqlldr userid=emcct/emcct@mis control=CD019CD004_NEW.ctl  log=CD019CD004_NEW.log

rem 22. Load CD019SO017之新代碼資料: CD019SO017_NEW
sqlldr userid=emcct/emcct@mis control=CD019SO017_NEW.ctl  log=CD019SO017_NEW.log

rem 23. Load CD020之新代碼資料: CD020_NEW
sqlldr userid=emcct/emcct@mis control=CD020_NEW.ctl  log=CD020_NEW.log

rem 24. Load CD021之新代碼資料: CD021_NEW
sqlldr userid=emcct/emcct@mis control=CD021_NEW.ctl  log=CD021_NEW.log

rem 25. Load CD022之新代碼資料: CD022_NEW
sqlldr userid=emcct/emcct@mis control=CD022_NEW.ctl  log=CD022_NEW.log

rem 26. Load CD022A之新代碼資料: CD022A_NEW
sqlldr userid=emcct/emcct@mis control=CD022A_NEW.ctl  log=CD022A_NEW.log

rem 27. Load CD022CD019之新代碼資料: CD022CD019_NEW
sqlldr userid=emcct/emcct@mis control=CD022CD019_NEW.ctl  log=CD022CD019_NEW.log

rem 28. Load CD025之新代碼資料: CD025_NEW
sqlldr userid=emcct/emcct@mis control=CD025_NEW.ctl  log=CD025_NEW.log

rem 29. Load CD026之新代碼資料: CD026_NEW
sqlldr userid=emcct/emcct@mis control=CD026_NEW.ctl  log=CD026_NEW.log

rem 30. Load CD027A之新代碼資料: CD027A_NEW
sqlldr userid=emcct/emcct@mis control=CD027A_NEW.ctl  log=CD027A_NEW.log

rem 31. Load CD027B之新代碼資料: CD027B_NEW
sqlldr userid=emcct/emcct@mis control=CD027B_NEW.ctl  log=CD027B_NEW.log

rem 32. Load CD027C之新代碼資料: CD027C_NEW
sqlldr userid=emcct/emcct@mis control=CD027C_NEW.ctl  log=CD027C_NEW.log

rem 33. Load CD027D之新代碼資料: CD027D_NEW
sqlldr userid=emcct/emcct@mis control=CD027D_NEW.ctl  log=CD027D_NEW.log

rem 34. Load CD041之新代碼資料: CD041_NEW
sqlldr userid=emcct/emcct@mis control=CD041_NEW.ctl  log=CD041_NEW.log

rem 35. Load CD049之新代碼資料: CD049_NEW
sqlldr userid=emcct/emcct@mis control=CD049_NEW.ctl  log=CD049_NEW.log

rem 36. Load CD051之新代碼資料: CD051_NEW
sqlldr userid=emcct/emcct@mis control=CD051_NEW.ctl  log=CD051_NEW.log

rem 37. Load CD055之新代碼資料: CD055_NEW
sqlldr userid=emcct/emcct@mis control=CD055_NEW.ctl  log=CD055_NEW.log

rem 38. Load CD059之新代碼資料: CD059_NEW
sqlldr userid=emcct/emcct@mis control=CD059_NEW.ctl  log=CD059_NEW.log
