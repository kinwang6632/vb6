Sqlplusw.exe emcct/emcct@mis @C:\Transform\NewCode\NewCheck\AddNewTempTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad

rem 0. Load  : Code_Check3
sqlldr userid=emcct/emcct@mis control=Code_Check3.ctl  log=Code_Check3.log

rem 1. Load CD004���s�N�X���: CD004_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD004_NEWTEMP.ctl  log=CD004_NEWTEMP.log

rem 2. Load CD005���s�N�X���: CD005_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD005_NEWTEMP.ctl  log=CD005_NEWTEMP.log

rem 3. Load CD006���s�N�X���: CD006_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD006_NEWTEMP.ctl  log=CD006_NEWTEMP.log

rem 4. Load CD007���s�N�X���: CD007_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD007_NEWTEMP.ctl  log=CD007_NEWTEMP.log

rem 5. Load CD008���s�N�X���: CD008_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD008_NEWTEMP.ctl  log=CD008_NEWTEMP.log

rem 6. Load CD008A���s�N�X���: CD008A_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD008A_NEWTEMP.ctl  log=CD008A_NEWTEMP.log

rem 7. Load CD009���s�N�X���: CD009_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD009_NEWTEMP.ctl  log=CD009_NEWTEMP.log

rem 8. Load CD011���s�N�X���: CD011_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD011_NEWTEMP.ctl  log=CD011_NEWTEMP.log

rem 9. Load CD011B���s�N�X���: CD011B_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD011B_NEWTEMP.ctl  log=CD011B_NEWTEMP.log

rem 10. Load CD012���s�N�X���: CD012_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD012_NEWTEMP.ctl  log=CD012_NEWTEMP.log

rem 11. Load CD013���s�N�X���: CD013_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD013_NEWTEMP.ctl  log=CD013_NEWTEMP.log

rem 12. Load CD014���s�N�X���: CD014_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD014_NEWTEMP.ctl  log=CD014_NEWTEMP.log

rem 13. Load CD015���s�N�X���: CD015_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD015_NEWTEMP.ctl  log=CD015_NEWTEMP.log

rem 14. Load CD016���s�N�X���: CD016_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD016_NEWTEMP.ctl  log=CD016_NEWTEMP.log

rem 15. Load CD018���s�N�X���: CD018_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD018_NEWTEMP.ctl  log=CD018_NEWTEMP.log

rem 16. Load CD019���s�N�X���: CD019_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD019_NEWTEMP.ctl  log=CD019_NEWTEMP.log

rem 17. Load CD020���s�N�X���: CD020_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD020_NEWTEMP.ctl  log=CD020_NEWTEMP.log

rem 18. Load CD021���s�N�X���: CD021_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD021_NEWTEMP.ctl  log=CD021_NEWTEMP.log

rem 19. Load CD022���s�N�X���: CD022_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD022_NEWTEMP.ctl  log=CD022_NEWTEMP.log

rem 20. Load CD025���s�N�X���: CD025_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD025_NEWTEMP.ctl  log=CD025_NEWTEMP.log

rem 21. Load CD026���s�N�X���: CD026_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD026_NEWTEMP.ctl  log=CD026_NEWTEMP.log

rem 22. Load CD049���s�N�X���: CD049_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD049_NEWTEMP.ctl  log=CD049_NEWTEMP.log

rem 23. Load CD051���s�N�X���: CD051_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD051_NEWTEMP.ctl  log=CD051_NEWTEMP.log

rem 24. Load CD055���s�N�X���: CD055_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD055_NEWTEMP.ctl  log=CD055_NEWTEMP.log

rem 25. Load CD059���s�N�X���: CD059_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD059_NEWTEMP.ctl  log=CD059_NEWTEMP.log