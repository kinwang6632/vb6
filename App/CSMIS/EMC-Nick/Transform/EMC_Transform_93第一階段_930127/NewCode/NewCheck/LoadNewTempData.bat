Sqlplusw.exe emcct/emcct@mis @C:\Transform\NewCode\NewCheck\AddNewTempTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad

rem 0. Load  : Code_Check3
sqlldr userid=emcct/emcct@mis control=Code_Check3.ctl  log=Code_Check3.log

rem 1. Load �s�N�X : CD009_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD009_NEWTEMP.ctl  log=CD009_NEWTEMP.log

rem 2. Load �s�N�X : CD012_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD012_NEWTEMP.ctl  log=CD012_NEWTEMP.log

rem 3. Load �s�N�X : CD013_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD013_NEWTEMP.ctl  log=CD013_NEWTEMP.log

rem 4. Load �s�N�X : CD015_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD015_NEWTEMP.ctl  log=CD015_NEWTEMP.log

rem 5. Load �s�N�X : CD016_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD016_NEWTEMP.ctl  log=CD016_NEWTEMP.log

rem 6. Load �s�N�X : CD020_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD020_NEWTEMP.ctl  log=CD020_NEWTEMP.log

rem 7. Load �s�N�X : CD021_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD021_NEWTEMP.ctl  log=CD021_NEWTEMP.log

rem 8. Load �s�N�X : CD026_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD026_NEWTEMP.ctl  log=CD026_NEWTEMP.log

rem 9. Load �s�N�X : CD051_NEWTEMP
sqlldr userid=emcct/emcct@mis control=CD051_NEWTEMP.ctl  log=CD051_NEWTEMP.log













