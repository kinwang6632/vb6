
rem ***** Load EMC�s�W�N�X *****
rem del *.log
rem del *.bad

rem 0.�R������Table
Sqlplusw.exe v30/v30@emc @C:\Transform\EMC_Transform2\DeleteTable.sql

rem 1. Load ���O���إN�X��(CD019)���s�N�X : CD019
sqlldr userid=v30/v30@emc control=CD019_Map.ctl  log=CD019_Map.log

rem 2. Load �W�D�s���N�X��(CD024)���s�N�X : CD024
sqlldr userid=v30/v30@emc control=CD024_Map.ctl  log=CD024_Map.log

rem 3. Load ���O�W�D�Ӷ���(CD019A)���s�N�X : CD019A
sqlldr userid=v30/v30@emc control=CD019A_Map.ctl  log=CD019A_Map.log

rem 4. Load �˾����O�N�X��(CD005)���s�N�X : CD005
sqlldr userid=v30/v30@emc control=CD005_Map.ctl  log=CD005_Map.log

rem 5. Load ��������O�N�X��(CD007)���s�N�X : CD007
sqlldr userid=v30/v30@emc control=CD007_Map.ctl  log=CD007_Map.log

rem 6. Load �~�W�s���N�X��(CD022)���s�N�X : CD022
sqlldr userid=v30/v30@emc control=CD022_Map.ctl  log=CD022_Map.log

rem 7. Load �]�ƶR��I�ڤ覡������(CD022A)���s�N�X : CD022A
sqlldr userid=v30/v30@emc control=CD022A_Map.ctl  log=CD022A_Map.log

rem 8. Load �q�ܥӧi���O�N�X��(CD008)���s�N�X : CD008
sqlldr userid=v30/v30@emc control=CD008_Map.ctl  log=CD008_Map.log

rem 9. Load �ӧi���e�N�X��(CD008A)���s�N�X : CD008A
sqlldr userid=v30/v30@emc control=CD008A_Map.ctl  log=CD008A_Map.log

rem 10. Load �ӧi���O���ӧi���e(CD008C)���s�N�X : CD008C
sqlldr userid=v30/v30@emc control=CD008C_Map.ctl  log=CD008C_Map.log

rem 11. Load �G�٭�]�N�X��(CD011)���s�N�X : CD011
sqlldr userid=v30/v30@emc control=CD011_Map.ctl  log=CD011_Map.log

rem 12. Load �]�Ƹ˸m�I�N�X��(CD056)���s�N�X : CD056
sqlldr userid=v30/v30@emc control=CD056_Map.ctl  log=CD056_Map.log

rem 13. Load �C�餶�����O�N�X��(CD009)���s�N�X : CD009
sqlldr userid=v30/v30@emc control=CD009_Map.ctl  log=CD009_Map.log

rem 14. Load �I�ں����N�X��(CD032)���s�N�X : CD032
sqlldr userid=v30/v30@emc control=CD032_Map.ctl  log=CD032_Map.log

rem 15. Load STB���M��]�N�X(CD061)���s�N�X : CD061
sqlldr userid=v30/v30@emc control=CD061_Map.ctl  log=CD061_Map.log

rem 16. Load �W�D���M��]�N�X(CD060)���s�N�X : CD060
sqlldr userid=v30/v30@emc control=CD060_Map.ctl  log=CD060_Map.log

rem 17. Load STB�t��N�X��(CD028)���s�N�X : CD028
sqlldr userid=v30/v30@emc control=CD028_Map.ctl  log=CD028_Map.log

rem 18. Load �q�����O�N�X��(CD057)���s�N�X : CD057
sqlldr userid=v30/v30@emc control=CD057_Map.ctl  log=CD057_Map.log







