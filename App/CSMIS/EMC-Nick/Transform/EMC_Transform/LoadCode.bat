rem ***** 91.11.21 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
Sqlplusw.exe kp/kp@emc @C:\Transform\EMC_Transform\AddTmpTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad

rem 1. Load �Ȥ����O�N�X��(CD004)���s�¹�Ӫ�: CD004_Map
sqlldr userid=kp/kp@emc control=CD004_Map.ctl  log=CD004_Map.log

rem 2. Load �Ȥ����O�N�X��(CD004)���s�N�X��: CD004_New
sqlldr userid=kp/kp@emc control=CD004_New.ctl  log=CD004_New.log



rem 3. Load ���O���إN�X��(CD019)���s�¹�Ӫ�: CD019_Map
sqlldr userid=kp/kp@emc control=CD019_Map.ctl  log=CD019_Map.log

rem 4. Load ���O���إN�X��(CD019)���s�N�X��: CD019_New
sqlldr userid=kp/kp@emc control=CD019_New.ctl  log=CD019_New.log



rem 5. Load �˾����O�N�X��(CD005)���s�¹�Ӫ�: CD005_Map
sqlldr userid=kp/kp@emc control=CD005_Map.ctl  log=CD005_Map.log

rem 6. Load �˾����O�N�X��(CD005)���s�N�X��: CD005_New
sqlldr userid=kp/kp@emc control=CD005_New.ctl  log=CD005_New.log



rem 7. Load ���O�覡�N�X��(CD031)���s�¹�Ӫ�: CD031_Map
rem sqlldr userid=kp/kp@emc control=CD031_Map.ctl  log=CD031_Map.log

rem 8. Load ���O�覡�N�X��(CD031)���s�N�X��: CD031_New
rem sqlldr userid=kp/kp@emc control=CD031_New.ctl  log=CD031_New.log



rem 9. Load �R��覡�N�X��(CD034)���s�¹�Ӫ�: CD034_Map
rem sqlldr userid=kp/kp@emc control=CD034_Map.ctl  log=CD034_Map.log

rem 10. Load �R��覡�N�X��(CD034)���s�N�X��: CD034_New
rem sqlldr userid=kp/kp@emc control=CD034_New.ctl  log=CD034_New.log



rem 11. Load ��������O�N�X��(CD007)���s�¹�Ӫ�: CD007_Map
sqlldr userid=kp/kp@emc control=CD007_Map.ctl  log=CD007_Map.log

rem 12. Load ��������O�N�X��(CD007)���s�N�X��: CD007_New
sqlldr userid=kp/kp@emc control=CD007_New.ctl  log=CD007_New.log



rem 13. Load �I�O�N�@�N�X��(CD020)���s�¹�Ӫ�: CD020_Map
sqlldr userid=kp/kp@emc control=CD020_Map.ctl  log=CD020_Map.log

rem 14. Load �I�O�N�@�N�X��(CD020)���s�N�X��: CD020_New
sqlldr userid=kp/kp@emc control=CD020_New.ctl  log=CD020_New.log



rem 15. Load �ؿv���A�N�X��(CD021)���s�¹�Ӫ�: CD021_Map
sqlldr userid=kp/kp@emc control=CD021_Map.ctl  log=CD021_Map.log

rem 16. Load �ؿv���A�N�X��(CD021)���s�N�X��: CD021_New
sqlldr userid=kp/kp@emc control=CD021_New.ctl  log=CD021_New.log



rem 17. Load �C�餶�����O�N�X��(CD009)���s�¹�Ӫ�: CD009_Map
sqlldr userid=kp/kp@emc control=CD009_Map.ctl  log=CD009_Map.log

rem 18. Load �C�餶�����O�N�X��(CD009)���s�N�X��: CD009_New
sqlldr userid=kp/kp@emc control=CD009_New.ctl  log=CD009_New.log















