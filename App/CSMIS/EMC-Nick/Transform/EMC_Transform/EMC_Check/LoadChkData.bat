rem ***** 91.11.22 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
Sqlplusw.exe v30/v30@mis @C:\Transform\EMC_Transform\EMC_Check\AddCodeTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad

rem 1. Load Table �P Column ��Ӫ�: Code_Map1
sqlldr userid=v30/v30@mis control=Code_Map1.ctl  log=Code_Map1.log














