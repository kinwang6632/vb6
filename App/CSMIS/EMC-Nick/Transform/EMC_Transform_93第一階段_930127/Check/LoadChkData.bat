rem ***** 91.11.22 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
Sqlplusw.exe emcct/emcct@mis @C:\Transform\Check\AddCodeTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad

rem 1. Load Table �P Column ��Ӫ�: Code_Check1
sqlldr userid=emcct/emcct@mis control=Code_Check1.ctl  log=Code_Check1.log














