rem ***** 92.12.27 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
Sqlplusw.exe emctc/emctc@ty @C:\Transform\Map\AddTmpTable.sql


rem ***** Load �ഫ�N�X��������� *****
rem del *.log
rem del *.bad


rem 1. Load CD009���s�¹�Ӫ�: CD009_MAP
sqlldr userid=emctc/emctc@ty control=CD009_MAP.ctl  log=CD009_MAP.log


rem 2. Load CD012���s�¹�Ӫ�: CD012_MAP
sqlldr userid=emctc/emctc@ty control=CD012_MAP.ctl  log=CD012_MAP.log


rem 3. Load CD013���s�¹�Ӫ�: CD013_MAP
sqlldr userid=emctc/emctc@ty control=CD013_MAP.ctl  log=CD013_MAP.log


rem 4. Load CD015���s�¹�Ӫ�: CD015_MAP
sqlldr userid=emctc/emctc@ty control=CD015_MAP.ctl  log=CD015_MAP.log


rem 5. Load CD016���s�¹�Ӫ�: CD016_MAP
sqlldr userid=emctc/emctc@ty control=CD016_MAP.ctl  log=CD016_MAP.log


rem 6. Load CD020���s�¹�Ӫ�: CD020_MAP
sqlldr userid=emctc/emctc@ty control=CD020_MAP.ctl  log=CD020_MAP.log


rem 7. Load CD021���s�¹�Ӫ�: CD021_MAP
sqlldr userid=emctc/emctc@ty control=CD021_MAP.ctl  log=CD021_MAP.log


rem 8. Load CD026���s�¹�Ӫ�: CD026_MAP
sqlldr userid=emctc/emctc@ty control=CD026_MAP.ctl  log=CD026_MAP.log


rem 9. Load CD051���s�¹�Ӫ�: CD051_MAP
sqlldr userid=emctc/emctc@ty control=CD051_MAP.ctl  log=CD051_MAP.log


