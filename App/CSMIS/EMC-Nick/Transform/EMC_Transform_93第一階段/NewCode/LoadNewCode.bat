rem ***** 91.11.22 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
Sqlplusw.exe emctc/emctc@ty @C:\Transform\NewCode\TruncateCodeTable.sql


rem ***** �إ߷s�N�X Table ��� *****

rem del *.log
rem del *.bad

rem 1. Load CD009���s�N�X���: CD009_NEW
sqlldr userid=emctc/emctc@ty control=CD009_NEW.ctl  log=CD009_NEW.log

rem 2. Load CD012���s�N�X���: CD012_NEW
sqlldr userid=emctc/emctc@ty control=CD012_NEW.ctl  log=CD012_NEW.log

rem 3. Load CD013���s�N�X���: CD013_NEW
sqlldr userid=emctc/emctc@ty control=CD013_NEW.ctl  log=CD013_NEW.log

rem 4. Load CD015���s�N�X���: CD015_NEW
sqlldr userid=emctc/emctc@ty control=CD015_NEW.ctl  log=CD015_NEW.log

rem 5. Load CD016���s�N�X���: CD016_NEW
sqlldr userid=emctc/emctc@ty control=CD016_NEW.ctl  log=CD016_NEW.log

rem 6. Load CD020���s�N�X���: CD020_NEW
sqlldr userid=emctc/emctc@ty control=CD020_NEW.ctl  log=CD020_NEW.log

rem 7. Load CD021���s�N�X���: CD021_NEW
sqlldr userid=emctc/emctc@ty control=CD021_NEW.ctl  log=CD021_NEW.log

rem 8. Load CD026���s�N�X���: CD026_NEW
sqlldr userid=emctc/emctc@ty control=CD026_NEW.ctl  log=CD026_NEW.log

rem 9. Load CD051���s�N�X���: CD051_NEW
sqlldr userid=emctc/emctc@ty control=CD051_NEW.ctl  log=CD051_NEW.log

