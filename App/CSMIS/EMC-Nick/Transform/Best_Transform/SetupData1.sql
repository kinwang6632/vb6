/*
  @C:\Transform\Best_Transform\SetupData1.SQL

  �إ߰򥻸�Ƥ@: ���e�Ӧۥ_��CATV��ư�, �ݩ�g��Ϭ����ɮפ��e

  By: Jackal
  Date: 2002.11.25
*/
set serveroutput on
set heading off
exec :t1 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< �إ߰򥻸�Ƥ@ >> ' || :t1;
spool C:\Transform\Best_Transform\Create1.log
print str 




Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD001
INSERT INTO CD001(SELECT * FROM BEST.CD001);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD001A
INSERT INTO CD001A(SELECT * FROM BEST.CD001A);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD002
INSERT INTO CD002(SELECT * FROM BEST.CD002);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD002CM003
INSERT INTO CD002CM003(SELECT * FROM BEST.CD002CM003);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD003
INSERT INTO CD003(SELECT * FROM BEST.CD003);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD017
INSERT INTO CD017(SELECT * FROM BEST.CD017);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD023
INSERT INTO CD023(SELECT * FROM BEST.CD023);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD040
INSERT INTO CD040(SELECT * FROM BEST.CD040);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD045
INSERT INTO CD045(SELECT * FROM BEST.CD045);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD050
INSERT INTO CD050(SELECT * FROM BEST.CD050);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD052
INSERT INTO CD052(SELECT * FROM BEST.CD052);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CD058
INSERT INTO CD058(SELECT * FROM BEST.CD058);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: CM003
INSERT INTO CM003(SELECT * FROM BEST.CM003);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: SO014
INSERT INTO SO014(SELECT * FROM BEST.SO014);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: SO016
INSERT INTO SO016(SELECT * FROM BEST.SO016);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: SO017
INSERT INTO SO017(SELECT * FROM BEST.SO017);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: SO026
INSERT INTO SO026(SELECT * FROM BEST.SO026);

Prompt �إ߰򥻸�Ƥ@,���e�Ӧۥ_��CATV��ư� Table: SO029
INSERT INTO SO029(SELECT * FROM BEST.SO029);

COMMIT;




exec :t2 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< ��ƫإߧ��� >> ' || :t2;
print str
exec :tt := TRUNC(86400*(to_date(:t2, 'YYYY/MM/DD HH24:MI:SS')-to_date(:t1, 'YYYY/MM/DD HH24:MI:SS')));
exec :hh := trunc(:tt/3600);
exec :ss := mod(:tt, 60);
exec :mm := (:tt - :hh*3600 - :ss) / 60;


exec :str := '�X�p����ɶ�: ' || :hh || '��' || :mm || '��' || :ss || '��';
print str 

spool off
set heading on
EXIT
