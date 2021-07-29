/*
  @C:\Transform\Best_Transform\SetupData2.SQL

  建立基本資料二: 內容來自LSC CM/DVS資料區, 屬於服務別相關檔案內容

  By: Jackal
  Date: 2002.11.20
*/
set serveroutput on
set heading off
exec :t1 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< 建立基本資料二 >> ' || :t1;
spool C:\Transform\Best_Transform\Create2.log
print str 




Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD004
INSERT INTO CD004(SELECT * FROM LSC.CD004);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD005
INSERT INTO CD005(SELECT * FROM LSC.CD005);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD006
INSERT INTO CD006(SELECT * FROM LSC.CD006);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD006A
INSERT INTO CD006A(SELECT * FROM LSC.CD006A);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD006B
INSERT INTO CD006B(SELECT * FROM LSC.CD006B);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD007
INSERT INTO CD007(SELECT * FROM LSC.CD007);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD008
INSERT INTO CD008(SELECT * FROM LSC.CD008);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD008A
INSERT INTO CD008A(SELECT * FROM LSC.CD008A);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD008C
INSERT INTO CD008C(SELECT * FROM LSC.CD008C);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD008B
INSERT INTO CD008B(SELECT * FROM LSC.CD008B);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD009
INSERT INTO CD009(SELECT * FROM LSC.CD009);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD010
INSERT INTO CD010(SELECT * FROM LSC.CD010);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD011
INSERT INTO CD011(SELECT * FROM LSC.CD011);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD011A
INSERT INTO CD011A(SELECT * FROM LSC.CD011A);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD011B
INSERT INTO CD011B(SELECT * FROM LSC.CD011B);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD012
INSERT INTO CD012(SELECT * FROM LSC.CD012);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD013
INSERT INTO CD013(SELECT * FROM LSC.CD013);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD014
INSERT INTO CD014(SELECT * FROM LSC.CD014);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD015
INSERT INTO CD015(SELECT * FROM LSC.CD015);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD016
INSERT INTO CD016(SELECT * FROM LSC.CD016);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD018
INSERT INTO CD018(SELECT * FROM LSC.CD018);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD019
INSERT INTO CD019(SELECT * FROM LSC.CD019);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD019CD004
INSERT INTO CD019CD004(SELECT * FROM LSC.CD019CD004);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD020
INSERT INTO CD020(SELECT * FROM LSC.CD020);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD021
INSERT INTO CD021(SELECT * FROM LSC.CD021);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD022
INSERT INTO CD022(SELECT * FROM LSC.CD022);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD022CD019
INSERT INTO CD022CD019(SELECT * FROM LSC.CD022CD019);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD024
INSERT INTO CD024(SELECT * FROM LSC.CD024);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD025
INSERT INTO CD025(SELECT * FROM LSC.CD025);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD026
INSERT INTO CD026(SELECT * FROM LSC.CD026);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD031
INSERT INTO CD031(SELECT * FROM LSC.CD031);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD031A
INSERT INTO CD031A(SELECT * FROM LSC.CD031A);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD032
INSERT INTO CD032(SELECT * FROM LSC.CD032);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD033
INSERT INTO CD033(SELECT * FROM LSC.CD033);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD034
INSERT INTO CD034(SELECT * FROM LSC.CD034);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD035
INSERT INTO CD035(SELECT * FROM LSC.CD035);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD036
INSERT INTO CD036(SELECT * FROM LSC.CD036);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD037
INSERT INTO CD037(SELECT * FROM LSC.CD037);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD043
INSERT INTO CD043(SELECT * FROM LSC.CD043);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD044
INSERT INTO CD044(SELECT * FROM LSC.CD044);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD045
INSERT INTO CD045(SELECT * FROM LSC.CD045);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD046
INSERT INTO CD046(SELECT * FROM LSC.CD046);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD049
INSERT INTO CD049(SELECT * FROM LSC.CD049);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD051
INSERT INTO CD051(SELECT * FROM LSC.CD051);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD053
INSERT INTO CD053(SELECT * FROM LSC.CD053);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD054
INSERT INTO CD054(SELECT * FROM LSC.CD054);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD056
INSERT INTO CD056(SELECT * FROM LSC.CD056);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CD059
INSERT INTO CD059(SELECT * FROM LSC.CD059);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CM001
INSERT INTO CM001(SELECT * FROM LSC.CM001);

Prompt 建立基本資料二,內容來自LSC CM/DVS資料區 Table: CM002
INSERT INTO CM002(SELECT * FROM LSC.CM002);


COMMIT;





exec :t2 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< 資料建立完成 >> ' || :t2;
print str
exec :tt := TRUNC(86400*(to_date(:t2, 'YYYY/MM/DD HH24:MI:SS')-to_date(:t1, 'YYYY/MM/DD HH24:MI:SS')));
exec :hh := trunc(:tt/3600);
exec :ss := mod(:tt, 60);
exec :mm := (:tt - :hh*3600 - :ss) / 60;


exec :str := '合計執行時間: ' || :hh || '時' || :mm || '分' || :ss || '秒';
print str 

spool off
set heading on
EXIT





