/*
  @C:\Transform\Best_Transform\SetupData1.SQL

  建立基本資料一: 內容來自北視CATV資料區, 屬於經營區相關檔案內容

  By: Jackal
  Date: 2002.11.25
*/
set serveroutput on
set heading off
exec :t1 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< 建立基本資料一 >> ' || :t1;
spool C:\Transform\Best_Transform\Create1.log
print str 




Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD001
INSERT INTO CD001(SELECT * FROM BEST.CD001);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD001A
INSERT INTO CD001A(SELECT * FROM BEST.CD001A);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD002
INSERT INTO CD002(SELECT * FROM BEST.CD002);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD002CM003
INSERT INTO CD002CM003(SELECT * FROM BEST.CD002CM003);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD003
INSERT INTO CD003(SELECT * FROM BEST.CD003);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD017
INSERT INTO CD017(SELECT * FROM BEST.CD017);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD023
INSERT INTO CD023(SELECT * FROM BEST.CD023);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD040
INSERT INTO CD040(SELECT * FROM BEST.CD040);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD045
INSERT INTO CD045(SELECT * FROM BEST.CD045);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD050
INSERT INTO CD050(SELECT * FROM BEST.CD050);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD052
INSERT INTO CD052(SELECT * FROM BEST.CD052);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CD058
INSERT INTO CD058(SELECT * FROM BEST.CD058);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: CM003
INSERT INTO CM003(SELECT * FROM BEST.CM003);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: SO014
INSERT INTO SO014(SELECT * FROM BEST.SO014);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: SO016
INSERT INTO SO016(SELECT * FROM BEST.SO016);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: SO017
INSERT INTO SO017(SELECT * FROM BEST.SO017);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: SO026
INSERT INTO SO026(SELECT * FROM BEST.SO026);

Prompt 建立基本資料一,內容來自北視CATV資料區 Table: SO029
INSERT INTO SO029(SELECT * FROM BEST.SO029);

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
