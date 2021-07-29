--描述哪些代碼檔,用到哪些資料檔的哪些欄位

DROP TABLE Code_Check1 CASCADE CONSTRAINT;
create table Code_Check1 (
	CodeTable varchar2(30),
	DataTable varchar2(30),
	DataColumn varchar2(30),
        DataChName  varchar2(100));


DROP TABLE Code_Check2 CASCADE CONSTRAINT;
create table Code_Check2 (
	CodeTable  varchar2(30),
	DataTable  varchar2(30),
	DataColumn varchar2(30),
        DataChColumn varchar2(100),
	CodeValue  varchar2(30),
        CodeName   varchar2(100),
	Description varchar2(500),
        RcdCnt  number(8) DEFAULT 0,
        CheckTime  Date);

DROP TABLE Code_Check3 CASCADE CONSTRAINT;
create table Code_Check3 (
	CodeTable varchar2(30),
	DataTable varchar2(30),
	DataColumn varchar2(30),
        DataChName  varchar2(100));


exit



