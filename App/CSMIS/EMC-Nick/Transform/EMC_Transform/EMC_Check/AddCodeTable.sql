--描述哪些代碼檔,用到哪些資料檔的哪些欄位

DROP TABLE Code_Map1 CASCADE CONSTRAINT;
create table Code_Map1 (
	CodeTable varchar2(30),
	DataTable varchar2(30),
	DataColumn varchar2(30));

CREATE INDEX I_Code_Map1_Index1 ON Code_Map1 (CodeTable,DataTable,DataColumn)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) ;



DROP TABLE Code_Map2 CASCADE CONSTRAINT;
create table Code_Map2 (
	CodeTable  varchar2(30),
	DataTable  varchar2(30),
	DataColumn varchar2(30),
	OldCodeNo  number(3),
	OldName    varchar2(20),
	OldRcdCnt  number(8) DEFAULT 0,
	NewCodeNo  number(3),
	NewName    varchar2(20),
	NewRcdCnt  number(8) DEFAULT 0);

CREATE INDEX I_Code_Map2_Index1 ON Code_Map2 (CodeTable,DataTable,DataColumn)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 500K MINEXTENTS 1 MAXEXTENTS UNLIMITED) ;
exit



