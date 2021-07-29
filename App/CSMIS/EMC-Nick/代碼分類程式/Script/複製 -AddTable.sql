PROMPT *** CateCodeList1  分類代碼檔 ***

DROP TABLE CateCodeList1 CASCADE CONSTRAINT;

CREATE TABLE CateCodeList1 (
	TableName		VARCHAR2(30),
	TableDescription	Varchar2(50),
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20),
	CONSTRAINT PK_CateCodeList1 PRIMARY KEY (TableName)
       );

DROP INDEX I_CateCodeList1_1;

CREATE INDEX I_CateCodeList1_1 ON CateCodeList1 (TableName);


PROMPT *** CateCodeList2  類別資料檔 ***

DROP TABLE CateCodeList2 CASCADE CONSTRAINT;

CREATE TABLE CateCodeList2 (
	CodeNo			Number(3),
	TableName		VARCHAR2(30),
	Description		VARCHAR2(20),
	RefNo			Number(3),
	ServiceType		Char(1),
	StopFlag		Number(1),
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20)
       );

DROP INDEX I_CateCodeList2_1;

CREATE INDEX I_CateCodeList2_1 ON CateCodeList2 (CodeNo);


PROMPT *** CateCodeList3  分類明細對照檔 ***

DROP TABLE CateCodeList3 CASCADE CONSTRAINT;

CREATE TABLE CateCodeList3 (
	CompCode		Number(3),
	TableName		VARCHAR2(30),
	MasterCodeNo		Number(3),
	DetailCodeNo		Number(3),
	StopFlag		Number(1),
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20)
       );

DROP INDEX I_CateCodeList3_1;

CREATE INDEX I_CateCodeList3_1 ON CateCodeList3 (CompCode,TableName) ;
