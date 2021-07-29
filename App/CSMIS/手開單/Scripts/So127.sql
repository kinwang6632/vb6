PROMPT *** So127  手開單之單據資料檔 ***

DROP TABLE So127 CASCADE CONSTRAINT;

CREATE TABLE So127 (
	CompCode		Number(3),
	Seq			Number(10),
	PaperNum		VARCHAR2(15),
	EmpNO		 	VARCHAR2(20),
	EmpName			VARCHAR2(30),
	GetPaperDate		Date,
	Status			Number(1),
	CustID			Number(8),
	CustName		Varchar2(20),
	CustTEL			Varchar2(20),
	BillNo			VARCHAR2(15),
	RealDate		Date,
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20),
	CONSTRAINT PK_So127 PRIMARY KEY (CompCode,Seq,PaperNum)
       );

DROP INDEX V30.I_So127_1;

CREATE INDEX V30.I_So127_1 ON SO127 (BillNo) TABLESPACE GICMIS_NDX_V30;
