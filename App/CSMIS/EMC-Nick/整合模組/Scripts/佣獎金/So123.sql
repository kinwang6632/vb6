PROMPT *** So123  客戶自來電佣金發放資料檔 ***

DROP TABLE So123 CASCADE CONSTRAINT;

CREATE TABLE So123 (
	CompCode	Number(3),
	CustID	        Number(8),
	ComputeYM	Varchar2(5),
	CommDate	Date,
	ChStopDate      Date,
	CsrID           VARCHAR2(20),
	CsrName         VARCHAR2(20),
	Comm            Number(10,2),
	OPERATOR	VARCHAR2(20),
	UPDTIME		VARCHAR2(20),
	CONSTRAINT PK_So123 PRIMARY KEY (CompCode,CustID)
       );
