PROMPT *** So124  佣金計算之已鎖帳年月檔 ***

DROP TABLE So124 CASCADE CONSTRAINT;

CREATE TABLE So124 (
	CompCode		Number(3),
	LockYM			Varchar2(5),
	OPERATOR		VARCHAR2(20),
	UPDTIME		  	VARCHAR2(20)
       );

