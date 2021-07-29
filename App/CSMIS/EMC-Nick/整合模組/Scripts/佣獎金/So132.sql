PROMPT *** So132  佣金計算之信用卡付費時之佣金歸屬人員檔 ***

DROP TABLE So132 CASCADE CONSTRAINT;

CREATE TABLE So132 (
	CompCode		Number(3),
	CUSTID			Number(8),
	CITEMCODE		Number(3),
	CITEMNAME		VARCHAR2(20),
	EMPNO			VARCHAR2(10),
	EMPNAME			VARCHAR2(20),
	DataCreator		VARCHAR2(20),
	UpdEn			VARCHAR2(20),
	UpdTime		  	Date
       );

