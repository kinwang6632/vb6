ROMPT *** SO508   A1 ight Run ***

DROP TABLE SO508 CASCADE CONSTRAINT;

create table SO508 (
	RptYM number(6), 
	CompCode number(3), 
	ServiceType char(1), 
	CitemCode number(3), 
	CitemName varchar2(20), 
	RealAmt number(10), 
	MonthPast number(10), 
	Month01 number(10), 
	CustCnt	number(8));

PROMPT *** SO508A   A1報表異常資料檔 ***

DROP TABLE SO508A CASCADE CONSTRAINT;

CREATE TABLE SO508A (
	BillNo VARCHAR2(15),
	Item NUMBER(3),
	CustId NUMBER(8),
	CitemCode NUMBER(3),
	CitemName VARCHAR2(20),
	RealDate DATE,
	RealStartDate DATE,
	RealStopDate DATE,
	RealPeriod NUMBER(3),
	RealAmt NUMBER(8),
	ErrorType VARCHAR2(50));
