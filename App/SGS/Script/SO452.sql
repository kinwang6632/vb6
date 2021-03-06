PROMPT *** SO452   監控GateWayIC卡訊息資料(即時查詢資料) ***

DROP TABLE SO452 CASCADE CONSTRAINT;

CREATE TABLE SO452 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	Status NUMBER(1),
	ICCardNo VARCHAR2(20),
	ICCardID VARCHAR2(20),
	CasID VARCHAR2(10),
	Feature VARCHAR2(300),
	Version VARCHAR2(10),
	SubName VARCHAR2(50),
	SubVocation VARCHAR2(30),
	SubAge NUMBER(3),
	SubIDCardNo VARCHAR2(15));
