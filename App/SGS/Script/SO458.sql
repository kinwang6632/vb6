PROMPT *** SO458   監控GateWay授權資訊資料(即時查詢資料) ***

DROP TABLE SO458 CASCADE CONSTRAINT;

CREATE TABLE SO458 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	Status NUMBER(1),
	ICCardNo VARCHAR2(20),
	ICCardID VARCHAR2(20),
	CasID VARCHAR2(10),
	ProductCode VARCHAR2(10));
