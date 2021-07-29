PROMPT *** SO456   監控GateWay產品定購資訊資料(即時查詢資料) ***

DROP TABLE SO456 CASCADE CONSTRAINT;

CREATE TABLE SO456 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	Status NUMBER(1),
	PurchaseID VARCHAR2(15),
	ICCardNo VARCHAR2(20),
	ICCardID VARCHAR2(20),
	CasID VARCHAR2(10),
	ProductCode VARCHAR2(10),
	StartTime DATE,
	EndTime DATE);
