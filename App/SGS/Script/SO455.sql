PROMPT *** SO455   �ʱ�GateWay���~�w�ʸ�T���(���v���) ***

DROP TABLE SO455 CASCADE CONSTRAINT;

CREATE TABLE SO455 (
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
