PROMPT *** SO454   監控GateWay產品定義訊息資料(即時查詢資料) ***

DROP TABLE SO454 CASCADE CONSTRAINT;

CREATE TABLE SO454 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	ProductCode VARCHAR2(10),
	ProductName VARCHAR2(100),
	CasID VARCHAR2(10),
	Provider VARCHAR2(40),
	ChannelName VARCHAR2(40));
