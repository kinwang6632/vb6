PROMPT *** SO453   �ʱ�GateWay���~�w�q�T�����(���v���) ***

DROP TABLE SO453 CASCADE CONSTRAINT;

CREATE TABLE SO453 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	ProductCode VARCHAR2(10),
	ProductName VARCHAR2(100),
	CasID VARCHAR2(10),
	Provider VARCHAR2(40),
	ChannelName VARCHAR2(40));
