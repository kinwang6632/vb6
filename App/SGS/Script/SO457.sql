PROMPT *** SO457   �ʱ�GateWay���v��T���(���v���) ***

DROP TABLE SO457 CASCADE CONSTRAINT;

CREATE TABLE SO457 (
	CompCode NUMBER(3),
	HandleEDateTime DATE,
	RealDateTime DATE,
	ExecDateTime DATE,
	Status NUMBER(1),
	ICCardNo VARCHAR2(20),
	ICCardID VARCHAR2(20),
	CasID VARCHAR2(10),
	ProductCode VARCHAR2(10));
