PROMPT *** SO452   �ʱ�GateWayIC�d�T�����(�Y�ɬd�߸��) ***

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
