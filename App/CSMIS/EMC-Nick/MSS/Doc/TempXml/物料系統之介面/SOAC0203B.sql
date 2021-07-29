PROMPT *** SOAC0203B   ICC���P�ƥ��� ***

DROP TABLE SOAC0203B CASCADE CONSTRAINT;

CREATE TABLE SOAC0203B (
	CompCode NUMBER(3),
	BatchNo VARCHAR2(15),
	ExecTime DATE,
	OldBatchNo VARCHAR2(15),
	MaterialNo VARCHAR2(15),
	FaciSNo VARCHAR2(15) NOT NULL,
	Description VARCHAR2(30),
	Spec VARCHAR2(50),
	VendorCode VARCHAR2(5),
	VendorName VARCHAR2(20),
	UpdTime VARCHAR2(19),
	UpdEn VARCHAR2(20));
