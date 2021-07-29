PROMPT *** SO461   ∫ ∫ﬁGateWay Request Log¿… ***

DROP TABLE SO461 CASCADE CONSTRAINT;

CREATE TABLE SO461 (
	CompCode NUMBER(3),
	ModeType NUMBER(2),
	HandleEDateTime DATE,
	QueryDateTime DATE,
	ExecDateTime DATE,
	InstQuery NUMBER(1),
	SrcCode VARCHAR2(10),
	DstCode VARCHAR2(10),
	DstMsgID VARCHAR2(10),
	XmlContent VARCHAR2(1000));
