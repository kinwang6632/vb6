PROMPT *** SO459   ∫ ∫ﬁGateWay NightRun ß@∑~Log¿… ***

DROP TABLE SO459 CASCADE CONSTRAINT;

CREATE TABLE SO459 (
	CompCode NUMBER(3),
	ModeType NUMBER(2),
	FirstFlag NUMBER(1),
	HandleEDateTime DATE,
	ExecDateTime DATE,
	TotalICCardNum NUMBER(10),
	TotalSubNum NUMBER(10),
	TotalProductNum NUMBER(10),
	TotalChannelNum NUMBER(10),
	ErrorCode VARCHAR2(20),
	ErrorMsg VARCHAR2(2000));
