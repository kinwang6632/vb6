PROMPT *** CD024A   ?????W?D?N?X?? ***

DROP TABLE CD024A CASCADE CONSTRAINT;

CREATE TABLE CD024A (
	CodeNo VARCHAR2(10),
	Description VARCHAR2(50) NOT NULL,
	RefNo NUMBER(3),
	UpdTime VARCHAR2(19),
	UpdEn VARCHAR2(20),
	CompCode NUMBER(3) NOT NULL,
	StopFlag NUMBER(1) NOT NULL,
	CONSTRAINT PK_CD024A PRIMARY KEY (CodeNo));
