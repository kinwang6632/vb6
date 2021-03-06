PROMPT *** CD067B   ???O?????? ***

DROP TABLE CD067B CASCADE CONSTRAINT;

CREATE TABLE CD067B (
	CodeNo NUMBER(3),
	TableName VARCHAR2(30),
	Description VARCHAR2(20),
	RefNo NUMBER(3),
	ServiceType CHAR(1),
	StopFlag NUMBER(1),
	OPERATOR VARCHAR2(20),
	UPDTIME VARCHAR2(20));

CREATE INDEX I_CD067B_1 ON CD067B (CodeNo)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 300K MINEXTENTS 1 MAXEXTENTS 200) TABLESPACE NTY_NDX;
