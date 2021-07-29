PROMPT *** CD067C   分類明細對照檔 ***

DROP TABLE CD067C CASCADE CONSTRAINT;

CREATE TABLE CD067C (
	CompCode NUMBER(3),
	TableName VARCHAR2(30),
	MasterCodeNo NUMBER(3),
	DetailCodeNo NUMBER(3),
	StopFlag NUMBER(1),
	OPERATOR VARCHAR2(20),
	UPDTIME VARCHAR2(20));

CREATE INDEX I_CD067C_1 ON CD067C (CompCode,TableName)
	PCTFREE 10 STORAGE (INITIAL 100K NEXT 300K MINEXTENTS 1 MAXEXTENTS 200) TABLESPACE GICMIS_NDX_V30;
