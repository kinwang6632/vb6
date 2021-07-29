PROMPT *** SO509   紀錄要被 move 到Snapshot site 上的table ***

DROP TABLE SO509 CASCADE CONSTRAINT;

CREATE TABLE SO509 (
	TableCode VARCHAR2(30),
	TableName VARCHAR2(60));
