/*
  ADD_Table.SQL: Create Tables
  Naming rule: INVXXX
  Date: 2000.12.11
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool C:\temp
print str 

/*
TABLESPACE 之值需要做適當之修改
*/


/*
	Callback Log (指令 200)檔 => CallbackLog200
*/

PROMPT *** Callback Log (指令 200)檔 ***
DROP TABLE CallbackLog200 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog200 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	Credit			VARCHAR2(7),
	Debit			VARCHAR2(7),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_200_1;
CREATE INDEX I_Callback_Log_200_1 on CallbackLog200 (TransactionNum)
TABLESPACE EMC;



PROMPT *** Callback Log (指令 201)檔 ***
DROP TABLE CallbackLog201 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog201 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	Credit			VARCHAR2(7),
	Debit			VARCHAR2(7),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_201_1;
CREATE INDEX I_Callback_Log_201_1 on CallbackLog201 (TransactionNum)
TABLESPACE EMC;



PROMPT *** Callback Log (指令 202)檔 ***
DROP TABLE CallbackLog202 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog202 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	ImsProductID		VARCHAR2(12),
	PurchaseDate		Date,
	WatchedStatus		CHAR(1),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_202_1;
CREATE INDEX I_Callback_Log_202_1 on CallbackLog202 (TransactionNum)
TABLESPACE EMC;




PROMPT *** Callback Log (指令 205)檔 ***
DROP TABLE CallbackLog205 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog205 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	PhoneNum1		VARCHAR2(16),
	PhoneNum2		VARCHAR2(16),
	PhoneNum3		VARCHAR2(16),
	AbnormalPhone		VARCHAR2(16),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_205_1;
CREATE INDEX I_Callback_Log_205_1 on CallbackLog205 (TransactionNum)
TABLESPACE EMC;




PROMPT *** Callback Log (指令 206)檔 ***
DROP TABLE CallbackLog206 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog206 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	Responding		CHAR(1),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_206_1;
CREATE INDEX I_Callback_Log_206_1 on CallbackLog206 (TransactionNum)
TABLESPACE EMC;




PROMPT *** Callback Log (指令 207)檔 ***
DROP TABLE CallbackLog207 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog207 (
	TransactionNum		CHAR(9),
	CmdType			CHAR(2),
	CreationDate		Date,
	CmdID			CHAR(4),
	SmartCardNum		VARCHAR2(20),
	StbNum			VARCHAR2(20),
	ProcessingFlag		Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_207_1;
CREATE INDEX I_Callback_Log_207_1 on CallbackLog207 (TransactionNum)
TABLESPACE EMC;




PROMPT *** Callback Log (指令 215)檔 ***
DROP TABLE CallbackLog215 CASCADE CONSTRAINT;

CREATE TABLE CallbackLog215 (
	TransactionNum			CHAR(9),
	CmdType				CHAR(2),
	CreationDate			Date,
	CmdID				CHAR(4),
	SmartCardNum			VARCHAR2(20),
	StbNum				VARCHAR2(20),
	OriginalTransactionNum		CHAR(9),	
	IccSuspended			CHAR(1),
	NumOfProducts			CHAR(2),
	ImsProductID			VARCHAR2(12),
	ProductSuspended		CHAR(1),	
	ProcessingFlag			Char(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);


drop index I_Callback_Log_215_1;
CREATE INDEX I_Callback_Log_215_1 on CallbackLog215 (TransactionNum)
TABLESPACE EMC;

commit;
spool off
set heading on



