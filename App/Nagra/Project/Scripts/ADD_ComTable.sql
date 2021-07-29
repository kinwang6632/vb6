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
	CA 高階指令檔 => INV001
*/

PROMPT *** Send_Nagra   CA 高階指令資料檔 ***
DROP TABLE SEND_NAGRA CASCADE CONSTRAINT;

CREATE TABLE SEND_NAGRA (
	HIGH_LEVEL_CMD_ID VARCHAR2(4),
	ICC_NO VARCHAR2(20),
	STB_NO VARCHAR2(20),
	SUBSCRIPTION_BEGIN_DATE DATE,
	SUBSCRIPTION_END_DATE DATE,
	NOTES VARCHAR2(2000),
	CMD_STATUS CHAR(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE,
	ERR_CODE VARCHAR2(50),
	ERR_MSG VARCHAR2(50),
	IMS_PRODUCT_ID VARCHAR2(12),
	ZIP_CODE VARCHAR2(5),
	CREDIT_MODE NUMBER(1),
	CREDIT NUMBER(7),
	CREDIT_LIMIT NUMBER(7),
	THRESHOLD_CREDIT NUMBER(7),
	EVENT_NAME VARCHAR2(32),
	PRICE NUMBER(5),
	CC_NUMBER_1 NUMBER(16),
	IP_ADDR VARCHAR2(15),
	CC_PORT NUMBER(5),
	CALLBACK_DATE VARCHAR2(8),
	CALLBACK_TIME VARCHAR2(6),
	CALLBACK_FREQUENCY CHAR(2),
	FIRST_CALLBACK_DATE DATE,
	PHONE_NUM_1 VARCHAR2(16),
	PHONE_NUM_2 VARCHAR2(16),
	PHONE_NUM_3 VARCHAR2(16),
	MIS_IRD_CMD_ID NUMBER(2),
	MIS_IRD_CMD_DATA VARCHAR2(100),
	SeqNo NUMBER(8),
	CompCode	Number(3),
	ReSentTimes	Number(3),
	ProcessingDate	Date
	);

drop index I_Send_Nagra_1;
CREATE INDEX I_Send_Nagra_1 ON Send_Nagra (CMD_STATUS, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_2;
CREATE INDEX I_Send_Nagra_2 ON Send_Nagra (CMD_STATUS,Operator, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_3;
CREATE INDEX I_Send_Nagra_3 ON Send_Nagra (CMD_STATUS,ReSentTimes, Operator, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_4;
CREATE INDEX I_Send_Nagra_4 ON Send_Nagra (CMD_STATUS,ReSentTimes, SeqNo)
TABLESPACE COM;

drop index I_Send_Nagra_5;
CREATE INDEX I_Send_Nagra_5 ON Send_Nagra (SeqNo)
TABLESPACE COM;


PROMPT *** CA sequence number 與 transaction number 的對應檔 ***
DROP TABLE CaSeqNoTransNumMappingData CASCADE CONSTRAINT;

CREATE TABLE CaSeqNoTransNumMappingData (
	SourceID	Varchar2(4),
	SeqNo		Varchar2(9),
	TransactionNum	Varchar2(9),
	UPDTIME DATE
);
drop index I_CaSeqNoTransNumMappingData_1;
CREATE INDEX I_CaSeqNoTransNumMappingData_1 ON CaSeqNoTransNumMappingData (SourceID, TransactionNum)
TABLESPACE COM;


drop index I_CaSeqNoTransNumMappingData_2;
CREATE INDEX I_CaSeqNoTransNumMappingData_2 ON CaSeqNoTransNumMappingData (SourceID, UPDTIME)
TABLESPACE COM;



PROMPT *** CaWareHouseLog   CA 之倉管開關機log檔 ***
DROP TABLE CaWareHouseLog CASCADE CONSTRAINT;

CREATE TABLE CaWareHouseLog (
	ICC_NO VARCHAR2(20),
	STB_NO VARCHAR2(20),
	ProductID VARCHAR2(1024),
	ACTION VARCHAR2(20),
	AuthorStartDate Date,
	AuthorStopDate Date,
	PinCode VARCHAR2(4),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);
drop index I_CaWareHouseLog_1;
CREATE INDEX I_CaWareHouseLog_1 ON CaWareHouseLog (ICC_NO)
TABLESPACE COM;

drop index I_CaWareHouseLog_2;
CREATE INDEX I_CaWareHouseLog_2 ON CaWareHouseLog (STB_NO)
TABLESPACE COM;

drop index I_CaWareHouseLog_3;
CREATE INDEX I_CaWareHouseLog_3 ON CaWareHouseLog (UPDTIME)
TABLESPACE COM;


PROMPT *** CaSendCmdLog   CA send command log檔 ***
DROP TABLE CaSendCmdLog CASCADE CONSTRAINT;

CREATE TABLE CaSendCmdLog (
	HIGH_LEVEL_CMD_ID 	VARCHAR2(4),
	TransactionNum		Varchar2(9),
	FullCmdText		Varchar2(1025),
	Operator		Varchar2(20),
	UPDTIME DATE
);
drop index I_CaSendCmdLog_1;
CREATE INDEX I_CaSendCmdLog_1 ON CaSendCmdLog (TransactionNum,FullCmdText)
TABLESPACE COM;



PROMPT *** CaCallbackData   CA 之callback data log檔 ***
DROP TABLE CaCallbackData CASCADE CONSTRAINT;

CREATE TABLE CaCallbackData (
	TransactionNum	Varchar2(9),
	CallbackData	Varchar2(2000),
	ProcessingState Char(1),
	UPDTIME DATE
);
drop index I_CaCallbackData_1;
CREATE INDEX I_CaCallbackData_1 ON CaCallbackData (TransactionNum)
TABLESPACE COM;


PROMPT *** CaCallbackDataErrorLog   處理 CA 之callback data 時之錯誤 log 檔 ***
DROP TABLE CaCallbackDataErrorLog CASCADE CONSTRAINT;

CREATE TABLE CaCallbackDataErrorLog (
	TransactionNum	Varchar2(9),
	CallbackData	Varchar2(2000),
	ErrorDesc	Varchar2(128),
	UPDTIME DATE
);
drop index I_CaCallbackDataErrorLog_1;
CREATE INDEX I_CaCallbackDataErrorLog_1 ON CaCallbackDataErrorLog (TransactionNum)
TABLESPACE COM;




PROMPT *** CaCmdResponseLog   CA command response log檔 ***
DROP TABLE CaCmdResponseLog CASCADE CONSTRAINT;

CREATE TABLE CaCmdResponseLog (
	NagraTransactionNum	Varchar2(9),	
	TransactionNum		Varchar2(9),
	FullResponseText	Varchar2(1025),
	CmdResponseID		Varchar2(4),
	UPDTIME DATE
);
drop index I_CaCmdResponseLog_1;
CREATE INDEX I_CaCmdResponseLog_1 ON CaCmdResponseLog (TransactionNum)
TABLESPACE COM;




PROMPT *** Send_Nagra_Log CA 高階指令之錯誤資料 log 檔 ***
DROP TABLE SEND_NAGRA_Log CASCADE CONSTRAINT;

CREATE TABLE SEND_NAGRA_Log (
	HIGH_LEVEL_CMD_ID VARCHAR2(4),
	ICC_NO VARCHAR2(20),
	STB_NO VARCHAR2(20),
	SUBSCRIPTION_BEGIN_DATE DATE,
	SUBSCRIPTION_END_DATE DATE,
	NOTES VARCHAR2(2000),
	CMD_STATUS CHAR(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE,
	ERR_CODE VARCHAR2(50),
	ERR_MSG VARCHAR2(50),
	IMS_PRODUCT_ID VARCHAR2(12),
	ZIP_CODE VARCHAR2(5),
	CREDIT_MODE NUMBER(1),
	CREDIT NUMBER(7),
	CREDIT_LIMIT NUMBER(7),
	THRESHOLD_CREDIT NUMBER(7),
	EVENT_NAME VARCHAR2(32),
	PRICE NUMBER(5),
	CC_NUMBER_1 NUMBER(16),
	IP_ADDR VARCHAR2(15),
	CC_PORT NUMBER(5),
	CALLBACK_DATE VARCHAR2(8),
	CALLBACK_TIME VARCHAR2(6),
	CALLBACK_FREQUENCY CHAR(2),
	FIRST_CALLBACK_DATE DATE,
	PHONE_NUM_1 VARCHAR2(16),
	PHONE_NUM_2 VARCHAR2(16),
	PHONE_NUM_3 VARCHAR2(16),
	MIS_IRD_CMD_ID NUMBER(2),
	MIS_IRD_CMD_DATA VARCHAR2(100),
	SeqNo NUMBER(8),
	CompCode	Number(3),
	ReSentTimes	Number(3),
	ProcessingDate	Date,
	LogDate Date);

drop index I_Send_Nagra_Log1;
CREATE INDEX I_Send_Nagra_Log_1 ON Send_Nagra_Log (CMD_STATUS, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_Log_2;
CREATE INDEX I_Send_Nagra_Log_2 ON Send_Nagra_Log (CMD_STATUS,Operator, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_Log_3;
CREATE INDEX I_Send_Nagra_Log_3 ON Send_Nagra_Log (CMD_STATUS,ReSentTimes, Operator, CompCode,ProcessingDate)
TABLESPACE COM;

drop index I_Send_Nagra_Log_4;
CREATE INDEX I_Send_Nagra_Log_4 ON Send_Nagra_Log (CMD_STATUS,ReSentTimes, SeqNo)
TABLESPACE COM;

drop index I_Send_Nagra_Log_5;
CREATE INDEX I_Send_Nagra_Log_5 ON Send_Nagra_Log (SeqNo)
TABLESPACE COM;

drop index I_Send_Nagra_Log_6;
CREATE INDEX I_Send_Nagra_Log_6 ON Send_Nagra_Log (LogDate)
TABLESPACE COM;




commit;
spool off
set heading on



