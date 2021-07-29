/*
  ADD_Table.SQL: Create Tables
  Naming rule: NDSXXX
  Date: 2001.07.18
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\APP\Start-TV\Log
print str 




PROMPT *** SEND_NDS   CA 指令檔 ***
DROP TABLE SEND_NDS CASCADE CONSTRAINT;
CREATE TABLE SEND_NDS (
	SeqNO Number(8), 						
	CompCode Number(3),
	CompName  VARCHAR2(20), 					
	CommandID VARCHAR2(4),
	SubscriberID VARCHAR2(8),
	IccNo VARCHAR2(20),
	SubBeginDate Date,
	SubEndDate Date,
	PinCode Number(4),
	PopulationID NUMBER(5), 
	RegionKey Varchar2(8),
	ReportbackAvailability Varchar2(1),
	ReportbackDate Number(2),
	Notes Varchar2(4000),
	CmdStatus Varchar2(1),
	ErrorCode Varchar2(5),
	ErrorDesc Varchar2(30),
	Operator VARCHAR2(20), 						
	ResentTimes Number(2),
	ProcessingDate Date,
	UpdTime DATE,
	PinControl Varchar2(1),
	ServiceID Varchar2(5),
	ResponseData Varchar2(1024)
	);
	CREATE INDEX I_SEND_NDS_1 ON SEND_NDS (CommandID)
	TABLESPACE COM_NDX;
	CREATE INDEX I_SEND_NDS_2 ON SEND_NDS (SeqNo)
	TABLESPACE COM_NDX;


PROMPT *** NDS002   CA 之倉管開關機log檔 ***
DROP TABLE NDS002 CASCADE CONSTRAINT;

CREATE TABLE NDS002 (
	ICC_NO VARCHAR2(12),
	Subscriber_ID VARCHAR2(16),
	ProductID VARCHAR2(1024),
	ACTION VARCHAR2(20),
	AuthorStartDate Date,
	AuthorStopDate Date,
	PinCode VARCHAR2(4),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE
);
drop index I_NDS002_1;
CREATE INDEX I_NDS002_1 ON NDS002 (Subscriber_ID)
TABLESPACE COM_NDX;

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
CREATE INDEX I_CaSendCmdLog_1 ON CaSendCmdLog (TransactionNum)
TABLESPACE COM_NDX;


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
TABLESPACE COM_NDX;




commit;
spool off
set heading on






