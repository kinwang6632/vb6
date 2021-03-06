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
	CA 高階指令檔 => INV001
*/

PROMPT *** Send_Nagra   CA 高階指令資料檔 ***
DROP TABLE SEND_NAGRA CASCADE CONSTRAINT;

CREATE TABLE SEND_NAGRA (
	HIGH_LEVEL_CMD_ID VARCHAR2(4),
	ICC_NO VARCHAR2(12),
	STB_NO VARCHAR2(16),
	SUBSCRIPTION_BEGIN_DATE DATE,
	SUBSCRIPTION_END_DATE DATE,
	NOTES VARCHAR2(2000),
	CMD_STATUS CHAR(1),
	OPERATOR VARCHAR2(20),
	UPDTIME DATE,
	ERR_CODE VARCHAR2(6),
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
	IRD_CMD_ID VARCHAR2(3),
	IRD_OPERATION VARCHAR2(3),
	IRD_DATA_LENGTH NUMBER(3),
	IRD_DATA VARCHAR2(96),
	SeqNo NUMBER(8),
	CONSTRAINT PK_SEND_NAGRA PRIMARY KEY (SeqNo));


drop index I_Send_Nagra_HIGHLEVELCMDID;
CREATE INDEX I_Send_Nagra_HIGHLEVELCMDID ON Send_Nagra (HIGH_LEVEL_CMD_ID)
TABLESPACE GICMIS_NDX;


commit;
spool off
set heading on



