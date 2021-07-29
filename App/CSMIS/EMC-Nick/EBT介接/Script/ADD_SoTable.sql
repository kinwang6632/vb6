
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\EMC\EBT介接\Log
print str 


PROMPT *** SO151  EMC客戶狀態異動寫入資料檔 (以各系統台建立此 table) ***
DROP TABLE SO151 CASCADE CONSTRAINT;
CREATE TABLE SO151 (
	SeqNo	Number(15) Not Null,
	EmcCompCode	Number(3) Not Null,
	EmcCustID	Number(8) Not Null,
	EmcNewCustStatusCode	Number(3) Not Null,
	EmcNewCustStatusDesc	Varchar2(20) Not Null,
	EmcOldCustStatusCode	Number(3) Not Null,
	EmcOldCustStatusDesc	Varchar2(20) Not Null,
	EbtCustID	Varchar2(20) Not Null,
	EbtContractNo	Varchar2(20) Not Null,
	EbtServiceType	Varchar2(2) Not Null,
	ProcessFlag	Varchar2(1) Not Null,
	UpdTime	Date,
	UpdEn	Varchar2(20),
	CONSTRAINT PK_SO151 PRIMARY KEY (SEQNO)
	);
drop index I_SO151_1;
	CREATE INDEX I_SO151_1 ON SO151 (EmcCompCode,EmcCustID,EbtCustID,EbtContractNo,ProcessFlag);

insert into SO151 values(1,7,123,2,'desc2',4,'desc4','ebt1','contract1','CM','W' ,null,'Howard');
insert into SO151 values(2,7,168,9,'desc9',5,'desc5','ebt2','contract2','CT','W' ,null,'Mary');
/*
	EMC與EBT客戶資料對應紀錄檔 => SO152
*/
PROMPT *** SO152   EMC與EBT客戶資料對應紀錄檔 (以各系統台建立此 table)***
DROP TABLE SO152 CASCADE CONSTRAINT;
CREATE TABLE SO152 (
	EmcCompCode	Number(3) Not Null,
	EmcCustID	Number(8) Not Null,
	EbtContractNo		Varchar2(20) Not Null,
	EbtCustID	Varchar2(20),
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100) Not Null,
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbTInstAddr	Varchar2(200) Not Null,
	EbTCustAddr	Varchar2(100),
	EbTBillAddr	Varchar2(200),
	EbtContractStatusCode	Varchar2(10),
	EbtContractStatusDesc	Varchar2(20),
	EbtFeePeriodCode	Varchar2(5),
	EbtFeePeriodDesc	Varchar2(20),
	EbtServiceType	Varchar2(2) Not Null,
	EbtAgentName	Varchar2(20),
	EbtAgentPhone	Varchar2(50),
	EbtAgentID	Varchar2(12),
	EbtAgentAddress	Varchar2(150),
	EbtIdCardId	Varchar2(12),
	EbtCompanyOwnerId	Varchar2(12),
	EbtNonProfitCompanyId	Varchar2(50),
	EbtCompanyId	Varchar2(10),
	IfMoveToOtherSo	Varchar2(1),
	EbtNotes	Varchar2(100),
	UpdTime	Date,
	UpdEn	Varchar2(20)
	);

drop index I_SO152_1;	
	CREATE INDEX I_SO152_1 ON SO152 (EmcCompCode,EmcCustID,EbtCustID,EbtContractNo);

insert into  SO152(EbtCustCName,EbTInstAddr,EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values('Howard','Taipei',7,2,'CM','CMCUS-01020296340');
insert into  SO152(EbtCustCName,EbTInstAddr,EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values('Carol','Taipei City',7,2,'CT','CPCUS-01020296360');
insert into  SO152(EbtCustCName,EbTInstAddr,EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values('Jackal','台北',7,2,'CT','CPCUS-01020296360');


PROMPT *** SO153   待賦予客編之EBT客戶資料檔(以各系統台建立此 table) ***
DROP TABLE SO153 CASCADE CONSTRAINT;
CREATE TABLE SO153 (
	EmcCompCode	Number(3) Not Null,
	EmcCustID	Number(8) ,
	EbtCatvID	Varchar2(6),
	EbtContractNo		Varchar2(20) Not Null,
	EbtCustID	Varchar2(20),
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100) Not Null,
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbTInstAddr	Varchar2(200) Not Null,
	EbTCustAddr	Varchar2(100),
	EbTBillAddr	Varchar2(200),
	EbtContractStatusCode	Varchar2(10),
	EbtContractStatusDesc	Varchar2(20),
	EbtFeePeriodCode	Varchar2(5),
	EbtFeePeriodDesc	Varchar2(20),
	EbtServiceType	Varchar2(2) Not Null,
	EbtAgentName	Varchar2(20),
	EbtAgentPhone	Varchar2(50),
	EbtAgentID	Varchar2(12),
	EbtAgentAddress	Varchar2(150),
	EbtIdCardId	Varchar2(12),
	EbtCompanyOwnerId	Varchar2(12),
	EbtNonProfitCompanyId	Varchar2(50),
	EbtCompanyId	Varchar2(10),
	ProcessFlag	Varchar2(1),
	ErrorCode	Number(5),
	ErrorMsg	Varchar2(50),
	EbtNotes	Varchar2(100),
	EmcNotes	Varchar2(100),
	UpdTime	Date,
	UpdEn	Varchar2(20),
	EbtModifySerialNo	VARCHAR2(30),
	SeqNo	Number(15) Not Null,
	ifSync	Varchar2(1),
	CONSTRAINT PK_SO153 PRIMARY KEY (SEQNO)
	);
drop index I_SO153_1;	
	CREATE INDEX I_SO153_1 ON SO153 (EmcCompCode,EmcCustID,EbtCustID,EbtContractNo,ProcessFlag,ifSync);

drop sequence S_So153_SeqNo;
create sequence S_So153_SeqNo
       MINVALUE 1
       MAXVALUE 9999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       CYCLE;

commit;
spool off
set heading on



