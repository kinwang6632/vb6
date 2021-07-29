
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\EMC\EBT介接\Log
print str 



/*
	EMC與EBT客戶資料對應紀錄檔 => EMC_EBT001
*/
PROMPT *** EMC_EBT001   EMC與EBT客戶資料對應紀錄檔 (以各系統台建立此 table)***
DROP TABLE EMC_EBT001 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT001 (
	EmcCompCode	Number(3),
	EmcCustID	Number(8),
	EbtContractNo		Varchar2(50),
	EbtCustID	Varchar2(20),
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100),
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbTInstAddr	Varchar2(200),
	EbTCustAddr	Varchar2(100),
	EbTBillAddr	Varchar2(200),
	EbtContractStatusCode	Varchar2(10),
	EbtContractStatusDesc	Varchar2(20),
	EbtFeePeriodCode	Varchar2(5),
	EbtFeePeriodDesc	Varchar2(20),
	EbtServiceType	Char(2),
	UpdTime	Date,
	UpdEn	Varchar2(20)
	);
	
insert into  EMC_EBT001(EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values(7,2,'CM','CMCUS-01020296340');
insert into  EMC_EBT001(EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values(7,2,'CT','CPCUS-01020296360');
insert into  EMC_EBT001(EmcCompCode,EmcCustID,EbtServiceType,EbtContractNo ) values(7,2,'CT','CPCUS-01020296860');

/*
	使用者資料檔 => EMC_EBT002
	STOPFLAG =0 表示不停用
*/
PROMPT *** EMC_EBT002   使用者資料檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT002 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT002 (
	UserID		Varchar2(8),
	UserName	Varchar2(30),
	UserPasswd	Varchar2(10),
	UserGroup	Number(5),
	IsSupervisor	Number(1),
	IsSysOp		Number(1),
	STOPFLAG	Number(1),
	UpdTime		Date,
	UpdEn		Varchar2(20),
	COMPCODE	Number(5),
	DEPTCODE	Number(5),
	TEL		VARCHAR2(20),
	TelExt		Number(5),
	EMail		VARCHAR2(50),
	Creater		VARCHAR2(20),
	CreateDate	Date,
	ifChangePasswd	VARCHAR2(1),
	CONSTRAINT PK_EMCEBT002 PRIMARY KEY (UserID)
	); 
insert into  EMC_EBT002(UserID,UserNam,UserPasswd,UserGroup, IsSupervisor, IsSysOp,STOPFLAG) values('a','Howard','a',1,1,1,0);
-- insert into  EMC_EBT002 values('b','Jackal','b',1,0,0,0);



PROMPT *** EMC_EBT003   使用者群組代碼檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT003 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT003 (
	CODENO		Number(5),
	DESCRIPTION	Varchar2(20),
	SONAME		Varchar2(120),
	SOCODE		Varchar2(50),
	STOPFLAG	Number(1),
	UpdTime		Date,
	UpdEn		Varchar2(20),
	CONSTRAINT PK_EMCEBT003 PRIMARY KEY (CODENO)
	); 
-- insert into  EMC_EBT003 values('1','EMC資訊部','陽明山,振道','9,7',0,null,null);
-- insert into  EMC_EBT003 values('2','EBT客服','振道,觀昇,豐盟','7,1,6',0,null,null);


PROMPT *** EMC_EBT004   使用者登入紀錄檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT004 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT004 (
	UserID		Varchar2(8),
	FromIP		Varchar2(20),
	LoginTime	Date,
	Status		Varchar2(20),
	Note		Varchar2(100)
	); 


PROMPT *** EMC_EBT005   使用者查詢紀錄檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT005 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT005 (
	UserID		Varchar2(8),
	FromIP		Varchar2(20),
	QueryTime	Date,
	SoName		Varchar2(10),
	QueryCond	Varchar2(10),
	QueryKeyValue	Varchar2(500),
	ResultCount	Number(6),
	Note		Varchar2(100)
	); 



PROMPT *** EMC_EBT006   EBT完工寫入資料檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT006 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT006 (
	EmcCustID	Varchar2(11),
	EbtCatvID	Varchar2(6) Not Null,
	EbtContractNo	Varchar2(20) Not Null,
	EbtCustID	Varchar2(20) Not Null,
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100) Not Null,
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName	Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName	Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbtInstAddr	Varchar2(200) Not Null,
	EbtCustAddr	Varchar2(100),
	EbtBillAddr	Varchar2(200),
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
	ProcessFlag	Varchar2(1) Not Null,
	ErrorCode	Number(5),
	ErrorMsg	Varchar2(200),
	EbtNotes	Varchar2(100),
	EmcNotes	Varchar2(100),
	CatvValid	Varchar2(1) Not Null,
	UpdTime	Date,
	UpdEn	Varchar2(20),
	CONSTRAINT PK_EMCEBT006 PRIMARY KEY (EbtContractNo)
	);

drop index I_EMCEBT006_1;
	CREATE INDEX I_EMCEBT006_1 ON EMC_EBT006 (EmcCustID,EbtCustID,EbtContractNo,ProcessFlag);

-- insert into  EMC_EBT006(EbtCustContactPhone, EbtCustID,EbtCustCName,EbtInstAddr,EbtServiceType,EmcCustID,EbtCatvID,EbtContractNo,ProcessFlag,CatvValid) values('27930566','1234','Howard','Taipei','CM','00777665438','040103','CMxxx','W','Y');
-- insert into  EMC_EBT006(EbtCustContactPhone, EbtCustID,EbtCustCName,EbtInstAddr,EbtServiceType,EbtCatvID,EbtContractNo,ProcessFlag,CatvValid) values('3030001','5678','RDP','Taipei','CT','040103','CM888','W','N');
-- insert into  EMC_EBT006(EbtCustContactPhone, EbtCustID,EbtCustCName,EbtInstAddr,EbtServiceType,EbtCatvID,EbtContractNo,ProcessFlag,CatvValid) values('55878196','778899','Carol','台北','CT','040103','CM789','W','N');


PROMPT *** EMC_EBT007   EBT客戶資料異動資料檔  (以 EBTUSER 建立此 table)***
DROP TABLE EMC_EBT007 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT007 (
	EmcCustID	Varchar2(11),
	EbtCatvID	Varchar2(6) Not Null,
	EbtContractNo	Varchar2(20) Not Null,
	EbtModifySerialNo	VARCHAR2(30) Not Null,
	EbtCustID	Varchar2(20) Not Null,
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100) Not Null,
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName	Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName	Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbtInstAddr	Varchar2(200) Not Null,
	EbtCustAddr	Varchar2(100),
	EbtBillAddr	Varchar2(200),
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
	EbtReqId	Varchar2(60) Not Null,
	EbtReqType	Varchar2(128) Not Null,
	OldEmcCustID	Varchar2(11),
	IfMoveToOtherSo	Varchar2(1) Not Null,
	ProcessFlag	Varchar2(1) Not Null,
	ErrorCode	Number(5),
	ErrorMsg	Varchar2(200),
	EbtNotes	Varchar2(100),
	CatvValid	Varchar2(1) ,
	UpdTime	Date,
	UpdEn	Varchar2(20),
	CONSTRAINT PK_EMCEBT007 PRIMARY KEY (EbtContractNo, EbtModifySerialNo)
	);
drop index I_EMCEBT007_1;
	CREATE INDEX I_EMCEBT007_1 ON EMC_EBT007 (EmcCustID,EbtCustID,EbtContractNo,ProcessFlag);


--跨系統台移機, 且沒有 CATV 客編
--insert into  EMC_EBT007(EBTCATVID,EBTCONTRACTNO,EBTMODIFYSERIALNO,EBTCUSTID,EBTCUSTCNAME,EBTINSTADDR,EBTSERVICETYPE,EBTREQID,EBTREQTYPE,IFMOVETOOTHERSO,PROCESSFLAG, EbtCustContactPhone, CatvValid, OldEmcCustID) 
--values('040103','567892','A60','cust19','Name3','taipei','CT','r8','r8desc','Y','W','55878196','N','00788990077');

--跨系統台移機, 且已經有 CATV 客編
--insert into  EMC_EBT007(EmcCustID,EBTCATVID,EBTCONTRACTNO,EBTMODIFYSERIALNO,EBTCUSTID,EBTCUSTCNAME,EBTINSTADDR,EBTSERVICETYPE,EBTREQID,EBTREQTYPE,IFMOVETOOTHERSO,PROCESSFLAG, EbtCustContactPhone, CatvValid, OldEmcCustID, EbtNotes) 
--values('00700000888','040103','567892','A602','cust20','Name3','taipei','CT','r8','r8desc','Y','W','55878196','Y','00788990077','notes168');

--一般異動資料,不是移機資料
--insert into  EMC_EBT007(EmcCustID,EBTCATVID,EBTCONTRACTNO,EBTMODIFYSERIALNO,EBTCUSTID,EBTCUSTCNAME,EBTINSTADDR,EBTSERVICETYPE,EBTREQID,EBTREQTYPE,IFMOVETOOTHERSO,PROCESSFLAG) 
--values('00700997654','040103','567890','A57','cust1','Name1','taipei','CM','r1','r1desc','X','W');

--同一系統台移機
--insert into  EMC_EBT007(EmcCustID,EBTCATVID,EBTCONTRACTNO,EBTMODIFYSERIALNO,EBTCUSTID,EBTCUSTCNAME,EBTINSTADDR,EBTSERVICETYPE,EBTREQID,EBTREQTYPE,IFMOVETOOTHERSO,PROCESSFLAG) 
--values('00700000123','040103','567891','A59','cust18','Name2','taipei','CT','r3','r3desc','N','W');


PROMPT *** EMC_EBT008  EMC客戶狀態異動寫入資料檔 (以 EBTUSER 建立此 table)***
DROP TABLE EMC_EBT008 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT008 (
	EmcCustID	Varchar2(11)  Not Null,
	EmcNewCustStatusCode	Number(3) Not Null,
	EmcNewCustStatusDesc	Varchar2(20) Not Null,
	EmcOldCustStatusCode	Number(3) Not Null,
	EmcOldCustStatusDesc	Varchar2(20) Not Null,
	EbtCustID	Varchar2(20) Not Null,
	EbtContractNo	Varchar2(20) Not Null,
	EbtServiceType	Varchar2(2) Not Null,
	ProcessFlag	Varchar2(1) Not Null,
	UpdTime	Date,
	UpdEn	Varchar2(20)
	);
	CREATE INDEX I_EMCEBT008_1 ON EMC_EBT008 (EmcCustID,EbtCustID,EbtContractNo);





PROMPT *** EMC_EBT009   EBT完工寫入資料 log 檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT009 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT009 (
	EmcCustID	Varchar2(11),
	EbtCatvID	Varchar2(6) Not Null,
	EbtContractNo	Varchar2(20) Not Null,
	EbtCustID	Varchar2(20) Not Null,
	EbtContractBDate	Date,
	EbtContractEDate	Date,
	EbtCustCName	Varchar2(100) Not Null,
	EbtCustContactPhone	Varchar2(20),
	EbtCustContactMobile	Varchar2(20),
	EbtCompOwnerName	Varchar2(50),
	EbtContactPhone	Varchar2(20),
	EbtItContactName	Varchar2(50),
	EbtItContactPhone	Varchar2(20),
	EbtItContactMobile	Varchar2(20),
	EbtItEMail	Varchar2(50),
	EbtInstAddr	Varchar2(200) Not Null,
	EbtCustAddr	Varchar2(100),
	EbtBillAddr	Varchar2(200),
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
	ProcessFlag	Varchar2(1) Not Null,
	ErrorCode	Number(5),
	ErrorMsg	Varchar2(200),
	EbtNotes	Varchar2(100),
	EmcNotes	Varchar2(100),
	CatvValid	Varchar2(1) Not Null,
	UpdTime	Date,
	UpdEn	Varchar2(20),
	CONSTRAINT PK_EMCEBT009 PRIMARY KEY (EbtContractNo)
	);


/*
	公司別資料檔 => EMC_EBT010
	STOPFLAG =0 表示不停用
*/
PROMPT *** EMC_EBT010   公司別資料檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT010 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT010 (
	CODENO		Number(5),
	DESCRIPTION	Varchar2(30),
	STOPFLAG	Number(1),
	UpdTime		Date,
	UpdEn		Varchar2(20),
	CONSTRAINT PK_EMCEBT010 PRIMARY KEY (CODENO)
	); 
-- insert into  EMC_EBT010 values(1,'EMC',0,null,null);
-- insert into  EMC_EBT010 values(2,'EBT',0,null,null);




/*
	部門別資料檔 => EMC_EBT011
	STOPFLAG =0 表示不停用
*/
PROMPT *** EMC_EBT011   部門別資料檔 (以 EBTUSER 建立此 table) ***
DROP TABLE EMC_EBT011 CASCADE CONSTRAINT;
CREATE TABLE EMC_EBT011 (
	CODENO		Number(5),
	DESCRIPTION	Varchar2(30),
	STOPFLAG	Number(1),
	UpdTime		Date,
	UpdEn		Varchar2(20),
	CONSTRAINT PK_EMCEBT011 PRIMARY KEY (CODENO)
	); 
-- insert into  EMC_EBT011 values(1,'資訊部',0,null,null);
-- insert into  EMC_EBT011 values(2,'管理部',0,null,null);
-- insert into  EMC_EBT011 values(3,'客服部',0,null,null);


commit;
spool off
set heading on



