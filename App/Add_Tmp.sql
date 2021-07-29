/*
@C:\gird\v400\script\Add_Tmp; 
@C:\Gird\v400\Script\�ݴ��{����2\Add_Tmp;

  ADD_TMP.SQL: Create temp tables
  Naming rule: TMP9999

  Date: 2000.07.07

  Modify: 2001.12.15 by Lawrence
          Mark SO032Tmp TempFile
          2002.08.23 �[SO003Z , GiTable , GiColumn , GiIndex
          2002.09.18 ��Tmp003 , Tmp005��������
          2002.10.24 �[Tmp010
          2003.05.23 �[Tmp011 ��Tmp007
          2003.07.23 ��STB�Ǹ������20
          2003.08.18 �[Tmp005 Schema (MediaBillNo, AccountNo)
          2003.08.19 �[Tmp004 Schema (CustId)
          2003.09.26 Tmp005�[ChargeTitle, ServiceType; Tmp004.VouString vc2(250)=>vc2(305)
          2003.10.20 Create Table�ɥ[NOLOGGING�H���Credo log
          2003.11.18 Mark SO003Z delete�Ϊk,�אּTruncate
          2003.11.27 ��gRownum�Ϊk��PK����
          2005.04.22 ���F�h�A�Ȥ@�_�P�ɥX��,Tmp009�[���ServiceType
          2005.08.12 for 1646 Modify Tmp011 Schema OldInstDate
          2005.09.03 �վ�CitemName, Title����
          2005.10.13 �վ�Tmp005.CGName����
          2005.10.25 �վ�Tmp003.CGName����
          2005.12.30 �[Tmp012
          2006.01.18 Tmp012 add InstDate, PRDate
          2006.05.16 Tmp009 add Notes
          2006.05.19 Tmp012 add Plan
          2006.06.06 Adj Address�������[�� for 2408
          2006.12.28 �ߵo��temp(tmp003, tmp004, tmp005)���a�}���[�j��126(��90)
          2007.01.03 PM�S���a�}���n�[�j��142, �w��Casper update���D����󻡩�
          2007.12.06 for 3398
          2008.03.11 for 3785 TMP004��VouString varchar2(357)�զ�VouString varchar2(380)
          2008.07.30 for 4030 TMP009�[���Item,CitemCode,CitemName,FaciSNO
          2008.12.02 for 4142 add Tmp013(�PSO033B���), Tmp014(�PSO032���)
          2008.12.03 for 4142 add TMP015(�ĤG�h�u�f��)
          2009.04.22 add Tmp016(�PSO033B����������)
          2009.05.13 add Tmp017(�PSO032���)
          2009.06.25 add Tmp018(�PSO032���), �B�z���M�ɦ���
          2009.09.16 #5169 TMP009�[���UCCode, STFlag
          2009.12.23 Tmp002.UpdSQL���ץ[�j600->1000

  -- �H�U��Pierre�ҭק�
          2010.06.15 �[�Ȧs��Tmp016Y, ��SF_Accounting.sql, SF_DisLevel20N���ϥ�
          2010.07.14 �NTmp004��VouString�����ץ�varchar2(380)�אּvarchar2(500)
*/

set heading off
variable str varchar2(80)
exec :str := '<< �إ߼Ȧs�� >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool c:\gird\Log\Add_Tmp.log
print str 

prompt TMP001 ���s�p���D��ƼȦs��
drop table TMP001;
create table TMP001 (StrtCode number(6), CustId number(8)) NOLOGGING;

PROMPT Tmp002   �a�}���s�����ɫݳB�z���log�� 
DROP TABLE Tmp002 CASCADE CONSTRAINT;
CREATE TABLE Tmp002 (
	CompCode NUMBER(3),
	AddrNo NUMBER(8),
	UpdSQL varchar2(1000)) NOLOGGING;

PROMPT Tmp003   �o�����ɧ@�~�Ȧs�� for �¶}��
DROP TABLE Tmp003 CASCADE CONSTRAINT;
CREATE TABLE Tmp003 (
	Cust_Id		varchar2(8),
	Invoice_No	varchar2(8),
	Title		varchar2(50),
	Address		varchar2(142),
	Cod_Invoic	char(1),
	Bill_No		varchar2(15),
	Date_c		date,
	C_Real		number(8),
	CGName		varchar2(40),
	Acct_Id		varchar2(7),
	CGCode		char(5),
	Ip_Dat		date,
	Ip_Nam		varchar2(8),
	Tax_Type		char(1),
	Tel_No		number(12),
	Ip_Tim		char(4),
	Zip_No		varchar2(5),
	Charg_addr	varchar2(142),
	Date_Start		date,
	Date_Stop		date,
	Notes		varchar2(40),
	ItemNo          	number(3),
	Cus_Comp		number(3)) NOLOGGING;

prompt TMP004 �o�����ɧ@�~�Ȧs�� for �s�}��
DROP TABLE Tmp004 CASCADE CONSTRAINT;
create table TMP004 (VouString varchar2(500),Sno number(8),CustId number(8),MFlag number(1),
	CONSTRAINT PK_TMP004 PRIMARY KEY (Sno,MFlag)) NOLOGGING;

PROMPT Tmp005   �o�����ɧ@�~�Ȧs�� for �s�}��(�����)
DROP TABLE Tmp005 CASCADE CONSTRAINT;
CREATE TABLE Tmp005 (
	Cust_Id		number(8),
	Invoice_No	varchar2(8),
	Title		varchar2(50),
	Address		varchar2(142),
	Bill_No		varchar2(15),
	Date_c		date,
	C_Real		number(8),
	CGName		varchar2(40),
        	Tax_Type		char(1),
        	ItemNo         	number(3),
     	MediaBillNo	varchar2(11),
        	AccountNo	varchar2(16),
        	ChargeTitle 	varchar2(50),
        	ServiceType 	char(1),
                  MFlag 		number(1)) NOLOGGING;

PROMPT Tmp006   A1����Sub Report
DROP TABLE Tmp006 CASCADE CONSTRAINT;
CREATE TABLE Tmp006 (
        Citemcod        NUMBER(3),
        RealAmt         NUMBER(8),
        Month01         NUMBER(8),
        Month02         NUMBER(8),
        Month03         NUMBER(8),
        Month04         NUMBER(8),
        Month05         NUMBER(8),
        Month06         NUMBER(8),
        Month07         NUMBER(8),
        Month08         NUMBER(8),
        Month09         NUMBER(8),
        Month10         NUMBER(8),
        Month11         NUMBER(8),
        Month12         NUMBER(8),
        MonthElse       NUMBER(8),
        Month20         NUMBER(8)) NOLOGGING;

PROMPT Tmp007   SF_GenPrtSno.Sql��Log��
DROP TABLE Tmp007 CASCADE CONSTRAINT;
CREATE TABLE Tmp007 (
        BillNo          Varchar2(15),
        Item            Number(3),
        CustId          Number(8),
        CustName        Varchar2(50),
        MduId           Varchar2(8),
        OldAddrNo       Number(8),
        OldAddress      Varchar2(90),
        NewAddrNo       Number(8),
        NewAddress      Varchar2(90)) NOLOGGING;

PROMPT Tmp008   SF_ChkDeposit ��Log��
DROP TABLE Tmp008 CASCADE CONSTRAINT;
CREATE TABLE Tmp008 (
        CustId          Number(8),
        CustName        Varchar2(50),
        Tel1            Varchar2(20),
        InstAddress     Varchar2(90),
        CustStatusName  Varchar2(20),
        Amount          Number(8),
        CitemName       Varchar2(40),
        RealDate        Date) NOLOGGING;

/*
PROMPT SO032Tmp   ���So032���X����B
DROP TABLE SO032Tmp CASCADE CONSTRAINT;
CREATE TABLE SO032Tmp (
        CustId     Number(8),
        Flag       Number(1) Default 0,
        UpdTime    Varchar2(19));
*/

PROMPT Tmp009 SF_GenCharge1 ��Log��
DROP TABLE Tmp009 CASCADE CONSTRAINT;
CREATE TABLE Tmp009 (
        CustId     Number(8),
        BillNo     Varchar2(15),
        StartDate  Date,
        StopDate   Date,
        ShouldDate Date,
        ShouldAmt  Number(8),
        ServiceType Char(1),
        Notes      Varchar2(100),
        Item     Number(3),
        CitemCode     Number(5),
        CitemName     Varchar2(40),
        FaciSNO     Varchar2(20),
        UCCode     Number(3),
        STFlag     Number(1) default 0) NOLOGGING;

PROMPT SO003Z���ƥ������
DROP TABLE SO003Z CASCADE CONSTRAINT;
CREATE TABLE SO003Z AS (SELECT * FROM SO003 WHERE CustId=0 and CitemCode=0 and CompCode=0 and SeqNo=0);
ALTER TABLE SO003Z NOLOGGING;
--DELETE FROM SO003Z;
TRUNCATE TABLE SO003Z;
--COMMIT;

PROMPT �إ�GiSchema������Ʈw
DROP TABLE GiTable CASCADE CONSTRAINT;
create table GiTable (TableName varchar2(30), Description varchar2(50), Note varchar2(100), PKColumn varchar2(120), 
		constraint PK_GiTable primary key (TableName));

DROP TABLE GiColumn CASCADE CONSTRAINT;
create table GiColumn (TableName varchar2(30), OrderNo number, Name varchar2(30), Caption varchar2(30), Type varchar2(10), 
		Length varchar2(10), NotNull char(1), KeyType varchar2(4), DefaultStr varchar2(22), CheckStr varchar2(50), 
		constraint PK_GiColumn primary key (TableName, OrderNo));

DROP TABLE GiIndex CASCADE CONSTRAINT;
create table GiIndex (TableName varchar2(30), IndexName varchar2(40), IndexColumn varchar2(4000), PctFreeVal number, InitialVal number, 
		NextExtVal number, MaxExtentsVal number, constraint PK_GiIndex primary key (TableName, IndexName));

PROMPT Tmp010   �a�}���s�����ɿ��~���log�� 
DROP TABLE Tmp010 CASCADE CONSTRAINT;
CREATE TABLE Tmp010 (
	CompCode NUMBER(3),
	AddrNo NUMBER(8),
	UpdSQL varchar2(600)) NOLOGGING;

PROMPT Tmp011   �]�Ƴ̷s���A�@����
DROP TABLE Tmp011 CASCADE CONSTRAINT;
CREATE TABLE Tmp011 (
        FaciSNo varchar2(20),
        BuyName varchar2(20),
        InstDate date,
        PRDate date,
        CustId number,
        SNo varchar2(15),
        PRSNo varchar2(15),
        OldInstDate date) NOLOGGING;

-- 2005.12.30 by Lawrence
PROMPT Tmp012   �򥻸�Ʈֹ�
DROP TABLE Tmp012 CASCADE CONSTRAINT;
CREATE TABLE Tmp012 (
        CustId varchar2(10),
        TFNServiceID varchar2(10),
        CPAreaCode varchar2(3),
        CPNumber varchar2(10),
        TFNStatus varchar2(80),
        ID varchar2(10),
        CustName varchar2(80),
        RecNo number,
        InstDate date,
        PRDate date,
        Plan varchar2(200)) NOLOGGING;

-- 2008.12.02 by Lawrence
PROMPT Tmp013��SO033B�Ȧs�����
DROP TABLE Tmp013 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE  Tmp013 ON COMMIT DELETE ROWS AS (SELECT * FROM SO033B WHERE Rownum=0) ;
ALTER TABLE Tmp013 MODIFY(CANCELFLAG DEFAULT 0);

PROMPT Tmp014��SO032�Ȧs�����
DROP TABLE Tmp014 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE Tmp014 ON COMMIT DELETE ROWS AS (SELECT * FROM SO032 WHERE Rownum=0);
ALTER TABLE Tmp014 MODIFY(Item DEFAULT 1, CancelFlag DEFAULT 0, PrtCount DEFAULT 0, Quantity DEFAULT 0, CustCount DEFAULT 1, 
       ProcessMark DEFAULT 0, ChEven DEFAULT 0, Budget DEFAULT 0, Punish DEFAULT 0,MergePrint DEFAULT 0, ApartFlag DEFAULT 0, 
       AdjustFlag DEFAULT 0, AddFlag DEFAULT 0, PackageStepNo DEFAULT 0, AdjustStat DEFAULT 0);

PROMPT Tmp015   �򥻸�Ʈֹ�
DROP TABLE Tmp015 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE Tmp015 (
        CustId number,
        FaciSeqNo varchar2(15),
        CompCode number,
        ServiceType varchar2(1),
        StartDate date,
        StopDate date,
        SecendDiscount number,
        CitemCode number,
        CitemName varchar2(40),
        ShouldAmt number,
        DAmount number,
        DisLevelCode number,
        Rate number) ON COMMIT DELETE ROWS;

-- 2009.04.22 by Lawrence
PROMPT Tmp016��SO033B���������
DROP TABLE Tmp016 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE Tmp016 ON COMMIT DELETE ROWS AS (SELECT * FROM SO033B WHERE Rownum=0);
ALTER TABLE Tmp016 MODIFY(CANCELFLAG DEFAULT 0);
CREATE INDEX TIDX1 ON Tmp016 (DBillNo,DItem); 
CREATE INDEX TIDX2 ON Tmp016 (BillNo,Item); 
CREATE INDEX TIDX3 ON Tmp016 (CustId); 

PROMPT Tmp017��SO032�Ȧs�����
DROP TABLE Tmp017 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE Tmp017 ON COMMIT DELETE ROWS AS (SELECT * FROM SO032 WHERE Rownum=0);
ALTER TABLE Tmp017 MODIFY(Item DEFAULT 1, CancelFlag DEFAULT 0, PrtCount DEFAULT 0, Quantity DEFAULT 0, CustCount DEFAULT 1, 
       ProcessMark DEFAULT 0, ChEven DEFAULT 0, Budget DEFAULT 0, Punish DEFAULT 0,MergePrint DEFAULT 0, ApartFlag DEFAULT 0, 
       AdjustFlag DEFAULT 0, AddFlag DEFAULT 0, PackageStepNo DEFAULT 0, AdjustStat DEFAULT 0);

PROMPT Tmp018��SO032�Ȧs�����, �B�z���M�ɦ���end of the session
DROP TABLE Tmp018 CASCADE CONSTRAINT;
CREATE GLOBAL TEMPORARY TABLE Tmp018 ON COMMIT PRESERVE ROWS AS (SELECT * FROM SO032 WHERE Rownum=0);
ALTER TABLE Tmp018 MODIFY(Item DEFAULT 1, CancelFlag DEFAULT 0, PrtCount DEFAULT 0, Quantity DEFAULT 0, CustCount DEFAULT 1, 
       ProcessMark DEFAULT 0, ChEven DEFAULT 0, Budget DEFAULT 0, Punish DEFAULT 0,MergePrint DEFAULT 0, ApartFlag DEFAULT 0, 
       AdjustFlag DEFAULT 0, AddFlag DEFAULT 0, PackageStepNo DEFAULT 0, AdjustStat DEFAULT 0);

Prompt Tmp016Y �ĤG�h�u�f�p��ɤ����`���log��, ��SF_Accounting.sql, SF_DisLevel20N���ϥ�
DROP TABLE Tmp016Y CASCADE CONSTRAINT;
CREATE TABLE Tmp016Y (
	ErrorMsg varchar2(200),
	CustId number(8),
	DCitemCode number(5),
	FaciSeqNo varchar2(15),
	BillNo varchar2(15),
	Item number(3),
	ServiceType char(1),
	CitemCode number(5),
	CitemName varchar2(40),
	StartDate1 date,
	StopDate1 date,
	StopDate2 date) NOLOGGING;

spool off
set heading on
