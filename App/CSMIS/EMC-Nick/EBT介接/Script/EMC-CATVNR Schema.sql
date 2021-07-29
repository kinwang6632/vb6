-------------------------------------------------
-- Export file for user EBTUSER                --
-- Created by OSX on 2004/07/27, ¤U¤È 03:41:58 --
-------------------------------------------------

prompt
prompt Creating table EMC_EBT002
prompt =========================
prompt
create table EMC_EBT002
(
  USERID         VARCHAR2(16) not null,
  USERNAME       VARCHAR2(60),
  USERPASSWD     VARCHAR2(20),
  USERGROUP      NUMBER(5),
  ISSUPERVISOR   NUMBER(1),
  ISSYSOP        NUMBER(1),
  UPDTIME        DATE,
  UPDEN          VARCHAR2(40),
  STOPFLAG       NUMBER(1),
  COMPCODE       NUMBER(5),
  DEPTCODE       NUMBER(5),
  TEL            VARCHAR2(40),
  TELEXT         NUMBER(5),
  EMAIL          VARCHAR2(100),
  CREATER        VARCHAR2(40),
  CREATEDATE     DATE,
  IFCHANGEPASSWD VARCHAR2(2)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT002
  add constraint PK_EMCEBT002 primary key (USERID)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT003
prompt =========================
prompt
create table EMC_EBT003
(
  CODENO      NUMBER(5) not null,
  DESCRIPTION VARCHAR2(120),
  STOPFLAG    NUMBER(1),
  UPDTIME     DATE,
  UPDEN       VARCHAR2(40),
  SONAME      VARCHAR2(240),
  SOCODE      VARCHAR2(100)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT003
  add constraint PK_EMCEBT003 primary key (CODENO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT004
prompt =========================
prompt
create table EMC_EBT004
(
  USERID    VARCHAR2(16),
  FROMIP    VARCHAR2(40),
  LOGINTIME DATE,
  STATUS    VARCHAR2(40),
  NOTE      VARCHAR2(200)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT005
prompt =========================
prompt
create table EMC_EBT005
(
  USERID        VARCHAR2(16),
  FROMIP        VARCHAR2(40),
  QUERYTIME     DATE,
  SONAME        VARCHAR2(20),
  QUERYCOND     VARCHAR2(100),
  QUERYKEYVALUE VARCHAR2(1000),
  RESULTCOUNT   NUMBER(6),
  NOTE          VARCHAR2(200)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT006
prompt =========================
prompt
create table EMC_EBT006
(
  EMCCUSTID             VARCHAR2(22),
  EBTCATVID             VARCHAR2(12) not null,
  EBTCONTRACTNO         VARCHAR2(40) not null,
  EBTCUSTID             VARCHAR2(40) not null,
  EBTCONTRACTBDATE      DATE,
  EBTCONTRACTEDATE      DATE,
  EBTCUSTCNAME          VARCHAR2(200) not null,
  EBTCUSTCONTACTPHONE   VARCHAR2(40),
  EBTCUSTCONTACTMOBILE  VARCHAR2(40),
  EBTCOMPOWNERNAME      VARCHAR2(100),
  EBTCONTACTPHONE       VARCHAR2(40),
  EBTITCONTACTNAME      VARCHAR2(100),
  EBTITCONTACTPHONE     VARCHAR2(40),
  EBTITCONTACTMOBILE    VARCHAR2(40),
  EBTITEMAIL            VARCHAR2(100),
  EBTINSTADDR           VARCHAR2(400) not null,
  EBTCUSTADDR           VARCHAR2(200),
  EBTBILLADDR           VARCHAR2(400),
  EBTCONTRACTSTATUSCODE VARCHAR2(20),
  EBTCONTRACTSTATUSDESC VARCHAR2(40),
  EBTFEEPERIODCODE      VARCHAR2(10),
  EBTFEEPERIODDESC      VARCHAR2(40),
  EBTSERVICETYPE        VARCHAR2(4) not null,
  EBTAGENTNAME          VARCHAR2(40),
  EBTAGENTPHONE         VARCHAR2(100),
  EBTAGENTID            VARCHAR2(24),
  EBTAGENTADDRESS       VARCHAR2(300),
  EBTIDCARDID           VARCHAR2(24),
  EBTCOMPANYOWNERID     VARCHAR2(24),
  EBTNONPROFITCOMPANYID VARCHAR2(100),
  EBTCOMPANYID          VARCHAR2(20),
  PROCESSFLAG           VARCHAR2(2) not null,
  ERRORCODE             NUMBER(5),
  ERRORMSG              VARCHAR2(400),
  EBTNOTES              VARCHAR2(200),
  EMCNOTES              VARCHAR2(200),
  CATVVALID             VARCHAR2(2) not null,
  UPDTIME               DATE,
  UPDEN                 VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT006
  add constraint PK_EMCEBT006 primary key (EBTCONTRACTNO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
create index I_EMCEBT006_1 on EMC_EBT006 (EMCCUSTID,EBTCUSTID,EBTCONTRACTNO,PROCESSFLAG)
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
grant select, insert, update on EMC_EBT006 to EBTD01;
grant select on EMC_EBT006 to EBTD02;

prompt
prompt Creating table EMC_EBT007
prompt =========================
prompt
create table EMC_EBT007
(
  EMCCUSTID             VARCHAR2(22),
  EBTCATVID             VARCHAR2(12) not null,
  EBTCONTRACTNO         VARCHAR2(40) not null,
  EBTMODIFYSERIALNO     VARCHAR2(60) not null,
  EBTCUSTID             VARCHAR2(40) not null,
  EBTCONTRACTBDATE      DATE,
  EBTCONTRACTEDATE      DATE,
  EBTCUSTCNAME          VARCHAR2(200) not null,
  EBTCUSTCONTACTPHONE   VARCHAR2(40),
  EBTCUSTCONTACTMOBILE  VARCHAR2(40),
  EBTCOMPOWNERNAME      VARCHAR2(100),
  EBTCONTACTPHONE       VARCHAR2(40),
  EBTITCONTACTNAME      VARCHAR2(100),
  EBTITCONTACTPHONE     VARCHAR2(40),
  EBTITCONTACTMOBILE    VARCHAR2(40),
  EBTITEMAIL            VARCHAR2(100),
  EBTINSTADDR           VARCHAR2(400) not null,
  EBTCUSTADDR           VARCHAR2(200),
  EBTBILLADDR           VARCHAR2(400),
  EBTCONTRACTSTATUSCODE VARCHAR2(20),
  EBTCONTRACTSTATUSDESC VARCHAR2(40),
  EBTFEEPERIODCODE      VARCHAR2(10),
  EBTFEEPERIODDESC      VARCHAR2(40),
  EBTSERVICETYPE        VARCHAR2(4) not null,
  EBTAGENTNAME          VARCHAR2(40),
  EBTAGENTPHONE         VARCHAR2(100),
  EBTAGENTID            VARCHAR2(24),
  EBTAGENTADDRESS       VARCHAR2(300),
  EBTIDCARDID           VARCHAR2(24),
  EBTCOMPANYOWNERID     VARCHAR2(24),
  EBTNONPROFITCOMPANYID VARCHAR2(100),
  EBTCOMPANYID          VARCHAR2(20),
  EBTREQID              VARCHAR2(120) not null,
  EBTREQTYPE            VARCHAR2(256) not null,
  OLDEMCCUSTID          VARCHAR2(22),
  IFMOVETOOTHERSO       VARCHAR2(2) not null,
  PROCESSFLAG           VARCHAR2(2) not null,
  ERRORCODE             NUMBER(5),
  ERRORMSG              VARCHAR2(400),
  EBTNOTES              VARCHAR2(200),
  CATVVALID             VARCHAR2(2),
  UPDTIME               DATE,
  UPDEN                 VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT007
  add constraint PK_EMCEBT007 primary key (EBTCONTRACTNO,EBTMODIFYSERIALNO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
create index I_EMCEBT007_1 on EMC_EBT007 (EMCCUSTID,EBTCUSTID,EBTCONTRACTNO,PROCESSFLAG)
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
grant select, insert, update on EMC_EBT007 to EBTD01;

prompt
prompt Creating table EMC_EBT008
prompt =========================
prompt
create table EMC_EBT008
(
  EMCCUSTID            VARCHAR2(22) not null,
  EMCNEWCUSTSTATUSCODE NUMBER(3) not null,
  EMCNEWCUSTSTATUSDESC VARCHAR2(40) not null,
  EMCOLDCUSTSTATUSCODE NUMBER(3) not null,
  EMCOLDCUSTSTATUSDESC VARCHAR2(40) not null,
  EBTCUSTID            VARCHAR2(40) not null,
  EBTCONTRACTNO        VARCHAR2(40) not null,
  EBTSERVICETYPE       VARCHAR2(4) not null,
  PROCESSFLAG          VARCHAR2(2) not null,
  UPDTIME              DATE,
  UPDEN                VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
create index I_EMCEBT008_1 on EMC_EBT008 (EMCCUSTID,EBTCUSTID,EBTCONTRACTNO)
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
grant select, update on EMC_EBT008 to EBTD01;
grant select on EMC_EBT008 to EBTD02;

prompt
prompt Creating table EMC_EBT009
prompt =========================
prompt
create table EMC_EBT009
(
  EMCCUSTID             VARCHAR2(22),
  EBTCATVID             VARCHAR2(12) not null,
  EBTCONTRACTNO         VARCHAR2(40) not null,
  EBTCUSTID             VARCHAR2(40) not null,
  EBTCONTRACTBDATE      DATE,
  EBTCONTRACTEDATE      DATE,
  EBTCUSTCNAME          VARCHAR2(200) not null,
  EBTCUSTCONTACTPHONE   VARCHAR2(40),
  EBTCUSTCONTACTMOBILE  VARCHAR2(40),
  EBTCOMPOWNERNAME      VARCHAR2(100),
  EBTCONTACTPHONE       VARCHAR2(40),
  EBTITCONTACTNAME      VARCHAR2(100),
  EBTITCONTACTPHONE     VARCHAR2(40),
  EBTITCONTACTMOBILE    VARCHAR2(40),
  EBTITEMAIL            VARCHAR2(100),
  EBTINSTADDR           VARCHAR2(400) not null,
  EBTCUSTADDR           VARCHAR2(200),
  EBTBILLADDR           VARCHAR2(400),
  EBTCONTRACTSTATUSCODE VARCHAR2(20),
  EBTCONTRACTSTATUSDESC VARCHAR2(40),
  EBTFEEPERIODCODE      VARCHAR2(10),
  EBTFEEPERIODDESC      VARCHAR2(40),
  EBTSERVICETYPE        VARCHAR2(4) not null,
  EBTAGENTNAME          VARCHAR2(40),
  EBTAGENTPHONE         VARCHAR2(100),
  EBTAGENTID            VARCHAR2(24),
  EBTAGENTADDRESS       VARCHAR2(300),
  EBTIDCARDID           VARCHAR2(24),
  EBTCOMPANYOWNERID     VARCHAR2(24),
  EBTNONPROFITCOMPANYID VARCHAR2(100),
  EBTCOMPANYID          VARCHAR2(20),
  PROCESSFLAG           VARCHAR2(2) not null,
  ERRORCODE             NUMBER(5),
  ERRORMSG              VARCHAR2(400),
  EBTNOTES              VARCHAR2(200),
  EMCNOTES              VARCHAR2(200),
  CATVVALID             VARCHAR2(2) not null,
  UPDTIME               DATE,
  UPDEN                 VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT009
  add constraint PK_EMCEBT009 primary key (EBTCONTRACTNO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT010
prompt =========================
prompt
create table EMC_EBT010
(
  CODENO      NUMBER(5) not null,
  DESCRIPTION VARCHAR2(60),
  STOPFLAG    NUMBER(1),
  UPDTIME     DATE,
  UPDEN       VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT010
  add constraint PK_EMCEBT010 primary key (CODENO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table EMC_EBT011
prompt =========================
prompt
create table EMC_EBT011
(
  CODENO      NUMBER(5) not null,
  DESCRIPTION VARCHAR2(60),
  STOPFLAG    NUMBER(1),
  UPDTIME     DATE,
  UPDEN       VARCHAR2(40)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
alter table EMC_EBT011
  add constraint PK_EMCEBT011 primary key (CODENO)
  using index 
  tablespace REPORT
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table TEST3
prompt ====================
prompt
create table TEST3
(
  DESCRIPTION VARCHAR2(6)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table TEST4
prompt ====================
prompt
create table TEST4
(
  DESCRIPTION VARCHAR2(6)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table TM1
prompt ==================
prompt
create table TM1
(
  USERID        VARCHAR2(16),
  FROMIP        VARCHAR2(40),
  QUERYTIME     DATE,
  SONAME        VARCHAR2(20),
  QUERYCOND     VARCHAR2(20),
  QUERYKEYVALUE VARCHAR2(1000),
  RESULTCOUNT   NUMBER(6),
  NOTE          VARCHAR2(200)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );

prompt
prompt Creating table TT3
prompt ==================
prompt
create table TT3
(
  TT VARCHAR2(6)
)
tablespace REPORT
  pctfree 10
  pctused 40
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
