
CREATE TABLE SO314
(
  CMDSEQNO         NUMBER(8),
  COMPCODE         NUMBER(3),
  MSGCODE          VARCHAR2(20 BYTE),
  MSGNAME          VARCHAR2(90 BYTE),
  OPERATORID       VARCHAR2(10 BYTE),
  UPDTIME          DATE,
  PROCESSINGDATE   DATE,
  CMDSTATUS        CHAR(1 BYTE),
  ERRCODE          VARCHAR2(50 BYTE),
  ERRMSG           VARCHAR2(100 BYTE),
  CUSTID           VARCHAR2(8 BYTE),
  CUSTNAME         VARCHAR2(50 BYTE),
  CUSTTEL          VARCHAR2(20 BYTE),
  CUSTADDRESS      VARCHAR2(90 BYTE),
  ACCOUNTID        VARCHAR2(20 BYTE),
  ACCOUNTPASSWORD  VARCHAR2(24 BYTE),
  SCHEMACODE1      VARCHAR2(3 BYTE),
  SCHEMACODE2      VARCHAR2(3 BYTE),
  USERIDNO         VARCHAR2(20 BYTE),
  USERNAME         VARCHAR2(50 BYTE),
  USERBIRTHDAY     DATE,
  DEVICETYPE       VARCHAR2(1 BYTE),
  DEVICEMODELNO    VARCHAR2(3 BYTE),
  DEVICESNO1       VARCHAR2(20 BYTE),
  DEVICESNO2       VARCHAR2(20 BYTE),
  STOPFLAG         NUMBER(1)                    DEFAULT 0,
  CANCELCODE       NUMBER(5),
  CANCELNAME       VARCHAR2(40 BYTE),
  CANCELDATE       DATE,
  SCMDATA          VARCHAR2(4000 BYTE)
)
TABLESPACE EMCNCC
PCTUSED    0
PCTFREE    10
INITRANS   1
MAXTRANS   255
STORAGE    (
            INITIAL          64K
            MINEXTENTS       1
            MAXEXTENTS       2147483645
            PCTINCREASE      0
            BUFFER_POOL      DEFAULT
           )
LOGGING 
NOCOMPRESS 
NOCACHE
NOPARALLEL
MONITORING;


CREATE UNIQUE INDEX PK_SO314 ON SO314
(CMDSEQNO)
LOGGING
TABLESPACE EMCNCC
PCTFREE    10
INITRANS   2
MAXTRANS   255
STORAGE    (
            INITIAL          64K
            MINEXTENTS       1
            MAXEXTENTS       2147483645
            PCTINCREASE      0
            BUFFER_POOL      DEFAULT
           )
NOPARALLEL;


ALTER TABLE SO314 ADD (
  CONSTRAINT PK_SO314
 PRIMARY KEY
 (CMDSEQNO)
    USING INDEX 
    TABLESPACE EMCNCC
    PCTFREE    10
    INITRANS   2
    MAXTRANS   255
    STORAGE    (
                INITIAL          64K
                MINEXTENTS       1
                MAXEXTENTS       2147483645
                PCTINCREASE      0
               ));



