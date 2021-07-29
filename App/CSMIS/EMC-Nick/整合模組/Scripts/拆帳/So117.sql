PROMPT *** SO117   拆帳作業之總和資料  ***

DROP TABLE SO117 CASCADE CONSTRAINT;

CREATE TABLE SO117 (
        Comp_ID                 VARCHAR2(3),
	BelongYM		VARCHAR2(6),
        TotalValidCounts        NUMBER(10),
        TotalCustCounts         NUMBER(10),
	ChannelCounts           VARCHAR2(4000),
        OPERATOR                VARCHAR2(20),
        UpdTime                 VARCHAR2(20),
	CONSTRAINT PK_So117 PRIMARY KEY (Comp_ID,BelongYM)
       );
