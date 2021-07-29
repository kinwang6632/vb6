PROMPT *** SO115   拆帳作業之客戶收視頻道拆帳結果檔  ***

DROP TABLE SO115 CASCADE CONSTRAINT;

CREATE TABLE SO115 (
        Comp_ID                 VARCHAR2(3),
	BelongYM		VARCHAR2(6),
	CustID                  NUMBER(8),
	Product_ID		Varchar2(4),
	Package_ID		Varchar2(3),
	Provider_ID		VARCHAR2(10),
	OrigionalPrice		NUMBER(10),
	RealCharge		NUMBER(10),	
        EMC_Benefit             NUMBER(10),
	So_Benefit              NUMBER(10),
	Provider_Benefit        NUMBER(10),
        OPERATOR                VARCHAR2(20),
        UpdTime                 VARCHAR2(20)
       );

