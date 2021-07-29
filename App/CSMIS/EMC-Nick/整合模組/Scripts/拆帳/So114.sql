PROMPT *** So114   拆帳作業之拆帳計算結果檔 ***

DROP TABLE So114 CASCADE CONSTRAINT;

CREATE TABLE So114 (
	Comp_ID			VARCHAR2(3),
	BelongYM		VARCHAR2(6),
	Begin_Date		Date,
	End_Date		Date,
	EMC_Income		Number(10),
	Income			Number(10),
	Outcome			Number(10),
	EMC_Benefit		Number(10),
	So_Benefit		Number(10),
	ProductID		Varchar2(4),
	ProviderID              Varchar2(10),    
	ProviderDesc            Varchar2(50),
	Provider_Benefit	VARCHAR2(250),
        SingleCounts            Number(10),
        PackageCounts           Number(10),
	FirstCount		Number(10),
	LastCount		Number(10),
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20),
	AddCount		Number(10),
	ReduceCount		Number(10),
	CONSTRAINT PK_So114 PRIMARY KEY (Comp_ID,BelongYM,ProductID)
       );
