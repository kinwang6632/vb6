PROMPT *** So119   拆帳作業之頻道歷史拆帳計算參數 ( 目前僅適用於基本公式 ) ***

DROP TABLE So119 CASCADE CONSTRAINT;

CREATE TABLE So119 (
	Comp_ID			VARCHAR2(3),
	BelongYM		VARCHAR2(6),
	Product_ID		VARCHAR2(4),
	Provider_ID             VARCHAR2(10), 
	Provider_Percent        Number(3),   
	So_Percent              Number(3),
	OPERATOR		VARCHAR2(20),
	UPDTIME			VARCHAR2(20),
	CONSTRAINT PK_So119 PRIMARY KEY (Comp_ID,BelongYM,Product_ID)
       );                                                 
