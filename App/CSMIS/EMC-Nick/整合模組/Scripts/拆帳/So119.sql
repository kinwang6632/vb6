PROMPT *** So119   ��b�@�~���W�D���v��b�p��Ѽ� ( �ثe�ȾA�Ω�򥻤��� ) ***

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
