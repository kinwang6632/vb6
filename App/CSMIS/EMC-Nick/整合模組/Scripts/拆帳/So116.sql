PROMPT *** SO116   ��b�@�~���Ȥ᦬���W�D������  ***

DROP TABLE SO116 CASCADE CONSTRAINT;

CREATE TABLE SO116 (
        Comp_ID                 VARCHAR2(3),
	BelongYM		VARCHAR2(6),
	CustID                  NUMBER(8),
	Channel_String          VARCHAR2(4000), 
        OPERATOR                VARCHAR2(20),
        UpdTime                 VARCHAR2(20),
	CONSTRAINT PK_So116 PRIMARY KEY (Comp_ID,BelongYM,CustID)
       );
