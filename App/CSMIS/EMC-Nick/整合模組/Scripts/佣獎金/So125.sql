PROMPT *** So125  佣金計算之特殊參數檔 ***

DROP TABLE So125 CASCADE CONSTRAINT;

CREATE TABLE So125 (
	CompCode		Number(3),
	Param1			Number(5),
	Param2			Number(5),
	Param3			Number(5),
	Param4			Number(5),
	Param5			Number(5),
	Param6			Number(5),
	Operator		VARCHAR2(20),
	UpdTime		  	VARCHAR2(20),
	Param7			Number(3),
	Param8			Number(3),
	Param9			Number(3)
       );

