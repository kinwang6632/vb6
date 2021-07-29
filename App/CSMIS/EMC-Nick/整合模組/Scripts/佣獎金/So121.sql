PROMPT *** So121  佣金計算之員工佣金檔 ***

DROP TABLE So121 CASCADE CONSTRAINT;

CREATE TABLE So121 (
	CompCode	Number(3),
	ComputeYM	Varchar2(5),
	BelongYM	Varchar2(5),
	EmpID		Varchar2(20),
	EmpName		Varchar2(20),
	Commission	Number(10,2),
	MediaCode	Number(3),
	MediaName	Varchar2(20),
	CompType	Char(1),
	ClanCompCode	Number(3),
	ClanCompName	Varchar2(20),
	OPERATOR	VARCHAR2(20),
	UPDTIME		VARCHAR2(20)
       );
