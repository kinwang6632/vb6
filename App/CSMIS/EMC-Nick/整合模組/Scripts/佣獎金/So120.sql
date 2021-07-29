PROMPT *** So120  佣金計算之抽佣比例/金額檔 ***

DROP TABLE So120 CASCADE CONSTRAINT;

CREATE TABLE So120 (
	CompCode		Number(3),
	CodeNo			Number(3),
	Description		VARCHAR2(20),
	PromoteCode		Number(3),
	PromoteName		VARCHAR2(20),
	DiscountCode		Number(10),
	DiscountName		VARCHAR2(20),
	PayUnit			Char(1),
	FirstCreditCard1        Number(10,2),
	FirstNotCreditCard1     Number(10,2),
	FirstCreditCard2        Number(10,2),
	FirstNotCreditCard2     Number(10,2),
	OtherCreditCard2        Number(10,2),
	OtherNotCreditCard2     Number(10,2),
	Ref_No			Number(3),
	OPERATOR		VARCHAR2(20),
	UPDTIME		  	VARCHAR2(20),
	ChannelViewDays 		Number(3)
       );
