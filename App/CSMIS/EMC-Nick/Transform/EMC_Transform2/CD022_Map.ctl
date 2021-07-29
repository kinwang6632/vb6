options (errors=200)
load data
infile 'CD022_Map.txt'
append into table CD022
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,RefNo DECIMAL EXTERNAL
,SnoLength DECIMAL EXTERNAL
,DefBuyCode DECIMAL EXTERNAL
,DefBuyName CHAR "substrb(rtrim(:DefBuyName),1,12)"
,SCardFlag DECIMAL EXTERNAL
,Deposit DECIMAL EXTERNAL
,BudgetPeriod DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,StopFlag DECIMAL EXTERNAL
)
