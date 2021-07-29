options (errors=200)
load data
infile 'CD022_NEW.txt'
append into table CD022
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
RefNo DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL,
SCardFlag DECIMAL EXTERNAL,
BudgetPeriod DECIMAL EXTERNAL,
UpperWord DECIMAL EXTERNAL,
CheckDouble DECIMAL EXTERNAL
)
