options (errors=200)
load data
infile 'CD015_NEW.txt'
append into table CD015
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
RefNo DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL,
Assort DECIMAL EXTERNAL,
BackAmt DECIMAL EXTERNAL
)
