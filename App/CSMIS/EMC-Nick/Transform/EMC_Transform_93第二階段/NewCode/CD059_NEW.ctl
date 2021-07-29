options (errors=200)
load data
infile 'CD059_NEW.txt'
append into table CD059
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,26)",
CompCode DECIMAL EXTERNAL,
Quantity DECIMAL EXTERNAL,
UnitPrice DECIMAL EXTERNAL,
StopFlag DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
