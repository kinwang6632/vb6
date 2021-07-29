options (errors=200)
load data
infile 'CD011A_NEW.txt'
append into table CD011A
fields terminated by X'09' TRAILING NULLCOLS
(MFCode DECIMAL EXTERNAL,
CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
