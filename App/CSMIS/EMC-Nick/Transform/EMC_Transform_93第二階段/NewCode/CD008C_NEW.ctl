options (errors=200)
load data
infile 'CD008C_NEW.txt'
append into table CD008C
fields terminated by X'09' TRAILING NULLCOLS
(ServiceCode DECIMAL EXTERNAL,
CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
