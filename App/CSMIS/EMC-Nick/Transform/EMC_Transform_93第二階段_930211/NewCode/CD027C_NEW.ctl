options (errors=200)
load data
infile 'CD027C_NEW.txt'
append into table CD027C
fields terminated by X'09' TRAILING NULLCOLS
(PackageNo CHAR "substrb(rtrim(:PackageNo),1,3)",
CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
