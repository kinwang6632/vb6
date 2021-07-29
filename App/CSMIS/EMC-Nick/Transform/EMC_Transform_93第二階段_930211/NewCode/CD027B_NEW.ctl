options (errors=200)
load data
infile 'CD027B_NEW.txt'
append into table CD027B
fields terminated by X'09' TRAILING NULLCOLS
(PackageNo CHAR "substrb(rtrim(:PackageNo),1,3)",
CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,26)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
FreeGift DECIMAL EXTERNAL
)
