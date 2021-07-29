options (errors=200)
load data
infile 'CD027A_NEW.txt'
append into table CD027A
fields terminated by X'09' TRAILING NULLCOLS
(PackageNo CHAR "substrb(rtrim(:PackageNo),1,3)",
CitemCode DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
CompCode DECIMAL EXTERNAL,
PeriodFlag DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
