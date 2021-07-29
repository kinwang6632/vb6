options (errors=200)
load data
infile 'CD019_NEW.txt'
append into table CD019
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
RefNo DECIMAL EXTERNAL,
Sign CHAR "substrb(rtrim(:Sign),1,1)",
PeriodFlag DECIMAL EXTERNAL,
Amount DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL,
BackAmount DECIMAL EXTERNAL,
MemberFlag DECIMAL EXTERNAL,
PackageFlag DECIMAL EXTERNAL
)
