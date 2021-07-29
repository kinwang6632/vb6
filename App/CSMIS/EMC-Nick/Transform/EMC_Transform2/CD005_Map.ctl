options (errors=200)
load data
infile 'CD005_Map.txt'
append into table CD005
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,RefNo DECIMAL EXTERNAL
,GroupNo DECIMAL EXTERNAL
,GroupName CHAR "substrb(rtrim(:GroupName),1,20)"
,PrintWONow DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,StopFlag DECIMAL EXTERNAL
)
