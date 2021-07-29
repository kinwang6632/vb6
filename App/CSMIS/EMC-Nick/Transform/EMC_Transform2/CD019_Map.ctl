options (errors=200)
load data
infile 'CD019_Map.txt'
append into table CD019
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
, Description CHAR "substrb(rtrim(:Description),1,20)"
,Amount DECIMAL EXTERNAL
,RefNo DECIMAL EXTERNAL
,sign CHAR "substrb(rtrim(:sign),1,1)"
,PeriodFlag DECIMAL EXTERNAL
,StopFlag DECIMAL EXTERNAL
,BackAmount DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,Period DECIMAL EXTERNAL
)
