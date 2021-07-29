options (errors=200)
load data
infile 'CD032_Map.txt'
append into table CD032
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,StopFlag DECIMAL EXTERNAL
)
