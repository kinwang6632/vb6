options (errors=200)
load data
infile 'CD008A_Map.txt'
append into table CD008A
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,RefNo DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,StopFlag DECIMAL EXTERNAL
)
