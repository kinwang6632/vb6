options (errors=200)
load data
infile 'CD019A_Map.txt'
append into table CD019A
fields terminated by X'09' TRAILING NULLCOLS
(CitemCode DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,26)"
,CodeNo CHAR "substrb(rtrim(:CodeNo),1,3)"
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
