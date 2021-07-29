options (errors=200)
load data
infile 'CD060_Map.txt'
append into table CD060
fields terminated by X'09' TRAILING NULLCOLS
(CompCode DECIMAL EXTERNAL
,CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,StopFlag DECIMAL EXTERNAL
)
