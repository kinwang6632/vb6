options (errors=200)
load data
infile 'CD028_Map.txt'
append into table CD028
fields terminated by X'09' TRAILING NULLCOLS
(CompCode DECIMAL EXTERNAL
,CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,Cost DECIMAL EXTERNAL
,StopFlag DECIMAL EXTERNAL
)
