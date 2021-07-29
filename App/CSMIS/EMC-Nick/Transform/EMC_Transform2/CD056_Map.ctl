options (errors=200)
load data
infile 'CD056_Map.txt'
append into table CD056
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,26)"
,StopFlag DECIMAL EXTERNAL
)
