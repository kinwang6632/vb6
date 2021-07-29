options (errors=200)
load data
infile 'CD057_Map.txt'
append into table CD057
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,RefNo DECIMAL EXTERNAL
,StopFlag DECIMAL EXTERNAL
)
