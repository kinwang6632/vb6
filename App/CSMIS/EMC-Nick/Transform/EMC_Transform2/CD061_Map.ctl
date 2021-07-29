options (errors=200)
load data
infile 'CD061_Map.txt'
append into table CD061
fields terminated by X'09' TRAILING NULLCOLS
(CompCode DECIMAL EXTERNAL
,CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,StopFlag DECIMAL EXTERNAL
)
