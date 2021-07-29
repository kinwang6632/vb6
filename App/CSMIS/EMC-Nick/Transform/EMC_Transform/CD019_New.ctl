options (errors=200)
load data
infile 'CD019_New.txt'
append into table CD019_New
fields terminated by X'09' TRAILING NULLCOLS
(NewCode DECIMAL EXTERNAL
, NewName CHAR "substrb(rtrim(:NewName),1,20)"
,Mark DECIMAL EXTERNAL)
