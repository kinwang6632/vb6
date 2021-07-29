options (errors=200)
load data
infile 'CD008C_Map.txt'
append into table CD008C
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL
,Description CHAR "substrb(rtrim(:Description),1,20)"
,ServiceCode DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
)
