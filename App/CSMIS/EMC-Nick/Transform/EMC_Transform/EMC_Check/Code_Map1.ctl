options (errors=200)
load data
infile 'Code_Map1.txt'
append into table Code_Map1
fields terminated by X'09' TRAILING NULLCOLS
(CodeTable CHAR "substrb(rtrim(:CodeTable),1,30)"
, DataTable CHAR "substrb(rtrim(:DataTable),1,30)"
, DataColumn CHAR "substrb(rtrim(:DataColumn),1,30)")
