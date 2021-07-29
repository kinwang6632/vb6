options (errors=200)
load data
infile 'CD009_Map.txt'
append into table CD009_Map
fields terminated by X'09' TRAILING NULLCOLS
(OldCode DECIMAL EXTERNAL
, OldName CHAR "substrb(rtrim(:OldName),1,20)",
NewCode DECIMAL EXTERNAL
, NewName CHAR "substrb(rtrim(:NewName),1,20)"
,Mark DECIMAL EXTERNAL)