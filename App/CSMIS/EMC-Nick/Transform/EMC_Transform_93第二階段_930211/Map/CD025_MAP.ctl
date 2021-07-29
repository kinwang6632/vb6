options (errors=200)
load data
infile 'CD025_MAP.txt'
append into table CD025_MAP
fields terminated by X'09' TRAILING NULLCOLS
(NewCode DECIMAL EXTERNAL,
NewName CHAR "substrb(rtrim(:NewName),1,20)",
OldCode DECIMAL EXTERNAL,
OldName CHAR "substrb(rtrim(:OldName),1,20)",
RefNo DECIMAL EXTERNAL,
Mark DECIMAL EXTERNAL
)
