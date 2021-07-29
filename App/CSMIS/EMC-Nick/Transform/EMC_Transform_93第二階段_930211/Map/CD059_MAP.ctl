options (errors=200)
load data
infile 'CD059_MAP.txt'
append into table CD059_MAP
fields terminated by X'09' TRAILING NULLCOLS
(NewCode DECIMAL EXTERNAL,
NewName CHAR "substrb(rtrim(:NewName),1,26)",
OldCode DECIMAL EXTERNAL,
OldName CHAR "substrb(rtrim(:OldName),1,26)",
RefNo DECIMAL EXTERNAL,
Mark DECIMAL EXTERNAL
)
