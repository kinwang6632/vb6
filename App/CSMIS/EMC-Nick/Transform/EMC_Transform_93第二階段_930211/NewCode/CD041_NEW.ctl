options (errors=200)
load data
infile 'CD041_NEW.txt'
append into table CD041
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
UnitCost DECIMAL EXTERNAL,
Type DECIMAL EXTERNAL,
RefNo DECIMAL EXTERNAL,
CompCode DECIMAL EXTERNAL
)
