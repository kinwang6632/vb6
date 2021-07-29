options (errors=200)
load data
infile 'CD051_NEW.txt'
append into table CD051
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
RefNo DECIMAL EXTERNAL,
CompCode DECIMAL EXTERNAL,
StopFlag DECIMAL EXTERNAL
)
