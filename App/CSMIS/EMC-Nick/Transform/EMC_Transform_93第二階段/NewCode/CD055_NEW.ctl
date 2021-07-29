options (errors=200)
load data
infile 'CD055_NEW.txt'
append into table CD055
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,26)",
RefNo DECIMAL EXTERNAL,
StopFlag DECIMAL EXTERNAL
)
