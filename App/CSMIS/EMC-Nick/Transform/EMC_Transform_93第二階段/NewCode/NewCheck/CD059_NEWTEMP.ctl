options (errors=200)
load data
infile 'CD059_NEW.txt'
append into table CD059_NEWTEMP
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,26)"
)
