options (errors=200)
load data
infile 'CD025_NEW.txt'
append into table CD025_NEWTEMP
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)"
)
