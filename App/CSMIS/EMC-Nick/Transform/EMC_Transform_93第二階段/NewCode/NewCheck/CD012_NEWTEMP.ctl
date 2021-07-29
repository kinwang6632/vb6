options (errors=200)
load data
infile 'CD012_NEW.txt'
append into table CD012_NEWTEMP
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)"
)
