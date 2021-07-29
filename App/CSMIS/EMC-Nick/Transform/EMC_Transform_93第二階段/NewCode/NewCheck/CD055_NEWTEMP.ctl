options (errors=200)
load data
infile 'CD055_NEW.txt'
append into table CD055_NEWTEMP
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,26)"
)
