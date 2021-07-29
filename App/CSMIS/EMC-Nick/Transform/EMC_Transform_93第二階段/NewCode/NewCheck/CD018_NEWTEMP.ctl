options (errors=200)
load data
infile 'CD018_NEW.txt'
append into table CD018_NEWTEMP
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,24)"
)
