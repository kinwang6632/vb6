options (errors=200)
load data
infile 'CD019CD004_NEW.txt'
append into table CD019CD004
fields terminated by X'09' TRAILING NULLCOLS
(CitemCode DECIMAL EXTERNAL,
ClassCode DECIMAL EXTERNAL,
Period DECIMAL EXTERNAL,
Amount DECIMAL EXTERNAL,
DayAmt DECIMAL EXTERNAL,
TruncAmt DECIMAL EXTERNAL,
PFlag1 DECIMAL EXTERNAL,
PFlag2 DECIMAL EXTERNAL,
CompCode DECIMAL EXTERNAL
)
