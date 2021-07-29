options (errors=200)
load data
infile 'CD022CD019_NEW.txt'
append into table CD022CD019
fields terminated by X'09' TRAILING NULLCOLS
(FaciCode DECIMAL EXTERNAL,
BuyCode DECIMAL EXTERNAL,
CitemCode DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL
)
