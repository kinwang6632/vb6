options (errors=200)
load data
infile 'CD022A_NEW.txt'
append into table CD022A
fields terminated by X'09' TRAILING NULLCOLS
(FaciCode DECIMAL EXTERNAL,
BuyCode DECIMAL EXTERNAL,
CitemCode DECIMAL EXTERNAL,
PTCode DECIMAL EXTERNAL,
BudgetPlan DECIMAL EXTERNAL,
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL
)
