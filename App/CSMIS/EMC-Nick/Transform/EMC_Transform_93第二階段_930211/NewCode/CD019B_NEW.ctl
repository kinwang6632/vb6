options (errors=200)
load data
infile 'CD019B_NEW.txt'
append into table CD019B
fields terminated by X'09' TRAILING NULLCOLS
(CitemCode DECIMAL EXTERNAL,
CitemName CHAR "substrb(rtrim(:CitemName),1,20)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
PTCode DECIMAL EXTERNAL,
PTName CHAR "substrb(rtrim(:PTName),1,20)",
Amount DECIMAL EXTERNAL,
Period DECIMAL EXTERNAL,
BudgetAmount DECIMAL EXTERNAL,
LastAmount DECIMAL EXTERNAL
)
