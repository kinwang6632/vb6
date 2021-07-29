options (errors=200)
load data
infile 'CD022A_Map.txt'
append into table CD022A
fields terminated by X'09' TRAILING NULLCOLS
(FaciCode DECIMAL EXTERNAL
,BuyCode DECIMAL EXTERNAL
,CitemCode DECIMAL EXTERNAL
,CitemName CHAR "substrb(rtrim(:CitemName),1,20)"
,PTCode DECIMAL EXTERNAL
,BudgetPlan DECIMAL EXTERNAL
,TotalBudget DECIMAL EXTERNAL
,Period DECIMAL EXTERNAL
,Amount DECIMAL EXTERNAL
,ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)"
,StopFlag DECIMAL EXTERNAL
,TotalAmt  DECIMAL EXTERNAL
)
