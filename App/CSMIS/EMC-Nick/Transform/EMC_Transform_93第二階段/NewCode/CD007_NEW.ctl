options (errors=200)
load data
infile 'CD007_NEW.txt'
append into table CD007
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo DECIMAL EXTERNAL,
Description CHAR "substrb(rtrim(:Description),1,20)",
RefNo DECIMAL EXTERNAL,
CitemCode1 DECIMAL EXTERNAL,
CitemName1 CHAR "substrb(rtrim(:CitemName1),1,20)",
CitemCode2 DECIMAL EXTERNAL,
CitemName2 CHAR "substrb(rtrim(:CitemName2),1,20)",
CitemCode3 DECIMAL EXTERNAL,
CitemName3 CHAR "substrb(rtrim(:CitemName3),1,20)",
FacilCode1 DECIMAL EXTERNAL,
FacilName1 CHAR "substrb(rtrim(:FacilName1),1,20)",
FacilCode2 DECIMAL EXTERNAL,
FacilName2 CHAR "substrb(rtrim(:FacilName2),1,20)",
FacilCode3 DECIMAL EXTERNAL,
FacilName3 CHAR "substrb(rtrim(:FacilName3),1,20)",
PrintWONow DECIMAL EXTERNAL,
GroupNo DECIMAL EXTERNAL,
GroupName CHAR "substrb(rtrim(:GroupName),1,20)",
ServiceType CHAR "substrb(rtrim(:ServiceType),1,1)",
StopFlag DECIMAL EXTERNAL,
FacilSelected DECIMAL EXTERNAL,
ReserveDay DECIMAL EXTERNAL
)
