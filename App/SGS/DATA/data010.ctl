options (errors=200)
load data
infile 'TXT\data010.txt'
append into table so004a
fields terminated by X'09' TRAILING NULLCOLS
(CustId DECIMAL EXTERNAL
, SNo CHAR(15) "upper(:SNo)"
, FaciCode DECIMAL EXTERNAL "ltrim(:FaciCode)"
, FaciSNo CHAR "SUBSTRB(LTRIM(rtrim(:FaciSNo)),1,14)"
, CvtId DECIMAL EXTERNAL "ltrim(:CvtId)"
, BuyCode DECIMAL EXTERNAL "Nvl(RTRIM(ltrim(:BuyCode)),1)"
, DueDate date "YYYYMMDD" "SUBSTRB(RTRIM(:DueDate), 1,8)"
, Amount DECIMAL EXTERNAL "ltrim(:Amount)"
, InstDate date "YYYYMMDD HH24MISS" "substrb(RTRIM(:InstDate), 1,8)"
, InstEn1 CHAR(10) "RTRIM(:InstEn1)"
, InstName1 CHAR(20) "RTRIM(:InstName1)"
, InstEn2 CHAR(10) "RTRIM(:InstEn2)"
, InstName2 CHAR(20) "RTRIM(:InstName2)"
, PRDate date "YYYYMMDD" "substrb(RTRIM(:PRDate), 1,8)"
, PREn1 CHAR(10) "RTRIM(:PREn1)"
, PRName1 CHAR(20) "RTRIM(:PRName1)"
, PREn2 CHAR(10) "RTRIM(:PREn2)"
, PRName2 CHAR(20) "RTRIM(:PRName2)"
-- , MediaCode DECIMAL EXTERNAL "ltrim(:MediaCode)"
, IntroId CHAR(10) "RTRIM(:IntroId)"
, Note CHAR(250) "RTRIM(:Note)"
, ServiceType CHAR "Nvl(LTRIM(:ServiceType),'C')"
, CompCode DECIMAL EXTERNAL "translate(:CompCode, :CompCode, '9')"
, UnitPrice DECIMAL EXTERNAL "ltrim(:UnitPrice)")
