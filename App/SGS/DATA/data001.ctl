options (errors=20000)
load data
infile 'TXT\data001.txt'
append into table so009a
fields terminated by X'09' TRAILING NULLCOLS
(CustId  DECIMAL EXTERNAL
, OldSNo CHAR(9) "upper(:OldSNo)"
, OldAddrNo  DECIMAL EXTERNAL "ltrim(:OldAddrNo)"
, OldAddress CHAR(60) "substrb(rtrim(:OldAddress),1,60)"
, PRCode  DECIMAL EXTERNAL "ltrim(:PRCode)"
, PRName CHAR(20) "rtrim(:PRName)"
, ReasonCode  DECIMAL EXTERNAL "ltrim(:ReasonCode)"
, ReasonName CHAR "substrb(rtrim(:ReasonName),1,20)"
, ReInstDate date "YYYYMMDD" "substrb(ltrim(:ReInstDate),1,8)"
, ReInstAddrNo  DECIMAL EXTERNAL "LTRIM(:ReInstAddrNo)"
, ReInstAddress CHAR(60) "rtrim(:ReInstAddress)"
, ResvTime date "YYYYMMDD HH24MISS" "rtrim(ltrim(:ResvTime))"
, FinTime date "YYYYMMDD HH24MISS" "RTRIM(ltrim(:FinTime))"
, AcceptTime date "YYYYMMDD HH24MISS" "rtrim(ltrim(:AcceptTime))"
, WorkerEn1 CHAR(10) "rtrim(:WorkerEn1)"
, WorkerName1 CHAR(20) "rtrim(:WorkerName1)"
, WorkerEn2 CHAR(10) "rtrim(:WorkerEn2)"
, WorkerName2 CHAR(20) "rtrim(:WorkerName2)"
, SignEn CHAR(10) "rtrim(:SignEn)"
, SignName CHAR(20) "rtrim(:SignName)"
, SignDate date "YYYYMMDD" "substrb(ltrim(:SignDate),1,8)"
, AcceptEn CHAR(10) "rtrim(:AcceptEn)"
, AcceptName CHAR(20) "rtrim(:AcceptName)"
, ReturnCode  DECIMAL EXTERNAL "ltrim(:ReturnCode)"
, ReturnName CHAR "substrb(rtrim(:ReturnName),1,20)"
, NewTel1 CHAR(20) "rtrim(:NewTel1)"
, NewChargeAddrNo  DECIMAL EXTERNAL "ltrim(:NewChargeAddrNo)"
, NewChargeAddress CHAR(60) "rtrim(:NewChargeAddress)"
, NewMailAddrNo  DECIMAL EXTERNAL "ltrim(:NewMailAddrNo)"
, NewMailAddress CHAR(60) "rtrim(:NewMailAddress)"
, PrtCount  DECIMAL EXTERNAL "ltrim(:PrtCount)"
, SatiCode  DECIMAL EXTERNAL "ltrim(:SatiCode)"
, SatiName CHAR(20) "rtrim(:SatiName)"
, Note CHAR(750) "rtrim(:Note)"
, ServCode CHAR(3) "ltrim(rtrim(:ServCode))"
, StrtCode  DECIMAL EXTERNAL "ltrim(:StrtCode)"
, NewTel2 CHAR(20) "rtrim(:NewTel2)"
, NewTel3 CHAR(20) "rtrim(:NewTel3)"
, WorkUnit  DECIMAL EXTERNAL "ltrim(:WorkUnit)"
, CompCode  DECIMAL EXTERNAL "translate(:CompCode, :CompCode, '9')"
, ServiceType CHAR "Nvl(ltrim(:ServiceType),'C')")
