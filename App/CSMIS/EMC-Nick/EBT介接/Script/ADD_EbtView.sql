set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\EMC\EBT介接\Log
print str 

prompt V_CT_EMCEBT003: SO003 資料檔(以 EBTUSER 建立此 view)
drop view V_CT_EMCEBT003;
create view V_CT_EMCEBT003 as
  select CustID, Period from emcct.SO003 where CitemCode=301;


prompt V_CT_EMCEBT001: web查詢資料檔(以 EBTUSER 建立此 view)
-- 以下給大新店,新唐城以外的 SO 用...
drop view V_CT_EMCEBT001;
create view V_CT_EMCEBT001 as
  select a.CompCode EmcCompCode, a.CompName EmcCompName, a.CustId EmcCustID, a.CustName EmcCustName,
	b.CustStatusName EmcCustStatusName,b.InstTime EmcInstTime,b.PrTime EmcPrTime,
	a.Tel1 EmcTel1, a.Tel2 EmcTel2,a.Tel3 EmcTel3, a.InstAddress EmcInstAddress, 
	a.ClassName1 EmcClassName1, a.ClassName2 EmcClassName2, a.ClassName3 EmcClassName3,
	d.MDUName EmcMDUName, d.BTName EmcBTName, d.PipelineName EmcPipelineName,
	d.NodeNo EmcNodeNo,e.Period EMCPERIOD,
	c.EbtContractNo, c.EbtCustID, c.EbtContractBDate, c.EbtContractEDate, c.EbtCustCName,
	c.EbtCustContactPhone, c.EbtCustContactMobile, c.EbtCompOwnerName,
	c.EbtContactPhone, c.EbtItContactName, c.EbtItContactPhone, c.EbtItContactMobile, c.EbtItEMail,
	c.EbTInstAddr, c.EbTCustAddr, c.EbTBillAddr, c.EbtContractStatusCode,
	c.EbtContractStatusDesc, c.EbtFeePeriodCode, c.EbtFeePeriodDesc,
	c.EbtServiceType,c.EbtAgentName, c.EbtAgentPhone,c.EbtAgentID,c.EbtAgentAddress,
	c.EbtIdCardId,c.EbtCompanyOwnerId,
	c.EbtNonProfitCompanyId,c.EbtCompanyId, c.IfMoveToOtherSo 
    from emcct.SO001 a , emcct.SO002 b, emcct.SO152 c , emcct.SO014 d, V_CT_EMCEBT003 e 
    where a.CustId=b.CustId and a.CompCode=b.CompCode and b.ServiceType='C' 
	and a.CompCode=d.CompCode and a.InstAddrNo=d.AddrNo 
	and a.CustId=e.CustId(+)
	and a.CompCode=c.EmcCompCode(+) and a.CustId=c.EmcCustID(+);


-- 以下給大新店以及新唐城的 SO 用...
drop view V_TC_EMCEBT001;
create view V_TC_EMCEBT001 as
  select a.CompCode EmcCompCode, a.CompName EmcCompName, a.CustId EmcCustID, a.CustName EmcCustName,
	b.CustStatusName EmcCustStatusName,b.InstTime EmcInstTime,b.PrTime EmcPrTime,
	a.Tel1 EmcTel1, a.Tel2 EmcTel2,a.Tel3 EmcTel3, a.InstAddress EmcInstAddress, 
	a.ClassName1 EmcClassName1, a.ClassName2 EmcClassName2, a.ClassName3 EmcClassName3,
	d.MDUName EmcMDUName, d.BTName EmcBTName, d.PipelineName EmcPipelineName,
	d.NodeNo EmcNodeNo,e.Period EMCPERIOD,
	c.EbtContractNo, c.EbtCustID, c.EbtContractBDate, c.EbtContractEDate, c.EbtCustCName,
	c.EbtCustContactPhone, c.EbtCustContactMobile, c.EbtCompOwnerName,
	c.EbtContactPhone, c.EbtItContactName, c.EbtItContactPhone, c.EbtItContactMobile, c.EbtItEMail,
	c.EbTInstAddr, c.EbTCustAddr, c.EbTBillAddr, c.EbtContractStatusCode,
	c.EbtContractStatusDesc, c.EbtFeePeriodCode, c.EbtFeePeriodDesc,
	c.EbtServiceType,c.EbtAgentName, c.EbtAgentPhone,c.EbtAgentID,c.EbtAgentAddress,
	c.EbtIdCardId,c.EbtCompanyOwnerId,
	c.EbtNonProfitCompanyId,c.EbtCompanyId, c.IfMoveToOtherSo 
    from emctc.SO001 a , emctc.SO002 b, emctc.SO152 c , emctc.SO014 d, V_TC_EMCEBT003 e 
    where a.CustId=b.CustId and a.CompCode=b.CompCode and b.ServiceType like 'C%' 
	and a.CompCode=d.CompCode and a.InstAddrNo=d.AddrNo 
	and a.CustId=e.CustId(+)
	and a.CompCode=c.EmcCompCode(+) and a.CustId=c.EmcCustID(+);

/* 用法: select * from V_CT_EMCEBT001*/




prompt V_CT_EMCEBT002: snapshot 資料檔(以 EBTUSER 建立此 view)
drop view V_CT_EMCEBT002;
create view V_CT_EMCEBT002 as
  select TABLECODE, COMPUTEDATE from emcct.SO509A;

/* 用法: select * from V_CT_EMCEBT001*/





spool off
set heading on



