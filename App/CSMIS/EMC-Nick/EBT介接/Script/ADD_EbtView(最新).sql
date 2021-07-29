prompt
prompt ============================
prompt 以下為 CSIS Web 查詢用 View, 共 14 個系統台
prompt 每個系統為 3 個 view, V_CT_EMCEBT003, V_CT_EMCEBT001, V_CT_EMCEBT002
prompt owner 為 EBTUSER
prompt ============================
prompt


/*
  振道 
*/

prompt
prompt Creating view V_CT_EMCEBT003 for 振道 
prompt ============================
prompt
create or replace view v_ct_emcebt003 as
select CustID, Period from emcct.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_CT_EMCEBT001 for 振道 
prompt ============================
prompt
create or replace view v_ct_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcct.SO001 a , emcct.SO002 b, emcct.SO152 c , emcct.SO014 d, emcct.SO017 e, v_ct_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/


prompt
prompt Creating view V_CT_EMCEBT002 for 振道
prompt ============================
prompt
create or replace view v_ct_emcebt002 as
select TABLECODE, COMPUTEDATE from emcct.SO509A
/



/*
  大安文山 
*/



prompt
prompt Creating view V_DW_EMCEBT003 for 大安文山
prompt ============================
prompt
create or replace view v_dw_emcebt003 as
select CustID, Period from emcdw.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_DW_EMCEBT001 for 大安文山
prompt ============================
prompt
create or replace view v_dw_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcdw.SO001 a , emcdw.SO002 b, emcdw.SO152 c , emcdw.SO014 d, emcdw.SO017 e, v_dw_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_DW_EMCEBT002 for 大安文山
prompt ============================
prompt
create or replace view v_dw_emcebt002 as
select TABLECODE, COMPUTEDATE from emcdw.SO509A
/





/*
  豐盟
*/


prompt
prompt Creating view V_FM_EMCEBT003 for 豐盟
prompt ============================
prompt
create or replace view v_fm_emcebt003 as
select CustID, Period from emcFM.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_FM_EMCEBT001 for 豐盟
prompt ============================
prompt
create or replace view v_fm_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcfm.SO001 a , emcfm.SO002 b, emcfm.SO152 c , emcfm.SO014 d, emcfm.SO017 e, v_fm_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_FM_EMCEBT002 for 豐盟
prompt ============================
prompt
create or replace view v_fm_emcebt002 as
select TABLECODE, COMPUTEDATE from emcFM.SO509A
/






/*
  金頻道
*/

prompt
prompt Creating view V_KP_EMCEBT003 for 金頻道
prompt ============================
prompt
create or replace view v_kp_emcebt003 as
select CustID, Period from emckp.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_KP_EMCEBT001 for 金頻道
prompt ============================
prompt
create or replace view v_kp_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emckp.SO001 a , emckp.SO002 b, emckp.SO152 c , emckp.SO014 d, emckp.SO017 e, v_kp_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_KP_EMCEBT002 for 金頻道
prompt ============================
prompt
create or replace view v_kp_emcebt002 as
select TABLECODE, COMPUTEDATE from emckp.SO509A
/






/*
  觀昇
*/


prompt
prompt Creating view V_KS_EMCEBT003 for 觀昇
prompt ============================
prompt
create or replace view v_ks_emcebt003 as
select CustID, Period from emcks.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_KS_EMCEBT001 for 觀昇
prompt ============================
prompt
create or replace view v_ks_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcks.SO001 a , emcks.SO002 b, emcks.SO152 c , emcks.SO014 d, emcks.SO017 e, v_ks_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_KS_EMCEBT002 for 觀昇
prompt ============================
prompt
create or replace view v_ks_emcebt002 as
select TABLECODE, COMPUTEDATE from emcks.SO509A
/





/*
  新頻道
*/

prompt
prompt Creating view V_NCC_EMCEBT003 for 新頻道
prompt =============================
prompt
create or replace view v_ncc_emcebt003 as
select CustID, Period from emcncc.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_NCC_EMCEBT001 for 新頻道
prompt =============================
prompt
create or replace view v_ncc_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcncc.SO001 a , emcncc.SO002 b, emcncc.SO152 c , emcncc.SO014 d, emcncc.SO017 e, v_ncc_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_NCC_EMCEBT002 for 新頻道
prompt =============================
prompt
create or replace view v_ncc_emcebt002 as
select TABLECODE, COMPUTEDATE from emcncc.SO509A
/





/*
  新台北
*/

prompt
prompt Creating view V_NTP_EMCEBT003 for 新台北
prompt =============================
prompt
create or replace view v_ntp_emcebt003 as
select CustID, Period from emcntp.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_NTP_EMCEBT001 for 新台北
prompt =============================
prompt
create or replace view v_ntp_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcntp.SO001 a , emcntp.SO002 b, emcntp.SO152 c , emcntp.SO014 d, emcntp.SO017 e, v_ntp_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_NTP_EMCEBT002 for 新台北
prompt =============================
prompt
create or replace view v_ntp_emcebt002 as
select TABLECODE, COMPUTEDATE from emcntp.SO509A
/





/*
  北桃園
*/             

prompt
prompt Creating view V_NTY_EMCEBT003 for 北桃園
prompt =============================
prompt
create or replace view v_nty_emcebt003 as
select CustID, Period from emcnty.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_NTY_EMCEBT001 for 北桃園
prompt =============================
prompt
create or replace view v_nty_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcnty.SO001 a , emcnty.SO002 b, emcnty.SO152 c , emcnty.SO014 d, emcnty.SO017 e, v_nty_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_NTY_EMCEBT002 for 北桃園
prompt =============================
prompt
create or replace view v_nty_emcebt002 as
select TABLECODE, COMPUTEDATE from emcnty.SO509A
/






/*
  南天
*/  


prompt
prompt Creating view V_NT_EMCEBT003 for 南天
prompt ============================
prompt
create or replace view v_nt_emcebt003 as
select CustID, Period from emcnt.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_NT_EMCEBT001 for 南天
prompt ============================
prompt
create or replace view v_nt_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcnt.SO001 a , emcnt.SO002 b, emcnt.SO152 c , emcnt.SO014 d, emcnt.SO017 e, v_nt_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_NT_EMCEBT002 for 南天
prompt ============================
prompt
create or replace view v_nt_emcebt002 as
select TABLECODE, COMPUTEDATE from emcnt.SO509A
/



/*
  屏南
*/  


prompt
prompt Creating view V_PN_EMCEBT003 for 屏南
prompt ============================
prompt
create or replace view v_pn_emcebt003 as
select CustID, Period from emcpn.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_PN_EMCEBT001 for 屏南
prompt ============================
prompt
create or replace view v_pn_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcpn.SO001 a , emcpn.SO002 b, emcpn.SO152 c , emcpn.SO014 d, emcpn.SO017 e, v_pn_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_PN_EMCEBT002 for 屏南
prompt ============================
prompt
create or replace view v_pn_emcebt002 as
select TABLECODE, COMPUTEDATE from emcpn.SO509A
/






/*
  新店
*/  


prompt
prompt Creating view V_ST_EMCEBT003 for 新店
prompt ============================
prompt
create or replace view v_st_emcebt003 as
select CustID, Period from emcst.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_ST_EMCEBT001 for 新店
prompt ============================
prompt
create or replace view v_st_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcst.SO001 a , emcst.SO002 b, emcst.SO152 c , emcst.SO014 d, emcst.SO017 e, v_st_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_ST_EMCEBT002 for 新店
prompt ============================
prompt
create or replace view v_st_emcebt002 as
select TABLECODE, COMPUTEDATE from emcst.SO509A
/







/*
  唐城
*/  


prompt
prompt Creating view V_TC_EMCEBT003 for 唐城
prompt ============================
prompt
create or replace view v_tc_emcebt003 as
select CustID, Period from emctc.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_TC_EMCEBT001 for 唐城
prompt ============================
prompt
create or replace view v_tc_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emctc.SO001 a , emctc.SO002 b, emctc.SO152 c , emctc.SO014 d, emctc.SO017 e, v_tc_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_TC_EMCEBT002 for 唐城
prompt ============================
prompt
create or replace view v_tc_emcebt002 as
select TABLECODE, COMPUTEDATE from emctc.SO509A
/





/*
  全聯
*/  

prompt
prompt Creating view V_UC_EMCEBT003 for 全聯
prompt ============================
prompt
create or replace view v_uc_emcebt003 as
select CustID, Period from emcuc.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_UC_EMCEBT001 for 全聯
prompt ============================
prompt
create or replace view v_uc_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcuc.SO001 a , emcuc.SO002 b, emcuc.SO152 c , emcuc.SO014 d, emcuc.SO017 e, v_uc_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_UC_EMCEBT002 for 全聯
prompt ============================
prompt
create or replace view v_uc_emcebt002 as
select TABLECODE, COMPUTEDATE from emcuc.SO509A
/








/*
  陽明山
*/  


prompt
prompt Creating view V_YMS_EMCEBT003 for 陽明山
prompt =============================
prompt
create or replace view v_yms_emcebt003 as
select CustID, Period from emcyms.SO003 where CitemCode=11010
/

prompt
prompt Creating view V_YMS_EMCEBT001 for 陽明山
prompt =============================
prompt
create or replace view v_yms_emcebt001 as
select a.CompCode EmcCompCode, 
       a.CompName EmcCompName, 
       a.CustId EmcCustID, 
       a.CustName EmcCustName,
       b.CustStatusName EmcCustStatusName,
       b.InstTime EmcInstTime,
       b.PrTime EmcPrTime,
       b.wipname1 EmcWipName1,
       b.wipname2 EmcWipName2,
       b.wipname3 EmcWipName3,
       a.Tel1 EmcTel1, 
       a.Tel2 EmcTel2,
       a.Tel3 EmcTel3, 
       a.InstAddress EmcInstAddress,
       a.ClassName1 EmcClassName1, 
       a.ClassName2 EmcClassName2, 
       a.ClassName3 EmcClassName3,
       e.Name EmcMDUName, 
       d.BTName EmcBTName, 
       e.PipelineName EmcPipelineName,
       d.NodeNo EmcNodeNo,
       f.Period EMCPERIOD,
       c.EbtContractNo, 
       c.EbtCustID, 
       c.EbtContractBDate, 
       c.EbtContractEDate, 
       c.EbtCustCName,
       c.EbtCustContactPhone, 
       c.EbtCustContactMobile, 
       c.EbtCompOwnerName,
       c.EbtContactPhone, 
       c.EbtItContactName, 
       c.EbtItContactPhone, 
       c.EbtItContactMobile, 
       c.EbtItEMail,
       c.EbTInstAddr, 
       c.EbTCustAddr, 
       c.EbTBillAddr, 
       c.EbtContractStatusCode,
       c.EbtContractStatusDesc, 
       c.EbtFeePeriodCode, 
       c.EbtFeePeriodDesc,
       c.EbtServiceType, 
       c.EbtAgentName, 
       c.EbtAgentPhone,
       c.EbtAgentID,
       c.EbtAgentAddress,
       c.EbtIdCardId,
       c.EbtCompanyOwnerId,
       c.EbtNonProfitCompanyId, 
       c.EbtCompanyId, 
       c.IfMoveToOtherSo
  from emcyms.SO001 a , emcyms.SO002 b, emcyms.SO152 c , emcyms.SO014 d, emcyms.SO017 e, v_yms_emcebt003 f
 where a.CustId = b.CustId 
   and a.CompCode = b.CompCode 
   and trim( b.ServiceType ) = 'C'
	 and a.CompCode = c.EmcCompCode(+) 
   and a.CustId = c.EmcCustID(+)   
 	 and a.CompCode = d.CompCode
   and a.InstAddrNo = d.AddrNo
   and a.compcode = e.compcode(+)
   and a.mduid = e.mduid(+)
	 and a.CustId = f.CustId(+)
/

prompt
prompt Creating view V_YMS_EMCEBT002 for 陽明山
prompt =============================
prompt
create or replace view v_yms_emcebt002 as
select TABLECODE, COMPUTEDATE from emcyms.SO509A
/

