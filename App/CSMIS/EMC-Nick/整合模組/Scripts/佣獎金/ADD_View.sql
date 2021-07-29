/*
  ADD_VIEW.SQL: Create views
  Naming rule: V_xxxxxx
  Date: 2003.05.30
*/
set heading off
variable str varchar2(80)
exec :str := '<< �إ߼Ȧs�� >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\Invoice\Log 
print str 

prompt V_User: �ϥΪ̸��
drop view V_User;
create view V_User as
  select IdentifyId1, IdentifyId2, UserId, Password, 
	 UserName, GroupID, DbConnSeq, UptTime, UptEn
    from INV004;

/* �Ϊk: select * from V_User*/


prompt V_MasterCustomer: mis�Ȥ�D���
drop view V_MasterCustomer;
create view V_MasterCustomer as
  select CustId, CustName from So001;

/* �Ϊk: select * from V_MasterCustomer*/

prompt V_DetailCustomer: mis�Ȥ���Ӹ��
drop view V_DetailCustomer;
create view V_DetailCustomer as
  select CompCode, CustId, ServiceType, 
	InvTitle, InvNo, InvAddress from So002;

/* �Ϊk: select * from V_DetailCustomer */

prompt V_MDCustomer: mis�Ȥ�D/���Ӹ��
drop view V_MDCustomer;
create view V_MDCustomer as
select   a.CUSTNAME, a.CompCode, a.CustId, a.MailAddress,
	 c.ServiceType, c.InvTitle, c.InvNo BUSINESSID, c.InvAddress,
	 b.ZipCode
         from So001 a, So014 b , So002 c
	 where a.MailAddrNo=b.AddrNo and a.CustID=c.CustID;
      

/* �Ϊk: select * from V_MDCustomer */
/*
prompt V_CommChargeData: mis�Ȥ�ú�O���
drop view V_CommChargeData;
create view V_CommChargeData as
  select 33 Source, a.CompCode, a.CustId, a.BillNo, a.Item,
	a.CitemCode, a.CitemName, a.ShouldDate, 
	a.ShouldAmt,a.RealDate,a.RealAmt,a.RealPeriod,
	a.RealStartDate,a.RealStopDate,a.ClctName,a.STName,
	c.Rate1,b.TaxCode  
 	from so033 a, cd019 b, cd033 c
	where a.CitemCode=b.CodeNo and b.TaxCode=c.CodeNo
	and (b.TaxCode is not null or c.TaxFlag=1)
	and a.CancelFlag=0
	and (a.GUINo is null or a.GUINo='') and a.InvoiceTime is null
  union
  select 34 Source, a.CompCode, a.CustId, a.BillNo, a.Item,
	a.CitemCode, a.CitemName, a.ShouldDate, 
	a.ShouldAmt,a.RealDate,a.RealAmt,a.RealPeriod,
	a.RealStartDate,a.RealStopDate,a.ClctName,a.STName,
	c.Rate1,b.TaxCode   
 	from so034 a, cd019 b, cd033 c
	where a.CitemCode=b.CodeNo and b.TaxCode=c.CodeNo
	and (b.TaxCode is not null or c.TaxFlag=1)
	and a.CancelFlag=0
	and (a.GUINo is null or a.GUINo='') and a.InvoiceTime is null;
*/

prompt V_ChargeInfo: mis�Ȥ�ú�O���
drop view V_ChargeInfo;
create view V_ChargeInfo as
  select 33 Source, a.CompCode, a.CustId, a.BillNo, a.Item,
	a.CitemCode, a.CitemName, a.ShouldDate, 
	a.ShouldAmt,a.RealDate,a.RealAmt,a.RealPeriod,
	a.RealStartDate,a.RealStopDate,a.ClctName,a.STName,
	c.Rate1, b.TaxCode   
 	from so033 a, cd019 b, cd033 c
	where a.CitemCode=b.CodeNo and b.TaxCode=c.CodeNo
	and (b.TaxCode is not null or c.TaxFlag=1)
	and a.CancelFlag=0
	and (a.GUINo is null or a.GUINo='') and a.InvoiceTime is null;

/* �Ϊk: select * from V_TranDateInfo */


prompt V_TranDateInfo: ���b������
drop view V_TranDateInfo;
create view V_TranDateInfo as
select   TYPE, TRANDATE, COMPCODE, SERVICETYPE from so062;
      

/* �Ϊk: select * from V_TranDataInfo */


spool off
set heading on



