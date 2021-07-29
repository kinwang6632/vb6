/*

   Night-run�{��, ���bSP_JDailyOp4��: SF_AdjCustCount.sql
   1. �վ�j�ӲΦ��ᤧ���O�ݩ�
   2. �վ�ƥ��D�����O���,�Y�O�Ȥ����O����401(�j�ӥN���)��custcount��>1,�B�I���[�@��>20030101,
	�h�վ��custcount=1
      ��: ���P���q�O, �����P���k, �аѥ[cursor���w�q

  IN	p_CompCode number:	���q�O�N�X
  return NUMBER: �B�z����
        >=0: �B�z����, ��ܥ��`����
        -99�Gp_RetMsg='��L���~'

   select a.custid,a.billno,a.custcount,a.citemcode,a.classcode,b.classcode1,a.mduid,b.mduid
     from so033 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
     a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
     a.custid=b.custid(+) and b.classcode1<>401 and  a.citemcode=301 

   @C:\EMC_Script\SF_AdjCustCount.sql

   set serveroutput on
   variable nn number
   exec :nn := SF_AdjCustCount(7);
   print nn

   Date: 2003.03.31
*/
create or replace function SF_AdjCustCount(p_CompCode number) return number
as
  v_Cnt number;

 

  -- ���q�O: 1, 2, 3, 7, 8, 9, 10, 11, 12 
  cursor cTable1 is
                  
    select 1 Type, a.custid, a.billno,a.citemcode
      from so033 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and b.classcode1<>401 and a.citemcode=301 
    union
    select 2 Type, a.custid, a.billno,a.citemcode
      from so034 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and b.classcode1<>401 and a.citemcode=301 ;

  -- ���q�O: 13, 14
  cursor cTable2 is
                  
    select 1 Type, a.custid, a.billno,a.citemcode
      from so033 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and b.classcode1<>410 and a.citemcode=301 
    union
    select 2 Type, a.custid, a.billno,a.citemcode
      from so034 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and b.classcode1<>410 and a.citemcode=301 ;

  -- ���q�O: 5, 6
  cursor cTable3 is
                  
    select 1 Type, a.custid, a.billno,a.citemcode 
      from so033 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and  
      a.citemcode=301 and a.custid not in (select maincustid from so017 where clctmethod=1)
    union
    select 2 Type, a.custid, a.billno,a.citemcode 
      from so034 a,so001 b,so002 c where a.realstopdate+1>=to_date('20030101','yyyymmdd') and 
      a.custid=c.custid(+) and c.custstatuscode=1 and a.custcount>1 and 
      a.custid=b.custid(+) and  
      a.citemcode=301 and a.custid not in (select maincustid from so017 where clctmethod=1);


BEGIN

  /**********************************************/

  v_Cnt := 0;

  /**********************************************/

  --***** 1. �վ�j�ӲΦ��ᤧ���O�ݩ�
  begin
    update so001 set chargetype=3 where mduid in (select mduid from so017 where clctmethod=1) and chargetype<>3;
  exception
    WHEN OTHERS THEN
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      ROLLBACK;  
      return -99;
  end;


  --***** 2. �վ�custcount=1
  -- ���q�O: 1, 2, 3, 7, 8, 9, 10, 11, 12 
  if p_CompCode in (1, 2, 3, 7, 8, 9, 10, 11, 12) then
    for cr1 in cTable1 loop
      if cr1.Type=1 then 
        update so033 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      else
        update so034 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      end if;
      v_Cnt := v_Cnt+1;
    end loop;
  end if;

  -- ���q�O: 13, 14
  if p_CompCode in (13, 14) then
    for cr1 in cTable2 loop
      if cr1.Type=1 then 
        update so033 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      else
        update so034 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      end if;
      v_Cnt := v_Cnt+1;
    end loop;
  end if;

  -- ���q�O: 5, 6
  if p_CompCode in (5, 6) then
    for cr1 in cTable3 loop
      if cr1.Type=1 then 
        update so033 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      else
        update so034 set custcount=1 where custid=cr1.custid and billno=cr1.billno and citemcode=cr1.citemcode;
      end if;
      v_Cnt := v_Cnt+1;
    end loop;
  end if;


  --  commit;
  DBMS_OUTPUT.PUT_LINE('*** �@�B�z'||to_char(v_Cnt)||'��.');
  return v_Cnt;

EXCEPTION

  WHEN OTHERS THEN

    ROLLBACK;
    DBMS_OUTPUT.PUT_LINE('???��Ʀ����D???');
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return -99;
END;

/

