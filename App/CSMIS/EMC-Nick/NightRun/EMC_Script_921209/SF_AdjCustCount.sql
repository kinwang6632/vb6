/*

   Night-run程式, 掛在SP_JDailyOp4內: SF_AdjCustCount.sql
   1. 調整大樓統收戶之收費屬性
   2. 調整備份主機收費資料,若是客戶類別不為401(大樓代表戶)但custcount值>1,且截止日加一天>20030101,
	則調整其custcount=1
      註: 不同公司別, 有不同做法, 請參加cursor的定義

  IN	p_CompCode number:	公司別代碼
  return NUMBER: 處理筆數
        >=0: 處理筆數, 表示正常完畢
        -99：p_RetMsg='其他錯誤'

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

 

  -- 公司別: 1, 2, 3, 7, 8, 9, 10, 11, 12 
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

  -- 公司別: 13, 14
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

  -- 公司別: 5, 6
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

  --***** 1. 調整大樓統收戶之收費屬性
  begin
    update so001 set chargetype=3 where mduid in (select mduid from so017 where clctmethod=1) and chargetype<>3;
  exception
    WHEN OTHERS THEN
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      ROLLBACK;  
      return -99;
  end;


  --***** 2. 調整custcount=1
  -- 公司別: 1, 2, 3, 7, 8, 9, 10, 11, 12 
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

  -- 公司別: 13, 14
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

  -- 公司別: 5, 6
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
  DBMS_OUTPUT.PUT_LINE('*** 共處理'||to_char(v_Cnt)||'筆.');
  return v_Cnt;

EXCEPTION

  WHEN OTHERS THEN

    ROLLBACK;
    DBMS_OUTPUT.PUT_LINE('???資料有問題???');
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return -99;
END;

/

