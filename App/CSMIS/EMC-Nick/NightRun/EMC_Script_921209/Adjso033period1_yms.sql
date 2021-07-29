/* for YMS
重新調整SO033之最後收費其數與金額
檔名: AdjSO033period.sql
PS : 請先將SO003備份一份後再執行本程式
       Ex: Create table so003t as (select * from so003);

@c:\EMC\TRANS\sql\AdjSO033period1_yms
variable nn number
variable msg varchar2(100)
exec :nn := AdjSO033period1(:msg);
print nn
print msg

Date: 2002.07.30   
By: Lawrence

Modify: 2002.09.28 by Lawrence 加條件與參數(請欲用者自行調整)


*/
create or replace function AdjSO033period1(p_RetMsg out varchar2) return number
  as
  x number;
  v_StartExecTime date;
  v_StopExecTime date;
  v_so033date       date;
  v_so033period   number;
  v_so033amount number;
  v_so034date      date;
  v_so034period   number;
  v_so034amount number;
  v_custid             number;
  v_errcnt             number;
  v_cnt                 number;
  v_citemcode      number;
  l_period             number;
  l_amount           number;

  -- 客戶類別代碼請自行調整
  cursor cc1 is
    select rowid,so033.* from SO033 where classcode=600 and uccode=50 and clctym=200211;

  -- 短收原因要過濾refno=1 ; stcode not in (?,?,..)
  cursor cc2(v_Cid number,v_Ccd number) is
    select realdate,realperiod,realamt from SO033 where custid=v_custid and citemcode=v_citemcode and realdate is not null  order by realdate desc;
 --   select realdate,realperiod,realamt from SO033 where custid=v_custid and citemcode=v_citemcode and realdate is not null and nvl(stcode,0)!=1 order by realdate desc;

  -- 短收原因要過濾refno=1 ; stcode not in (?,?,..)
  cursor cc3(v_Cid number,v_Ccd number) is
    select realdate,realperiod,realamt from SO034 where custid=v_custid and citemcode=v_citemcode and realdate is not null  order by realdate desc;
--    select realdate,realperiod,realamt from SO034 where custid=v_custid and citemcode=v_citemcode and realdate is not null and nvl(stcode,0)!=1 order by realdate desc;

BEGIN

  v_StartExecTime := sysdate;		-- 開始執行時間
  v_errcnt:=0;
  v_cnt:=0;
  delete from tmp002;

  for cr1 in cc1 loop

     v_CustId:=cr1.custid;
     v_CitemCode:=cr1.citemcode;
      -- 抓該客戶於so033中之最後一筆資料
      v_so033date:=null;
      v_so033period:=0;
      v_so033amount:=0;
      begin
          open cc2(v_CustId,v_CitemCode);
              fetch cc2 into v_so033date, v_so033period, v_so033amount;
          close cc2;
     exception
         when others then
            v_so033date:=null;
            v_so033period:=0;
            v_so033amount:=0;
     end;
      -- 抓該客戶於so034中之最後一筆資料
      v_so034date:=null;
      v_so034period:=0;
      v_so034amount:=0;
      begin
          open cc3(v_CustId,v_CitemCode);
              fetch cc3 into v_so034date, v_so034period, v_so034amount;
          close cc3;
     exception
         when others then
            v_so034date:=null;
            v_so034period:=0;
            v_so034amount:=0;
     end;
     if v_so033date is null and v_so034date is not null then
        if v_so034period in (1,2,3,4,5) then
           l_period:=3;
           l_amount:=1640;
        elsif v_so034period in (6,7,8,9) then
           l_period:=6;
           l_amount:=3250;
        elsif v_so034period >=10 then
           l_period:=12;
           l_amount:=6500;
        else 
           l_period:=v_so034period;
           l_amount:=v_so034amount;
        end if;
        update so033 set realperiod=l_period,shouldamt=l_amount,realamt=0,realstopdate=add_months(realstartdate,l_period)-1 where rowid=cr1.rowid;
        v_cnt:=v_cnt+1;
    elsif v_so033date is not null and v_so034date is null then
        if v_so033period in (1,2,3,4,5) then
           l_period:=3;
           l_amount:=1640;
        elsif v_so033period in (6,7,8,9) then
           l_period:=6;
           l_amount:=3250;
        elsif v_so033period >=10 then
           l_period:=12;
           l_amount:=6500;
         else 
           l_period:=v_so033period;
           l_amount:=v_so033amount;
        end if;
        update so033 set realperiod=l_period,shouldamt=l_amount,realamt=0,realstopdate=add_months(realstartdate,l_period)-1 where rowid=cr1.rowid;
        v_cnt:=v_cnt+1;
    elsif v_so033date is not null and v_so034date is not null and v_so033date>v_so034date then

        if v_so033period in (1,2,3,4,5) then
           l_period:=3;
           l_amount:=1640;
        elsif v_so033period in (6,7,8,9) then
           l_period:=6;
           l_amount:=3250;
        elsif v_so033period >=10 then
           l_period:=12;
           l_amount:=6500;
         else 
           l_period:=v_so033period;
           l_amount:=v_so033amount;
        end if;
        update so033 set realperiod=l_period,shouldamt=l_amount,realamt=0,realstopdate=add_months(realstartdate,l_period)-1 where rowid=cr1.rowid;
        v_cnt:=v_cnt+1;
    elsif v_so033date is not null and v_so034date is not null and v_so033date<=v_so034date then
        if v_so034period in (1,2,3,4,5) then
           l_period:=3;
           l_amount:=1640;
        elsif v_so034period in (6,7,8,9) then
           l_period:=6;
           l_amount:=3250;
        elsif v_so034period >=10 then
           l_period:=12;
           l_amount:=6500;
         else 
           l_period:=v_so034period;
           l_amount:=v_so034amount;
        end if;
        update so033 set realperiod=l_period,shouldamt=l_amount,realamt=0,realstopdate=add_months(realstartdate,l_period)-1 where rowid=cr1.rowid;
        v_cnt:=v_cnt+1;
    else
        insert into tmp002 (compcode,addrno,UPDSQL) values(cr1.compcode,cr1.custid,to_char(v_so033date,'YYYYMMDD')||'/'||to_char(v_so034date,'YYYYMMDD'));
        v_errcnt:=v_errcnt+1;
    end if;   

  END loop;

  v_StopExecTime := sysdate;
  p_RetMsg := '執行完畢, 共花費'||to_char(trunc(86400*(v_StopExecTime-v_StartExecTime)))||'秒, 共更新:'||to_char(v_cnt)||' 找不到資料:'||to_char(v_errcnt);
--  commit;
  return 0;

EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE(x);
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    DBMS_OUTPUT.PUT_LINE('執行有誤, 共花費'||to_char(trunc(86400*(v_StopExecTime-v_StartExecTime)))||'秒');
    ROLLBACK;
    return -99;
END;
/

