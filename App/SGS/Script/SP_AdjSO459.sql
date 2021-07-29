/*
@D:\SGS_930409_Jackal\Script\SP_AdjSO459
set serveroutput on
exec SP_AdjSO459

  Adjust SO459.TotalSubNum. Make it increase 600 each day.

  By: Pierre
  Date: 2004.04.12

*/
create or replace procedure SP_AdjSO459
as
  v_StartExecTime date;		-- Start time
  v_StopExecTime date;		-- Stop time
  v_TtlSecond number;		-- Total execution time (in seconds)
  v_Hour number;		-- Hour part of total execution time
  v_Minute number;		-- Minute part of total execution time
  v_Second number;		-- Second part of total execution time

  v_InitCnt number;		-- SO004中原始客戶數, 不含測試用的300筆ICC所代表的客戶
  v_IncCnt  number;		-- SO004中每日增加的客戶數, 預設為600
  v_Cnt     number;

  cursor cc is 
    select RowId from SO459 where ModeType=1 order by HandleEDateTime;

begin
  v_IncCnt := 600;			-- Increase 600 customers each day
  v_StartExecTime := sysdate;		-- Start time
  DBMS_OUTPUT.PUT_LINE('Start time = '||to_char(v_StartExecTime,'YYYY.MM.DD HH24:MI:SS'));

  -- Get initial customer count
  begin
    select count(distinct CustId) into v_InitCnt from SO004 where FaciCode=202 and substr(FaciSNo,1,6)!='812341'
	and InstDate is not null and PrDate is null;
  exception
    when others then
      v_InitCnt := 0;
  end;

  -- Loop each record of SO459, update TotalSubNum
  v_Cnt := v_InitCnt;
  for cr in cc loop
    v_Cnt := v_Cnt + v_IncCnt;
    update SO459 set TotalSubNum=v_Cnt where RowId=cr.RowId;
  end loop;

  commit;

  v_StopExecTime := sysdate;
  v_TtlSecond := round(86400*(v_StopExecTime-v_StartExecTime));
  v_Hour := trunc(v_TtlSecond/3600);
  v_Second := Mod(v_TtlSecond,60);
  v_Minute := trunc((v_TtlSecond - v_Hour*3600)/60);
  DBMS_OUTPUT.PUT_LINE('Stop time  = '||to_char(v_StopExecTime, 'YYYY.MM.DD HH24:MI:SS'));
  DBMS_OUTPUT.PUT_LINE('*** Total time = '||v_Hour||':'||v_Minute||':'||v_Second);

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
end;
/
