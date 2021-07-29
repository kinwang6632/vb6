/*
@D:\App\EMC\Snapshot\SP_RefreshGroupPackage

EXEC SP_RefreshGroupPackage
PRINT ReturnNo
PRINT RetMsg


  By: Howard
  Date: 2003.06.19
*/

CREATE OR REPLACE PROCEDURE SP_RefreshGroupPackage
  AS
  v_RefreshGroup varchar2(100);
  v_ErrorMsg varchar2(4000);
  v_date1 date;
  v_date2 date;
  v_RetMsg varchar2(4000);
  v_TotalSecs varchar2(20,5);

begin
  v_date1 := sysdate;
  v_ErrorMsg := '';
  v_RetMsg := '';


  begin
    v_RefreshGroup := 'EMCYMS.YMS1';
    DBMS_REFRESH.REFRESH('EMCYMS.YMS1');
  exception
    when others then
      v_ErrorMsg := v_ErrorMsg || ' refresh ' || v_RefreshGroup || ' 時發生錯誤.';
  end;





  begin
    v_RefreshGroup := 'EMCYMS.YMS2';
    DBMS_REFRESH.REFRESH('EMCYMS.YMS2');
  exception
    when others then
      v_ErrorMsg := v_ErrorMsg || ' refresh ' || v_RefreshGroup || ' 時發生錯誤.';
  end;



  begin
    v_RefreshGroup := 'EMCYMS.YMS3';
    DBMS_REFRESH.REFRESH('EMCYMS.YMS3');
  exception
    when others then
      v_ErrorMsg := v_ErrorMsg || ' refresh ' || v_RefreshGroup || ' 時發生錯誤.';
  end;

  begin
    v_RefreshGroup := 'EMCYMS.YMS4';
    DBMS_REFRESH.REFRESH('EMCYMS.YMS4');
  exception
    when others then
      v_ErrorMsg := v_ErrorMsg || ' refresh ' || v_RefreshGroup || ' 時發生錯誤.';
  end;

  v_date2 := sysdate;

  v_TotalSecs := to_Char(trunc((v_date2 - v_date1)*24*60*60));

  if (v_ErrorMsg='')  or (v_ErrorMsg Is NULL )then
    v_RetMsg := '順利執行完畢' ;
  else
    v_RetMsg := v_ErrorMsg; 
  end if;



  v_RetMsg := v_RetMsg || '--總執行時間: ' || v_TotalSecs || ' 秒';

  insert into SO141(EXECBEGINTIME,EXECENDTIME, RESULT) values(v_date1, v_date2,  v_RetMsg);
  commit;




exception
  when others then
    DBMS_OUTPUT.PUT_LINE(v_RetMsg);  

end;
/