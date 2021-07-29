/*
  -- compile與測試指令
  @c:\emc_script\SP_FillChEvenCode
  set serveroutput on
  variable p_RetCode number
  variable p_RetMsg varchar2(500)
  exec  SP_FillChEvenCode(:p_RetCode,:p_RetMsg);
  print P_Retcode
  print p_RetMsg

  將收費資料的備註欄中頻道結清原因取出, 填入正式的結清原因代碼欄位中.
  檔名: SP_FillEvenCode.sql

  說明: V4.0從92.08月初起於收費資料檔加上頻道結清原因欄位(ChEvenCode, ChEvenName), 頻道結清時會將
	原因填入此二欄位. 但舊版的頻道結清程式會將原因資料填在收費資料的備註欄, 故需寫一支程式將
	舊資料的備註欄中頻道結清原因取出, 填入新欄位中, 以便報表統計分析.

  流程:
  Loop 所有備註有內容的收費資料(SO033, SO034, SO035)
	. 取出頻道結清原因字串: 取"結清原因 : "與", 原項目 :"之間的字串
	. 取道結清原因字串所對應至CD060的代碼
	. 若找到代碼, 則更新該筆收費資料的頻道結清原因欄位
	. 若找不到代碼, 則Error count加1


  OUT	p_RetCode number: 結果代碼
	p_RetMsg VARCHAR2： 結果訊息, 呼叫端變數至少需300 bytes

        >0: 資料批號, 表示正常完畢
	-1: p_RetMsg='參數錯誤'
        -99：p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2002.08.11
*/
create or replace procedure SP_FillChEvenCode(p_RetCode OUT NUMBER, P_RetMsg OUT Varchar2)
is
--  p_RetCode number;
--  p_RetMsg varchar2(500);

  v_StartExecTime	date;
  v_EndExecTime		date;
  c39			char(1) := chr(39);
  v_SQL			varchar2(1000);
  v_ErrorCnt		number := 0;
  v_OkCnt		number := 0;
  v_ChEvenCode		number;
  v_ChEvenName		varchar2(20);

  s1 varchar2(100);
  s2 varchar2(100);
  x1 number;
  x2 number;
  len1 number;
  len2 number;

  cursor cc33 is 
    select RowId, Note from SO033 where Note is not null;
  cursor cc34 is 
    select RowId, Note from SO034 where Note is not null;
  cursor cc35 is 
    select RowId, Note from SO035 where Note is not null;


begin
  v_StartExecTime := sysdate;		-- 開始執行時間

  s1 := '結清原因 : ';
  s2 := ', 原項目 :';
  len1 := lengthb(s1);
  len2 := lengthb(s2);

  -- 處理SO033
  for cr1 in cc33 loop
    -- 取出頻道結清原因字串: 取"結清原因 : "與", 原項目 :"之間的字串
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- 有包含s1字串, 才往下做
      ---** DBMS_OUTPUT.PUT_LINE('x1='||x1);
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      ---** DBMS_OUTPUT.PUT_LINE('x2='||x2);
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
      ---** DBMS_OUTPUT.PUT_LINE('v_ChEvenName=['||v_ChEvenName||']');
    else
      v_ChEvenName := null;
    end if;

    if v_ChEvenName is not null then
      begin
	-- 取道結清原因字串所對應至CD060的代碼
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- 若找到代碼, 則更新該筆收費資料的頻道結清原因欄位
	update SO033 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;


  -- 處理SO034
  for cr1 in cc34 loop
    -- 取出頻道結清原因字串: 取"結清原因 : "與", 原項目 :"之間的字串
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- 有包含s1字串, 才往下做
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
    else
      v_ChEvenName := null;
    end if;
    if v_ChEvenName is not null then
      begin
	-- 取道結清原因字串所對應至CD060的代碼
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- 若找到代碼, 則更新該筆收費資料的頻道結清原因欄位
	update SO034 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;


  -- 處理SO035
  for cr1 in cc35 loop
    -- 取出頻道結清原因字串: 取"結清原因 : "與", 原項目 :"之間的字串
    x1 := instrb(cr1.Note, s1);
    if x1 > 0 then		-- 有包含s1字串, 才往下做
      x2 := instrb(cr1.Note, s2);
    else
      x2 := 0;
    end if;
    if x2 > x1 then
      v_ChEvenName := trim(substrb(cr1.Note, x1+len1, x2-x1-len1));
    else
      v_ChEvenName := null;
    end if;
    if v_ChEvenName is not null then
      begin
	-- 取道結清原因字串所對應至CD060的代碼
        select CodeNo into v_ChEvenCode from CD060 where Description = trim(v_ChEvenName);

	-- 若找到代碼, 則更新該筆收費資料的頻道結清原因欄位
	update SO035 set ChEvenCode=v_ChEvenCode, ChEvenName=v_ChEvenName where RowId=cr1.RowId and ChEvenCode is null;
	if SQL%rowcount > 0 then
	  v_OkCnt := v_OkCnt + 1;
	end if;
      exception
	when others then
	  v_ErrorCnt := v_ErrorCnt + 1;
      end;
    end if;
  end loop;

  v_EndExecTime := sysdate;		-- 結束執行時間
  p_RetMsg := '調整完畢, 共花費'||to_char(trunc(86400*(v_EndExecTime-v_StartExecTime)))||'秒, 調整筆數: '||v_OkCnt;
  p_RetCode := 0;
  commit;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := '其他錯誤: '||SQLERRM;
    p_RetCode := -99;
end;
/