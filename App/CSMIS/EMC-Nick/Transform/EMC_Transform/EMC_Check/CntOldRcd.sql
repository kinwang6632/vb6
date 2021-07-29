/*
@C:\Transform\EMC_Transform\EMC_Check\CntOldRcd.sql

set serveroutput on
variable nn number
variable msg varchar2(200)
exec :nn := CntOldRcd(:msg);

print nn
print msg
  
  計算待轉換代碼檔"轉檔前"的各個代碼值, 於現有各個資料檔中的資料數目.
  注意: 1. 此程式於正式轉資料程式"前"執行
	2. 做完此程式後, 才可以drop掉資料檔的index
	3. 假設所有代碼對照表命名為: CDxxx_MAP

  1. 計算依據為Code_Map1, 結構應有下列欄位:
	CodeTable	varchar2(30)
	DataTable	varchar2(30)
	DataColumn	varchar2(30)

  2. 計算結果儲存於Code_Map2, 結構應有下列欄位:
	CodeTable	varchar2(30)
	DataTable	varchar2(30)
	DataColumn	varchar2(30)
	OldCodeNo	number(3)
	OldRcdCnt	number(8)
	NewCodeNo	number(3)
	NewRcdCnt	number(8)	

  OUT p_RetMsg VARCHAR2：結果訊息
  Return NUMBER：結果代碼
        0: 正常完畢
	-1: p_RetMsg='參數錯誤'
        -99：p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2002.11.22
	2002.11.25 改: 加上代碼不需轉換的處理
	2002.11.26 改: 加上記錄"新舊"代碼的名稱
	2002.11.27 改: 補強資料檔中有用到的代碼, 但舊代碼中卻無該代碼的處理--
			再至對照表中取舊代碼名稱及新代碼
	2002.12.03 改: 若對照表有新舊代碼都一樣者, 則此程式再加上註明對照表中
			"未提及"新舊都一樣者(新代碼名稱填null).

*/

create or replace function CntOldRcd(p_RetMsg out varchar2) return number
as
  v_SQL1 varchar2(2000);
  v_SQL2 varchar2(2000);
  v_SQL3 varchar2(2000);
  v_CodeNo number;
  v_RcdCnt number;
  v_NewCodeNo number;
  v_OldName varchar2(20);
  v_NewName varchar2(20);

  TYPE CurTyp IS REF CURSOR;  --自訂cursor型態, 供新版dynamic SQL
  v_DyCursor1 CurTyp;
  v_DyCursor2 CurTyp;
  v_DyCursor3 CurTyp;

  cursor cc1 is
	select * from Code_Map1;

begin
  -- 先清除結果檔
  delete from Code_Map2;

  -- 開始計算
  for cr1 in cc1 loop	-- loop Code_Map1 取到的每一筆資料

    -- 檢查資料檔中有用到的代碼, 但舊代碼中卻無該代碼, 
    -- 則再檢查該代碼的對照表中, 是否有加入對照資料, 若有, 則以該對照表的舊名稱及新代碼為主,
    -- 若還找不到, 則表示為異常資料, 
    -- 若有此種代碼, 一樣要計算資料檔中的資料數, 並記錄之==>要請問客戶如何處理?
    v_SQL3 := 'select distinct '||cr1.DataColumn||' from '||cr1.DataTable||' minus (select CodeNo from '||
		cr1.CodeTable||')';
    begin
      OPEN v_DyCursor3 FOR v_SQL3;
    exception
      when others then
	rollback;
	p_RetMsg := 'SQL3錯誤: ' || v_SQL3;
      return -3;
    end;
    loop
      FETCH v_DyCursor3 INTO v_CodeNo;  
      EXIT WHEN v_DyCursor3%NOTFOUND;
      --DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      if v_CodeNo is not null then
        -- 再檢查該代碼的對照表中, 是否有加入對照資料, 若有, 則以該對照表的舊名稱及新代碼為主
        v_SQL2 := 'select OldName, NewCode from '||cr1.CodeTable||'_Map where OldCode='||v_CodeNo;
        begin
	  OPEN v_DyCursor2 FOR v_SQL2;
        exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2錯誤: ' || v_SQL2;
	  return -7;
        end;
        v_NewCodeNo := null;
        v_OldName := null;
        FETCH v_DyCursor2 INTO v_OldName, v_NewCodeNo; 
        CLOSE v_DyCursor2;


	-- 取資料數 v_RcdCnt
	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
	begin
	  OPEN v_DyCursor2 FOR v_SQL2;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2錯誤: ' || v_SQL2;
	  return -5;
	end;
	v_RcdCnt := 0;
	FETCH v_DyCursor2 INTO v_RcdCnt;  
	CLOSE v_DyCursor2;

	-- 新增至結果檔, 此種資料可能無舊代碼名稱OldName ==> 要請問客戶如何處理?
	insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, OldName, NewCodeNo) values
	    (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, v_CodeNo, v_RcdCnt, v_OldName, v_NewCodeNo);
      end if;
    end loop;
    CLOSE v_DyCursor3;

    v_SQL1 := 'select CodeNo, Description from '||cr1.CodeTable||' order by CodeNo';
    begin
      OPEN v_DyCursor1 FOR v_SQL1;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
        return -2;
    end;

    loop       -- loop cr1.CodeTable 的取到的每一筆資料代碼
      v_CodeNo := null;
      FETCH v_DyCursor1 INTO v_CodeNo, v_OldName;
      EXIT WHEN v_DyCursor1%NOTFOUND;
      -- DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      -- 取該CodeTable於DataTable中DataColumn等於該CodeNo的資料數
      if v_CodeNo is not null then
	-- 取對應之新代碼值 v_NewCodeNo, 有可能是null(無對應新代碼)
	-- 假設所有代碼對照表命名為: CDxxx_MAP
	-- 假設所有新代碼暫存檔命名為: CDxxx_NEW
	v_NewCodeNo := null;
	v_NewName := null;
	v_SQL3 := 'select NewCode, NewName from '||cr1.CodeTable||'_Map where OldCode='||v_CodeNo;
	begin
	  OPEN v_DyCursor3 FOR v_SQL3;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL3錯誤: ' || v_SQL3;
	  return -4;
	end;
	FETCH v_DyCursor3 INTO v_NewCodeNo, v_NewName;  
	CLOSE v_DyCursor3;

	-- 取資料數 v_RcdCnt
	v_SQL2 := 'select count(*) from '||cr1.DataTable||' where '||cr1.DataColumn||'='||v_CodeNo;
	begin
	  OPEN v_DyCursor2 FOR v_SQL2;
	exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL2錯誤: ' || v_SQL2;
	  return -6;
	end;

	v_RcdCnt := 0;
	FETCH v_DyCursor2 INTO v_RcdCnt;  

	-- 新增至結果檔
	if v_NewCodeNo is null then	-- 若v_NewCodeNo為空, 則表示該代碼不需轉換
	  v_NewCodeNo := v_CodeNo;
	  v_NewName := null;		-- 2002.12.03 此種於對照表中未列之代碼, 將其新代碼名稱填null, 以便檢查.
	  -- v_NewName := v_OldName;
	end if;
	insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, OldName, NewCodeNo, NewName) values
	    (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, v_CodeNo, v_RcdCnt, v_OldName, v_NewCodeNo, v_NewName);

	CLOSE v_DyCursor2;
      end if;

    end loop;   -- Cursor 迴圈
    CLOSE v_DyCursor1;  -- close cursor

  end loop;

   commit;
  p_RetMsg := '正常完畢';
  return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := '其他錯誤';
    return -99;
end;
/
