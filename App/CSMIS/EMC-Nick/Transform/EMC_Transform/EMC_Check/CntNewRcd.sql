/*
@C:\Transform\EMC_Transform\EMC_Check\CntNewRcd.sql

set serveroutput on
variable nn number
variable msg varchar2(200)
exec :nn := CntNewRcd(:msg);

print nn
print msg
  
  計算轉換代碼檔"轉檔後"的各個代碼值, 於各個調整後資料檔中的資料數目.
  注意: 1. 此程式於正式轉資料程式"後"執行
	2. 執行此程式前, 也要先建立各資料檔的index

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
  Date: 2002.11.22 初稿
	2002.11.25 加上對未轉換代碼的計算
	2002.11.26 改: 加上記錄"新"代碼的名稱 (取消, 改在轉檔前就統計了)

*/

create or replace function CntNewRcd(p_RetMsg out varchar2) return number
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

  -- 開始計算
  for cr1 in cc1 loop	-- loop Code_Map1 取到的每一筆資料
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
      v_NewName := null;

      FETCH v_DyCursor1 INTO v_CodeNo, v_NewName;
      EXIT WHEN v_DyCursor1%NOTFOUND;
      -- DBMS_OUTPUT.PUT_LINE(cr1.CodeTable||', '||cr1.DataTable||', '||cr1.DataColumn||', '||v_CodeNo);

      -- 取該CodeTable於DataTable中DataColumn等於該CodeNo的資料數
      if v_CodeNo is not null then
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
	-- 更新至結果檔
	update Code_Map2 set NewRcdCnt=nvl(NewRcdCnt,0)+v_RcdCnt where CodeTable=cr1.CodeTable and 
		DataTable=cr1.DataTable and DataColumn=cr1.DataColumn and NewCodeNo=v_CodeNo;

	-- 若找不到, 則可能是新舊代碼一樣, 故更新之
	if SQL%ROWCOUNT = 0 then
	  update Code_Map2 set NewRcdCnt=nvl(NewRcdCnt,0)+v_RcdCnt, NewCodeNo=v_CodeNo, NewName=v_NewName
		where CodeTable=cr1.CodeTable and DataTable=cr1.DataTable and DataColumn=cr1.DataColumn 
		and OldCodeNo=v_CodeNo and (NewCodeNo=v_CodeNo or NewCodeNo is null);
	  if SQL%ROWCOUNT = 0 then	-- 若還找不到, 則再新增一筆 (不應發生此種情形)
	    insert into Code_Map2 (CodeTable, DataTable, DataColumn, OldCodeNo, OldRcdCnt, NewCodeNo, NewRcdCnt, NewName) 
	      values (cr1.CodeTable, cr1.DataTable, cr1.DataColumn, null, 0, v_CodeNo, v_RcdCnt, v_NewName);
	  end if;
	end if;
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
