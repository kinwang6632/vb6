/*
@C:\EMC_Script\SF_AdjPTCode

set serveroutput on
variable nn number
variable msg varchar2(1000)

exec :nn := SF_AdjPTCode(:msg);
print nn
print msg

  根據PTCode!=1的每戶最近單據資料(其PTCode可能為3或4), 若該戶於SO003中的PTCode為空值, 
  則調整該戶於SO003中的PTCode.
  檔名: SF_AdjPTCode.sql

  每一組修改至SO003的會有log至SO080, 各欄位如下: (執行時, 會先清除SO080的資料)
	. RcdNo:	資料序號
	. CustId:	客編
	. Type:		收費項目代碼(CitemCode)
	. Content:	存修改的PTCode值, 格式='PTCode=???'
	. CompCode:	公司別
	. ServiceType:	服務別

  OUT	p_RetMsg VARCHAR2： 結果訊息, 呼叫端變數至少需100 bytes

  return NUMBER: 結果代碼
        =0: 表示正常完畢
	-1: p_RetMsg='參數錯誤'
        -2: p_RetMsg='取 Sequence S_SeqNo1 的次一可用值失敗'
        -3: p_RetMsg='SQL錯誤: ' || v_SQL
        -4: p_RetMsg='新增資料失敗, 訊息='||SQLERRM;
        -99: p_RetMsg='其他錯誤'

  By: Pierre
  Date: 2003.06.24
	2003.06.25 改: 加log至SO080
*/
create or replace function SF_AdjPTCode(p_RetMsg out varchar2) return number
as
  c39		char(1) := chr(39);
  v_StartExecTime date;
  v_PTName 	varchar2(20);
  v_RcdCnt	number := 0;
  v_UpdCnt	number := 0;

  v_CompCode	number;
  v_ServiceType char(1);

  -- 註: 只改SO003中同客編/同收費項目只有一筆者
  cursor cc1 is
    select custid,citemcode,ptcode, count(*) from (
	SELECT custid,citemcode,ptcode FROM SO033 
	where nvl(ptcode,1)!=1
	union
	SELECT custid,citemcode,ptcode FROM SO034 
	where nvl(ptcode,1)!=1 )
	group by CustId, CitemCode, PtCode
	having count(*)=1;

begin
  -- 參數檢查

  v_StartExecTime := sysdate;   -- 開始執行時間

  begin
    select CompCode, ServiceType into v_CompCode, v_ServiceType from SO041 where RowNum<=1;
  exception
    when others then
      p_RetMsg := '取公司別/服務別錯誤: '||SQLERRM;
      return -2;
  end;

  -- 清除log檔
  delete from SO080;
  commit;

  for cr1 in cc1 loop
    begin
      v_RcdCnt := v_RcdCnt + 1;
      begin
	select Description into v_PTName from CD032 where CodeNo=cr1.PtCode;
      exception 
	when others then
	  v_PTName := null;
      end;

      update SO003 set PtCode=cr1.PtCode, PtName=v_PTName where CustId=cr1.CustId and 
	CitemCode=cr1.CitemCode and PtCode is null;
      v_UpdCnt := v_UpdCnt + SQL%ROWCOUNT;

      begin
	Insert into SO080 (RcdNo, CustId, Type, Content, CompCode, ServiceType) Values 
	  (v_RcdCnt, cr1.CustId, cr1.CitemCode, 'PTCode='||cr1.PtCode, v_CompCode, v_ServiceType);
      exception
	when others then
	  DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      end;

    exception 
      when others then
	DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    end;
  end loop;

  commit;
  p_RetMsg := '執行完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒, 讀取筆數='||v_RcdCnt||', 更新SO003筆數='||v_UpdCnt;
  return 0;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := SQLERRM;
    return -99;
end;
/