/*
--@C:\Gird\CSMis\Script\TBC_DailyOp\SF_DailyOp4_30.sql;
@C:\Gird\v300\Script\SF_DailyOp4_30.sql;

variable nn number
variable msg varchar2(1000)

exec :nn := SF_DailyOp4_30(7, 'C', 'MAINTAIN', 1, :msg);
print nn
print msg

  批次作業__現有各類客戶數統計報表(3.0版使用)
  檔名: SF_DailyOp4_30.sql

  注意: 1.需建立Oracle物件: SO504(Table), S_SeqNo1(Sequence)

  IN	p_CompCode number:	公司別代碼
	p_ServiceType varchar2:	服務類別, 例: 'C', 'I'
	p_Operator varchar2:	操作者名稱
	p_Option number:	操作選項,1=行政區, 2=服務區

  OUT	p_RetMsg VARCHAR2： 結果訊息, 呼叫端變數至少需100 bytes

  return NUMBER: 結果代碼, 為該批資料於SO504的資料批號
        >0: 資料批號, 表示正常完畢
	-1: p_RetMsg='參數錯誤'
        -2: p_RetMsg='取 Sequence S_SeqNo1 的次一可用值失敗'
        -3: p_RetMsg='SQL錯誤: ' || v_SQL
        -4: p_RetMsg='新增資料失敗, 訊息='||SQLERRM;
        -99: p_RetMsg='其他錯誤'

  By: Stanley
  Date: 2002.10.02


*/
create or replace function SF_DailyOp4_30(p_CompCode number, p_ServiceType varchar2, 
	p_Operator varchar2, p_Option number, p_RetMsg out varchar2) return number
as
  c39		  char(1) := chr(39);
  v_StartExecTime date;
  v_Now		  varchar2(10);
  v_SeqNo	  number;
  v_FldName       varchar2(30);
  v_Table         varchar2(30);
  v_SQL 	  varchar2(500);
  v_SQL1          varchar2(1000);
  v_Where         varchar2(500);
  v_temp          number;

  x_DataCode      number;
  x_ServAreaCode  varchar2(3);
  x_CustCnt       number;


  type CurTyp is ref cursor;  --自訂cursor型態
  v_DyCursor1 CurTyp;          --供dynamic SQL

begin

  v_StartExecTime := sysdate;   -- 開始執行時間
  v_Now := sysdate;
  v_Table := 'SO001 A, SO002 B';

  -- 參數檢查
  if p_CompCode is null or p_ServiceType is null or p_Option is null or p_Operator is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  -- 若有操作選項條件,則依值決定欄位內容(1=行政區, 2=服務區)
  if p_Option is not null then
    if p_option = 1 then
      v_FldName := 'C.AreaCode';
      v_Table := v_Table||', SO014 C';
      v_SQL := v_SQL || ' and A.InstAddrNo = C.AddrNo and C.CompCode= '||p_CompCode;
    elsif p_option = 2 then
      v_FldName := 'A.ServCode';
    end if;
  end if;

  -- 若有公司別條件
  if p_CompCode is not null then
    v_SQL := v_SQL || ' and B.CompCode = ' || p_CompCode;
  end if;

  -- 若有服務別條件
  if p_ServiceType is not null then
    v_SQL := v_SQL || ' and B.ServiceType =' ||c39||p_ServiceType||c39;
  end if;

  --取 Sequence S_SeqNo1 的次一可用值
  begin
    Select S_SeqNo1.Nextval Into v_SeqNo From Dual;
  exception
    when others then
      p_RetMsg := '取 Sequence S_SeqNo1 的次一可用值失敗';
      return -2;
  end;

  delete from SO504 where SeqNo = v_SeqNo;	-- 清除暫存檔中同一批號的資料

  ------依客戶狀態統計資料--------------------
  begin
    v_SQL1 := 'Select B.CustStatusCode,'||v_FldName||',Count(A.CustId) As intCount'||
              ' From '||v_Table||' Where A.CustId = B.CustId And A.CompCode = B.CompCode'||
              v_SQL||' Group By B.CustStatusCode,'||v_FldName||
              ' Union Select A.CodeNo,B.CodeNo,0 from CD035 A,CD002 B';
    open v_DyCursor1 for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_SQL1;
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return -3;
  end;

  loop
    fetch v_DyCursor1 Into x_DataCode,x_ServAreaCode,x_CustCnt;
    exit when v_DyCursor1%NOTFOUND;

    	begin
	  insert into SO504 (SeqNo, CompCode, ServiceType, RptDate, RptTime, Operator, DataType,
                 ServAreaCode, DataCode, CustCnt) values
		(v_SeqNo, p_CompCode, p_ServiceType, to_date(v_Now), v_StartExecTime, p_Operator, 1,
                 x_ServAreaCode, x_DataCode, x_CustCnt);

	exception
	  when others then
	    rollback;
    	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    p_RetMsg := '新增資料失敗, 訊息='||SQLERRM;
	    return -4;
	end;

  end loop;
  close v_DyCursor1;

  ------依收費方式統計資料--------------------
  v_SQL1 := '';
  begin
    v_SQL1 := 'Select B.CMCode,'||v_FldName||',Count(A.CustId) As intCount'||
              ' From '||v_Table||' Where A.CustId = B.CustId And A.CompCode = B.CompCode'||
              v_SQL||' AND B.CustStatusCode =1'||' Group By B.CMCode,'||v_FldName||
              ' Union Select A.CodeNo,B.CodeNo,0 from CD031 A,CD002 B';
    open v_DyCursor1 for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_SQL1;
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return -3;
  end;

  loop
    fetch v_DyCursor1 Into x_DataCode,x_ServAreaCode,x_CustCnt;
    exit when v_DyCursor1%NOTFOUND;

    	begin

	  insert into SO504 (SeqNo, CompCode, ServiceType, RptDate, RptTime, Operator, DataType,
                 ServAreaCode, DataCode, CustCnt) values
		(v_SeqNo, p_CompCode, p_ServiceType, to_date(v_Now), v_StartExecTime, p_Operator, 2,
                 x_ServAreaCode, x_DataCode, x_CustCnt);
	exception
	  when others then
	    rollback;
    	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    p_RetMsg := '新增資料失敗, 訊息='||SQLERRM;
	    return -4;
	end;

  end loop;
  close v_DyCursor1;

  ------依客戶類別統計資料--------------------
  v_SQL1 := '';
  begin
    v_SQL1 := 'Select A.ClassCode1,'||v_FldName||',Count(A.CustId) As intCount'||
              ' From '||v_Table||' Where A.CustId = B.CustId And A.CompCode = B.CompCode'||
              v_SQL||' AND B.CustStatusCode =1'||' Group By A.ClassCode1,'||v_FldName||
              ' Union Select A.CodeNo,B.CodeNo,0 from CD004 A,CD002 B';
    open v_DyCursor1 for v_SQL1;
  exception
    when others then
      rollback;
      p_RetMsg := 'SQL錯誤: ' || v_SQL1;
      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
      return -3;
  end;

  loop
    fetch v_DyCursor1 Into x_DataCode,x_ServAreaCode,x_CustCnt;
    exit when v_DyCursor1%NOTFOUND;

    	begin

	  insert into SO504 (SeqNo, CompCode, ServiceType, RptDate, RptTime, Operator, DataType,
                 ServAreaCode, DataCode, CustCnt) values
		(v_SeqNo, p_CompCode, p_ServiceType, to_date(v_Now), v_StartExecTime, p_Operator, 3,
                 x_ServAreaCode, x_DataCode, x_CustCnt);
	exception
	  when others then
	    rollback;
    	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    p_RetMsg := '新增資料失敗, 訊息='||SQLERRM;
	    return -4;
	end;

  end loop;
  close v_DyCursor1;

  commit;
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒';
  return v_SeqNo;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;
    p_RetMsg := SQLERRM;
    return -99;
end;
/