/*
@c:\emc_script\SF_DailyOp3

set serveroutput on
variable nn number
variable msg varchar2(300)
exec :nn := SF_DailyOp3('testv30', 1, '20020101', '20020131', 'IM', 'Pierre', :msg);

--exec :nn := SF_DailyOp3('EMCYMS,EMCLK,EMCKP,EMCNTP', 1, '20020902', '20020902', 'IM', 'Pierre', :msg);
print nn
print msg

  批次作業__客服員派工完工收費統計報表 (EMC專用)
  說明: 會依據傳來之各Owner, 取各Owner資料區的工單所附收費資料來統計CSR績效
  檔名: SF_DailyOp3.sql

  注意: 1.需建立Oracle物件: SO503(Table), V_SO503Charge(View), S_SeqNo1(Sequence)
	2.需grant物件權限給主資料區: cd019, cd039, cm003, so007/8/9, V_SO503Charge, SO041

  IN	p_OwnerList varchar2:	Owner list, 例, 'EMCYMS,EMCLK,EMCKP,EMCNTP'
	-- p_CompCode number:	公司別代碼
	-- p_ServiceType char(1):	服務類別, 例: 'C', 'I'
	p_DateType number;	1=完工日, 2=受理日
	p_RptDate1  varchar2:	統計起始日, 西曆'YYYYMMDD'
	p_RptDate2  varchar2:	統計截止日, 西曆'YYYYMMDD'
	p_BillType varchar2:	單據類別, 例, 'IMP'
	p_Operator varchar2:	操作者名稱

  OUT	p_RetMsg VARCHAR2： 結果訊息, 呼叫端變數至少需100 bytes

  return NUMBER: 結果代碼, 為該批資料於SO503的資料批號
        >0: 資料批號, 表示正常完畢
	-1: p_RetMsg='參數錯誤'
	-2: p_RetMsg='取各種費用金額失敗, 收費項目='
	-3: p_RetMsg='SQL錯誤: ' || v_SQL2
	-4: p_RetMsg='取 Sequence S_SeqNo1 的次一可用值失敗'
	-5: p_RetMsg='新增資料失敗, 訊息='||SQLERRM;
	-6: p_RetMsg='環境錯誤: 業者參數檔中無服務別等於C的資料'
	-7: p_RetMsg='環境錯誤: 公司別代碼檔中無此公司別的資料, 公司別=' || v_CompCode
        -99：p_RetMsg='其他錯誤'

  SO503 (SeqNo		number(8),		-- 批號
	CompCode 	number(3),		-- 公司別代碼
	CompName 	varchar2(20),		-- 公司別名稱
	ServiceType 	char(1),		-- 服務別
	AcceptEn 	varchar2(10),		-- 受理/介紹人員代碼
	AcceptName 	varchar2(20),		-- 受理/介紹人員姓名
	Cnt1		number(8) default 0,	-- '裝復機費=牌價'數量
	Cnt2		number(8) default 0,	-- '裝復機費0<X<牌價'數量
	Cnt3		number(8) default 0,	-- '裝復機費=0'數量
	Cnt4		number(8) default 0,	-- '分機裝置費'數量
	Cnt5		number(8) default 0,	-- '分機訊號費'數量
	CntTotal	number(8) default 0,	-- '合計'數量
	RptDate		date,			-- 報表資料日
	ReportTime	date,			-- 報表執行時間
	Operator	varchar2(20)		-- 操作者名稱
	MediaCode	number(3)		-- 媒體介紹類別代碼
	MediaName	varchar2(20)		-- 媒體介紹類別名稱
	Cnt6		number(8) default 0,	-- 'eBox(賣)'數量
	Cnt7		number(8) default 0,	-- 'eBox(租)'數量
	);

  By: Pierre
  Date: 2002.09.10
	2002.09.11 改: 只有統計I單, 且I單條件由受理人員改為介紹類別為客服部
	2002.09.12 改: 若項目為多台之分機訊號費, 則要加上對應台數, 而不是筆數
	2002.09.13 改: SO503加上媒體介紹類別代碼之區隔, 以及MediaCode=11(TM組)的資料也加入統計
	2002.09.19 改bug: 'Acceptime' ==> 'AcceptTime'
	2002.12.12 改: 因應轉代碼資料完畢, MediaCode 1=>10, 11=>18
	2002.12.13 改: 因應轉代碼資料完畢, CitemCode已改掉了

	2003.03.22 改: 因應公司別代碼檔(CD039)於各平台已不是只有一筆, 可能會多筆, 所以取公司別
			與名稱的方式改為去SO041取系統別(就是公司別), 再去CD039中取公司名稱
			注意: 需再grant SO041的權限給主資料區的操作者
	2003.06.06 改: 依據92.06.06的規格討論做修改, 請參考"客戶修正版-派工完工收費統計表規格_920606.doc"內容

*/
--create or replace function SF_DailyOp3(p_OwnerList varchar2, p_CompCode number, p_ServiceType char, 
create or replace function SF_DailyOp3(p_OwnerList varchar2, 
	p_DateType number, p_RptDate1 varchar2, p_RptDate2 varchar2, p_BillType varchar2,
	p_Operator varchar2, p_RetMsg out varchar2) return number
as
  c39		char(1) := chr(39);
  v_RetCode 	number;
  v_RetMsg 	varchar2(1000);
  v_StartExecTime date;
  v_StopExecTime date;
  v_SQL 	varchar2(500);
  v_SQL2 	varchar2(1000);
  v_SQL3	varchar2(500);
  v_SQLcsr	varchar2(500);
  v_SQLcharge	varchar2(200);
  v_Tmp 	number;
  v_DateFldName varchar2(20);
  v_TableName 	varchar2(20);
  v_AllBillType varchar2(10) := 'IMP';
  v_SeqNo	number;
  v_Now		date;

  v_InstFeeCode 	number := 101;		-- 1.裝機費項目代碼
  v_ReInstFeeCode	number := 106;		-- 2.復機費項目代碼
  v_InstFee2Code	number := 108;		-- 3.分機裝置費項目代碼
  v_BCFeeCode		number := 301;		-- 4.基本台訊號費項目代碼
  v_BCFee2Code1		number := 306;		-- 5.分機訊號費項目代碼1(1台)
  v_BCFee2Code2		number := 111;		-- 5.分機訊號費項目代碼2(1台)
  v_BCFee2Code3		number := 112;		-- 5.分機訊號費項目代碼3(2台)
  v_BCFee2Code4		number := 113;		-- 5.分機訊號費項目代碼4(3台)
  v_BCFee2Code5		number := 114;		-- 5.分機訊號費項目代碼5(4台)
  v_BCFee2Code6		number := 115;		-- 5.分機訊號費項目代碼6(5台)

  v_eBoxSaleFeeCode1	number := 801;		-- 6. eBox(賣)一次付清項目代碼
  v_eBoxSaleFeeCode2	number := 802;		-- 6. eBox(賣)分期付款項目代碼
  v_eBoxRentFeeCode	number := 804;		-- 7. eBox(租)保證金項目代碼

/*
  v_InstFeeCode 	number := 3;		-- 1.裝機費項目代碼
  v_ReInstFeeCode	number := 4;		-- 2.復機費項目代碼
  v_InstFee2Code	number := 6;		-- 3.分機裝置費項目代碼
  v_BCFeeCode		number := 1;		-- 4.基本台訊號費項目代碼
  v_BCFee2Code1		number := 21;		-- 5.分機訊號費項目代碼1(1台)
  v_BCFee2Code2		number := 22;		-- 5.分機訊號費項目代碼2(1台)
  v_BCFee2Code3		number := 23;		-- 5.分機訊號費項目代碼3(2台)
  v_BCFee2Code4		number := 24;		-- 5.分機訊號費項目代碼4(3台)
  v_BCFee2Code5		number := 25;		-- 5.分機訊號費項目代碼5(4台)
  v_BCFee2Code6		number := 66;		-- 5.分機訊號費項目代碼6(5台)
*/
  v_InstFee 		number := 0;		-- 1.裝機費牌價
  v_ReInstFee		number := 0;		-- 2.復機費牌價
  v_InstFee2		number := 0;		-- 3.分機裝置費牌價
  v_BCFee		number := 0;		-- 4.基本台訊號費
  v_BCFee2		number := 0;		-- 5.分機訊號費

  v_eBoxSaleFee1	number := 0;		-- 6. eBox(賣)一次付清牌價
  v_eBoxSaleFee2	number := 0;		-- 6. eBox(賣)分期付款牌價 ==> 此項不能直接比較
  v_eBoxRentFee		number := 0;		-- 7. eBox(租)保證金牌價

  v_Cnt1	number;		-- 1.		==> [裝復機費=牌價]數值
  v_Cnt2	number;		-- 2.		==> [裝復機費0<X<牌價]數值
  v_Cnt3	number;		-- 3.		==> [裝復機費=0]數值
  v_Cnt4	number;		-- 4.		==> [分機裝置費]數值
  v_Cnt5	number;		-- 5.		==> [分機訊號費]數值
  v_Cnt6	number;		-- 6.		==> [eBox(賣)]數值
  v_Cnt7	number;		-- 7.		==> [eBox(租)]數值
  v_CntTotal	number;		-- 合計

  v_FlagA	number;		-- 是否有該項目
  v_FlagB	number;		-- 
  v_FlagC	number;		-- 
  v_FlagD	number;		-- 
  v_FlagE	number;		-- 
  v_FlagF	number;		-- 
  v_FlagG	number;		-- 

  v_Cnt0	number;		-- record count
  I		number;
  v_Index	number;
  v_OwnerList	varchar2(500);
  v_OwnerCnt	number;
  v_FirstOwner	varchar2(20);	-- 第一資料區, 即目前資料區
  v_CompCode	number;		-- 該資料區的公司別代碼
  v_CompName 	varchar2(20);	-- 該資料區的公司別名稱
  v_ServiceType char(1);

  x_EmpNo	varchar2(10);	-- 該CSR員工代號
  x_EmpName	varchar2(30);	-- 該CSR員工姓名
  x_MediaCode	number;		-- 該工單媒體介紹類別代碼
  x_MediaName	varchar2(20);	-- 該工單媒體介紹類別名稱
  x_SNo		varchar2(15);	-- 暫時單號
  x_WoCode	number;		-- 派工類別
  x_CitemCode	number;		-- 收費項目代碼
  x_RealAmt	number;		-- 實收金額

  TYPE BillAry IS TABLE OF char INDEX BY BINARY_INTEGER; --宣告一個Table的型態(用法和陣列同)
  BillType BillAry;		-- array for Bill types

  TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
  OwnerName Varchar2Ary;	-- array for owners
  DataTableName Varchar2Ary;	-- array for Data tables

  type CurTyp is ref cursor;  --自訂cursor型態
  v_DyCursor1 CurTyp;          --供dynamic SQL
  v_DyCursor2 CurTyp;          --供dynamic SQL
  v_DyCursor3 CurTyp;          --供dynamic SQL


/*
  cursor CSR is	
    select EmpNo, EmpName from CM003 where WorkClass = 'A';    -- All CSR

  cursor Charge(v_BillNo varchar2) is
    select * from SO034 where BillNo = v_BillNo;
*/
begin

  v_StartExecTime := sysdate;		-- 開始執行時間
  v_Now := sysdate;

/*
  -- 取單據類別放入BillType array中
  for i in 1..length(p_BillType) loop
    BillType(i) := substr(p_BillType, i, 1);
  end loop;

  -- 決定資料檔種類: 目前為SO007(I單), SO009(P單)  ==> 92.06.06 取消
  DataTableName(1) := 'SO007';
  DataTableName(2) := 'SO009';
  DBMS_OUTPUT.PUT_LINE('['||DataTableName(1)||']');
  DBMS_OUTPUT.PUT_LINE('['||DataTableName(2)||']');
*/

  --切割Owner字串至OwnerName陣列中
  v_OwnerCnt := 0;
  if p_OwnerList is not null then
    v_OwnerList := p_OwnerList;
    v_Index := 1;
    I := 1;
    while v_OwnerList is not null loop
      v_Index := instr(v_OwnerList, ',');
      if v_Index > 0 then
	begin
          OwnerName(I) := ltrim(rtrim(substr(v_OwnerList, 1, v_Index-1)));
	  v_OwnerList := substr(v_OwnerList, v_Index+1);
          I:=I+1;
	exception
	  when others then
	    v_OwnerList := null;
	end;
      else
        OwnerName(I) := rtrim(ltrim(v_OwnerList));
        v_OwnerList := null;
      end if;
    end loop;
    v_OwnerCnt := I;
    v_FirstOwner := OwnerName(1);		-- 記住第一個owner name
  end if;

/*  for j in 1..v_OwnerCnt loop
    DBMS_OUTPUT.PUT_LINE(OwnerName(j));
  end loop; */


  -- 取各種項目之牌價
  begin
    v_Tmp := v_InstFeeCode;
    select nvl(Amount,0) into v_InstFee from CD019 where CodeNo = v_InstFeeCode;
    v_Tmp := v_ReInstFeeCode;
    select nvl(Amount,0) into v_ReInstFee from CD019 where CodeNo = v_ReInstFeeCode;
/*
    v_Tmp := v_InstFee2Code;
    select nvl(Amount,0) into v_InstFee2 from CD019 where CodeNo = v_InstFee2Code;
    v_Tmp := v_BCFeeCode;
    select nvl(Amount,0) into v_BCFee from CD019 where CodeNo = v_BCFeeCode;
    v_Tmp := v_BCFee2Code;
    select nvl(Amount,0) into v_BCFee2 from CD019 where CodeNo = v_BCFee2Code;
*/
    v_Tmp := v_eBoxSaleFeeCode1;
    select nvl(Amount,0) into v_eBoxSaleFee1 from CD019 where CodeNo = v_eBoxSaleFeeCode1;
    v_Tmp := v_eBoxRentFeeCode;
    select nvl(Amount,0) into v_eBoxRentFee from CD019 where CodeNo = v_eBoxRentFeeCode;
  exception
    when others then
      p_RetMsg := '取各種費用金額失敗, 收費項目=' || v_Tmp;
      return -2;
  end;


  -- 若有統計日範圍
  if p_DateType = 1 then
    v_DateFldName := 'FinTime';
  elsif p_DateType = 2 then
    v_DateFldName := 'AcceptTime';
  else
    p_RetMsg := '參數錯誤';
    return -1;
  end if;
  if p_RptDate1 is not null then
    v_SQL := v_SQL || ' and '||v_DateFldName||'>=' || GiPackage.QryDTString0(to_date(p_RptDate1,'YYYYMMDD')) ||
		' and '||v_DateFldName||'<' || GiPackage.QryDTString0(to_date(p_RptDate2,'YYYYMMDD')) || '+1';
  end if;

  v_SQL := substr(v_SQL, 5);		-- 除去開頭的 ' and'
  v_Cnt0 := 0;

  --取 Sequence S_SeqNo1 的次一可用值
  begin
    Select S_SeqNo1.Nextval Into v_SeqNo From Dual;
  exception
    when others then
      p_RetMsg := '取 Sequence S_SeqNo1 的次一可用值失敗';
      return -4;
  end;
  delete from SO503 where SeqNo = v_SeqNo;	-- 清除同一批號的資料

  for i in 1..v_OwnerCnt loop		-- Loop 每一Owner
    v_ServiceType := 'C';
    v_CompCode := null;

    -- 取該資料區之預設公司別
    begin
      -- v_SQL3 := 'select CodeNo, Description from '||OwnerName(i)||'.CD039 where RowNum=1';
      v_SQL3 := 'select SysId from '||OwnerName(i)||'.SO041 where ServiceType like '||c39||'%C%'||c39||' and RowNum=1';
      open v_DyCursor1 for v_SQL3;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL錯誤: ' || v_SQL3;
        return -3;
    end;
    fetch v_DyCursor1 INTO v_CompCode; 
    close v_DyCursor1;

    if v_CompCode is null then
      rollback;
      p_RetMsg := '環境錯誤: 業者參數檔中無服務別等於C的資料';
      return -6;
    end if;

    --根據公司別, 去CD039中取公司名稱
    begin
      v_SQL3 := 'select Description from '||OwnerName(i)||'.CD039 where CodeNo='||v_CompCode;
      open v_DyCursor1 for v_SQL3;
    exception
      when others then
        rollback;
        p_RetMsg := '環境錯誤: 公司別代碼檔中無此公司別的資料, 公司別=' || v_CompCode;
        return -7;
    end;
    fetch v_DyCursor1 INTO v_CompName; 
    close v_DyCursor1;
    DBMS_OUTPUT.PUT_LINE(v_CompCode||','||v_CompName);

/*
    -- 該Owner資料區的CM003.WorkClass = 'C'者, 為CSR
    begin
      v_SQLcsr := 'select EmpNo, EmpName from '||OwnerName(i)||'.CM003 where WorkClass='||c39||'A'||c39;    -- All CSR
      open v_DyCursor1 for v_SQLcsr;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL錯誤: ' || v_SQLcsr;
	DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
        return -3;
    end;
*/

    -- 以工單之介紹媒介為客服部或TM組當受理人員來統計
    -- 2002.12.12 改: MediaCode 1=>10, 11=>18
    begin
      v_SQLcsr := 'select IntroId EmpNo, IntroName EmpName, MediaCode, MediaName from '||OwnerName(i)||'.SO007 where '||
		v_SQL||' and FinTime is not null and (MediaCode=10 or MediaCode=18) and IntroId is not null'||
		' group by MediaCode, MediaName, IntroId, IntroName';    -- All CSR and TM group
      open v_DyCursor1 for v_SQLcsr;
    exception
      when others then
        rollback;
        p_RetMsg := 'SQL錯誤: ' || v_SQLcsr;
	DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
        return -3;
    end;

    loop 				-- Loop 每一CSR
      fetch v_DyCursor1 INTO x_EmpNo, x_EmpName, x_MediaCode, x_MediaName; 
      exit when v_DyCursor1%NOTFOUND;

      -- initialize variables
      v_Cnt1 := 0;
      v_Cnt2 := 0;
      v_Cnt3 := 0;
      v_Cnt4 := 0;
      v_Cnt5 := 0;
      v_Cnt6 := 0;
      v_Cnt7 := 0;
      v_CntTotal := 0;

      --DBMS_OUTPUT.PUT_LINE(x_EmpNo||', '||x_EmpName);

	v_SQL2 := 'select SNo, InstCode WoCode from '||OwnerName(i)||'.SO007';
        v_SQL2 := v_SQL2||' where '||v_SQL||' and FinTime is not null and IntroId='||c39||x_EmpNo||c39||
			' and MediaCode='||x_MediaCode;
        begin
	  open v_DyCursor2 for v_SQL2;
        exception
	  when others then
	    rollback;
	    p_RetMsg := 'SQL錯誤: ' || v_SQL2;
	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    return -3;
        end;
        loop       	-- loop 每一筆裝機單資料
	  v_FlagA := 0;
	  v_FlagB := 0;
	  v_FlagC := 0;
	  v_FlagD := 0;
	  v_FlagE := 0;
	  v_FlagF := 0;
	  v_FlagG := 0;

	  fetch v_DyCursor2 INTO x_SNo, x_WoCode; 
	  exit when v_DyCursor2%NOTFOUND;
	  v_SQLcharge := 'select CitemCode, RealAmt from '||OwnerName(i)||'.V_SO503Charge where BillNo='||c39||x_SNo||c39;
	  begin
	    open v_DyCursor3 for v_SQLcharge;
	  exception
	    when others then
	      rollback;
	      p_RetMsg := 'SQL錯誤: ' || v_SQLcharge;
	      DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	      return -3;
	  end;
	  loop		-- loop 該受理人員之該張工單之相關收費, 做以下處理
	    fetch v_DyCursor3 INTO x_CitemCode, x_RealAmt; 
	    exit when v_DyCursor3%NOTFOUND;
	    v_Cnt0 := v_Cnt0 + 1;

	    -- 1. 若項目為裝機費，且金額=[裝機費牌價]，或項目為復機費，且金額=[復機費牌價]，則[裝復機費=牌價]數值加1
	    if x_CitemCode=v_InstFeeCode and x_RealAmt=v_InstFee or 
		x_CitemCode=v_ReInstFeeCode and x_RealAmt=v_ReInstFee then
	      v_FlagA := 1;
	      v_Cnt1 := v_Cnt1 + 1;

	    -- 2. 若項目為裝機費，且0<金額<[裝機費牌價]，或項目為復機費，且0<金額<[復機費牌價]，則[裝復機費0<X<牌價]數值加1
	    elsif x_CitemCode=v_InstFeeCode and x_RealAmt>0 and x_RealAmt<v_InstFee or 
		x_CitemCode=v_ReInstFeeCode and x_RealAmt>0 and x_RealAmt<v_ReInstFee then
	      v_FlagB := 1;
	      v_Cnt2 := v_Cnt2 + 1;

	    -- 3. 若項目為裝機費，且金額=0，或項目為復機費，且金額=0，則[裝復機費=0]數值加1
	    elsif x_CitemCode=v_InstFeeCode and x_RealAmt=0 or x_CitemCode=v_ReInstFeeCode and x_RealAmt=0 then
	      v_FlagC := 1;
	      v_Cnt3 := v_Cnt3 + 1;

	    -- 4. 若項目為分機裝置費，則[分機裝置費]數值加1
	    elsif x_CitemCode=v_InstFee2Code then
	      v_FlagD := 1;
	      v_Cnt4 := v_Cnt4 + 1;

	    -- 5. 項目為分機訊號費，且該工單尚有分機裝置費項目(代碼108), 則該分機訊號費不計算, 否則[分機訊號費]數值加1
	    elsif x_CitemCode in (v_BCFee2Code1,v_BCFee2Code2,v_BCFee2Code3,v_BCFee2Code4,v_BCFee2Code5,v_BCFee2Code6) then
	      v_FlagE := 1; 

	    -- 6. 項目為eBox(賣)一次付清，且金額=[eBox賣牌價], 或  項目為eBox(賣)分期, 且金額>0, 則[eBox(賣)]數值加1
	    elsif x_CitemCode=v_eBoxSaleFeeCode1 and x_RealAmt=v_eBoxSaleFee1 or 
		x_CitemCode=v_eBoxSaleFeeCode1 and x_RealAmt>0 then
	      v_FlagF := 1;
	      v_Cnt6 := v_Cnt6 + 1;

	    -- 7. 項目為eBox(租)的保證金 , 且金額>0, 則[eBox(租)]數值加1
	    elsif x_CitemCode=v_eBoxRentFeeCode and x_RealAmt>0 then
	      v_FlagG := 1;
	      v_Cnt7 := v_Cnt7 + 1;
	    end if;

	    if v_FlagE=1 and v_FlagD=0 then	-- 5. 此項只有該工單無分機裝置費時, 才計算
	      if x_CitemCode=v_BCFee2Code1 or x_CitemCode=v_BCFee2Code2 then	-- 加上對應台數, 不是筆數, 2002.09.12
		v_Cnt5 := v_Cnt5 + 1;
	      elsif x_CitemCode=v_BCFee2Code3 then
		v_Cnt5 := v_Cnt5 + 2;
	      elsif x_CitemCode=v_BCFee2Code4 then
		v_Cnt5 := v_Cnt5 + 3;
	      elsif x_CitemCode=v_BCFee2Code5 then
		v_Cnt5 := v_Cnt5 + 4;
	      elsif x_CitemCode=v_BCFee2Code6 then
		v_Cnt5 := v_Cnt5 + 5;
	      end if;
	    end if;
	
	  end loop;	-- loop 該受理人員之該張工單之相關收費

	end loop;	-- loop 每一筆裝機單資料
	close v_DyCursor2;

      -- 合計
      v_CntTotal := v_Cnt1 + v_Cnt2 + v_Cnt3 + v_Cnt4 + v_Cnt5 + v_Cnt6 + v_Cnt7;

      -- 將結果加入此程式統計當時資料區內的統計檔SO503
      if v_CntTotal>0 then	-- 若總數大於0, 才需要加入統計檔
	begin
	  insert into SO503 (SeqNo, CompCode, CompName, ServiceType, AcceptEn, AcceptName, 
		Cnt1, Cnt2, Cnt3, Cnt4, Cnt5, Cnt6, Cnt7, CntTotal, RptDate, ReportTime, Operator, MediaCode, MediaName) 
		values (v_SeqNo, v_CompCode, v_CompName, v_ServiceType, x_EmpNo, x_EmpName, 
		v_Cnt1, v_Cnt2, v_Cnt3, v_Cnt4, v_Cnt5, v_Cnt6, v_Cnt7, v_CntTotal, to_date(p_RptDate1,'YYYYMMDD'), 
		v_Now, p_Operator, x_MediaCode, x_MediaName);
	exception
	  when others then
    	    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
	    rollback;
	    p_RetMsg := '新增資料失敗, 訊息='||SQLERRM;
	    return -5;
	end;
      end if;

    end loop;		-- Loop 每一CSR
    close v_DyCursor1;

  end loop;		-- Loop 每一Owner


  commit;
  p_RetMsg := '產生完畢, 共花費'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||'秒, v_Cnt0='||v_Cnt0;

  return v_SeqNo;

exception
  when others then
    DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    rollback;

    p_RetMsg := SQLERRM;
    return -99;
end;
/