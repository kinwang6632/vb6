 
create or replace function SF_NewWip3
  (p_CustId in number) return number
  as
  --v_AcceptTime date;
  --v_PrCode number;
  v_NewWipCode3 number := 0;
  v_NewWipName3 varchar2(12);
  v_IgnoreCode number;
  v_RefNo number;
 
  -- 該客戶於停拆移機未完工log(SO073)檔中,最近的一筆log資料
  cursor cc1(vc_IgnoreCode number) is
    select RefNo from SO073 where CustId=p_CustId and PRCode!=vc_IgnoreCode
	order by AcceptTime desc;
 
begin
  -- Check parameters
  if p_CustId is null then
    return -1;
  end if;
 
  -- 不檢查派工類別為'D/M '者(RefNo=4), 故先取'D/M '之類別代碼
  begin
    select CodeNo into v_IgnoreCode from CD007 where RefNo=4;
  exception
    when others then
      return -3;
  end;
 
  -- 取派工狀態類別代碼檔(CD036)代碼=0的名稱
  begin
    select Description into v_NewWipName3 from CD036 where CodeNo=0;
  exception 
    when others then
      return -2;
  end;
 
  -- 取該客戶於停拆移機未完工log(SO073)檔中,最近的一筆log資料的停拆機類別代碼的參考號
  -- 若有資料, 判斷該參考號, 來決定將該客戶的派工狀態三改為何值
  -- 若無資料, 則將派工狀態三改為0
  open cc1(v_IgnoreCode);
  fetch cc1 into v_RefNo;
  if cc1%notfound then
    v_RefNo := 0;
  end if;
  close cc1;
 
  if v_RefNo=0 then
    v_NewWipCode3 := 0;		-- 無
  elsif v_RefNo=3 then
    v_NewWipCode3 := 11;	-- 移機中
  elsif v_RefNo=2 or v_RefNo=5 then
    v_NewWipCode3 := 12;	-- 拆機中
  elsif v_RefNo=1 then
    v_NewWipCode3 := 13;	-- 停機中
  elsif v_RefNo=4 then
    v_NewWipCode3 := 14;	-- 移拆中
  end if;
  if v_RefNo!=0 then
    select Description into v_NewWipName3 from CD036 where CodeNo=v_NewWipCode3; -- 取名稱
  end if;
  update SO001 set WipCode3=v_NewWipCode3, WipName3=v_NewWipName3 where CustId=p_CustId;
 
  commit;
  return 0;
 
exception
  when others then
    rollback;
    return -99;
end;
/
