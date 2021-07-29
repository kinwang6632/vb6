create or replace function SF_StrtCompute
  (p_RetMsg out varchar2) return number
  as
  v_Count number;

  cursor cc1 is 
    select CodeNo from CD017 order by CodeNo;

begin

  -- 取所有有效客戶之客戶編號與裝機地址的街道代碼, 存至暫存檔中
  delete from TMP0001;
  insert into TMP0001 (StrtCode, CustId) 
    (select B.StrtCode, A.CustId from SO001 A, SO014 B 
	    where A.InstAddrNo=B.AddrNo and A.CustStatusCode=1);

  -- Loop 每一街道
  --   計算該街道之有效戶數
  --   更新街道代碼檔中之有效戶數欄位
  for cr1 in cc1 loop
    select count(*) into v_Count from TMP0001 where StrtCode=cr1.CodeNo;
    update CD017 set InstCnt=v_Count where CodeNo=cr1.CodeNo;
  end loop;

  commit;
  p_RetMsg := '執行完畢';
  return 0;

exception
  when others then
    rollback;
    p_RetMsg := '其他錯誤';
    return -99;
end;
/