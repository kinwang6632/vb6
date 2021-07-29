
create or replace function SF_MduCompute
  (p_MduId in Varchar2, p_RetMsg out Varchar2) return number
  As
  v_CountA number;
  v_CountB number;
  v_CountC number;

begin
  -- 檢查參數
  if (p_MduId is null) then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  -- 建檔戶數[A]: 客戶資料檔中大樓編號=<該大樓>, 
  -- 且客戶狀態!=促銷中(5), 註銷(4)
  select count(*) into v_CountA from SO001 where MduId= p_MduId 
     and CustStatusCode!=4 and CustStatusCode!=5;

  -- 有效戶數[B]: 客戶資料檔中大樓編號=<該大樓>, 且客戶狀態=正常(1)
  select count(*) into v_CountB from SO001 where 
     MduId=p_MduId and CustStatusCode=1;

  -- 停拆/待拆戶數[C]: 建檔戶數[A] - 有效戶數[B]
  v_CountC := v_CountA - v_CountB;

  -- 更新該大樓資料中各種戶數欄位資料
  update SO017 set AcceptCnt=v_CountA, InstCnt=v_CountB, 
     UnInstCnt=v_CountC where MduId=p_MduId;

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
