
create or replace function SF_GetAddrMap
  (p_TableName in varchar2, p_AddrSort in varchar2, p_CircuitNo out varchar2, 
   p_AreaCode out varchar2, p_ServCode out varchar2) return number
  as
  v_Flag number := 0;	-- 若找到, 則設為1
  v_TS1 varchar2(10);
  v_TS2 varchar2(10);
  v_TS3 varchar2(10);
  v_TS4 varchar2(10);
  v_TN1 number;
  v_TN2 number;
  v_TN3 number;
  v_TN4 number;

  -- 從cursor取得每筆資料欄位的對應變數
  v_AddrSortA	varchar2(53);
  v_AddrSortB	varchar2(53);
  v_Noe		number;
  v_CircuitNo	varchar2(20);
  v_AreaCode	varchar2(3);
  v_ServCode	varchar2(3);


  cursor cc1 is
    select AddrSortA, AddrSortB, Noe, CircuitNo, AreaCode, ServCode
      from SO016 where StrtCode=to_number(substrb(p_AddrSort,1,6)) and
      p_AddrSort >= AddrSortA and p_AddrSort <= AddrSortB
      order by Ord desc;

begin
  if (p_TableName is null) or (p_AddrSort is null) then
    return -1;
  end if;
/*
  p_CircuitNo := '168.95.192.1';
  p_AreaCode  := 'A01';
  p_ServCode  := 'B02';
*/
  -- 設定傳回值的初值
  p_CircuitNo := '';
  p_AreaCode  := '';
  p_ServCode  := '';

  open cc1;
  loop
    fetch cc1 into v_AddrSortA, v_AddrSortB, v_Noe, v_CircuitNo, v_AreaCode, v_ServCode;
    exit when cc1%notfound or v_Flag=1;
/*
    v_TS1 := ltrim(substrb(p_AddrSort, 15, 4));
    v_TS2 := ltrim(substrb(p_AddrSort, 25, 4));
    v_TS3 := ltrim(substrb(p_AddrSort, 35, 4));
    v_TS4 := ltrim(substrb(p_AddrSort, 45, 4));

    v_TN1 := 0;
    v_TN2 := 0;  
    v_TN3 := 0;  
    v_TN4 := 0; 

    if v_TS1 is not null then
    begin
      v_TN1 := nvl(to_number(v_TS1),0);
    exception
      when others then
        v_TN1 := 0;
    end;
    end if;
    if v_TS2 is not null then
    begin
      v_TN2 := nvl(to_number(v_TS2),0);
    exception
      when others then
        v_TN2 := 0;
    end;
    end if;
    if v_TS3 is not null then
    begin
      v_TN3 := nvl(to_number(v_TS3),0);
    exception
      when others then
        v_TN3 := 0;
    end;
    end if;
    if v_TS4 is not null then
    begin
      v_TN4 := nvl(to_number(v_TS4),0);
    exception
      when others then
        v_TN4 := 0;
    end;
    end if;
*/

    -- 找到第一筆資料, 即取其相關資料傳回
    v_Flag := 1;
    p_CircuitNo := v_CircuitNo;
    p_AreaCode  := v_AreaCode;
    p_ServCode  := v_ServCode;
    exit;

  end loop;
  close cc1;

  if v_Flag = 1 then
    return 0;
  else
    return -2;
  end if;

exception
  when others then
    if cc1%isopen then
      close cc1;
    end if;
    return -3;
end;
/