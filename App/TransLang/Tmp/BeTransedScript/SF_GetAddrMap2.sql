
CREATE OR REPLACE FUNCTION SF_GetAddrMap2
  (p_AddrSort IN VARCHAR2, 
   p_Noe1 in number, p_Noe2 in number, p_Noe3 in number, p_Noe4 in number, 
   p_ZipCode OUT VARCHAR2,
   p_CircuitNo OUT VARCHAR2,
   p_AreaCode OUT VARCHAR2,
   p_AreaName OUT VARCHAR2,
   p_ServCode OUT VARCHAR2,
   p_ServName OUT VARCHAR2,
   p_ClctAreaCode OUT VARCHAR2,
   p_ClctAreaName OUT VARCHAR2,
   p_ClctEn OUT VARCHAR2,
   p_ClctName OUT VARCHAR2,
   p_BTCode OUT VARCHAR2,
   p_BTName OUT VARCHAR2,
   p_MduId OUT VARCHAR2,
   p_MduName OUT VARCHAR2) RETURN NUMBER
  as
  v_Flag number := 0;	-- 若找到, 則設為1
  v_StrtCode number;

  -- 從cursor取得每筆資料欄位的對應變數
  v_Ord		number;
  --v_AddrSortA	varchar2(53);
  --v_AddrSortB	varchar2(53);
  v_Noe		number;
  v_Lane	varchar2(10);
  v_Alley	varchar2(10);
  v_Alley2	varchar2(10);
  --v_No1A	number;

  v_ZipCode	varchar2(5);
  v_CircuitNo	varchar2(20);
  v_AreaCode	varchar2(3);
  v_AreaName	varchar2(12);
  v_ServCode	varchar2(3);
  v_ServName	varchar2(12);
  v_ClctAreaCode varchar2(3);
  v_ClctAreaName varchar2(12);
  v_ClctEn 	varchar2(10);
  v_ClctName 	varchar2(20);
  v_BTCode 	number;
  v_BTName 	varchar2(12);
  v_MduId 	varchar2(8);
  v_MduName 	varchar2(30);

/*
    select AddrSortA, AddrSortB, Noe, ZipCode, CircuitNo, 
	AreaCode, AreaName, ServCode, ServName,
	ClctAreaCode, ClctAreaName, ClctEn, ClctName, 
	BTCode, BTName, MduId, MDUName
*/
  cursor cc1 is
    select Ord, Noe, Lane, Alley, Alley2
      from SO016 where StrtCode=to_number(substrb(p_AddrSort,1,6)) and
      p_AddrSort >= AddrSortA and p_AddrSort <= AddrSortB
      order by Ord desc;

begin
  -- Check parameters
  if (p_AddrSort is null) then
    return -1;
  end if;

  -- 設定傳回值的初值
  p_ZipCode   := '';
  p_CircuitNo := '';
  p_AreaCode  := '???';
  p_AreaName  := '(未設)';
  p_ServCode  := '???';
  p_ServName  := '(未設)';
  p_ClctAreaCode := '???';
  p_ClctAreaName := '(未設)';
  p_ClctEn    := '???';
  p_ClctName  := '(未設)';
  p_BTCode    := '';
  p_BTName    := '';
  p_MduId     := '';
  p_MduName   := '';

  v_StrtCode := to_number(substrb(p_AddrSort,1,6));	-- 街道代碼
  open cc1;
  loop
    fetch cc1 into v_Ord, v_Noe, v_Lane, v_Alley, v_Alley2;
    exit when cc1%notfound or v_Flag=1;

    -- 若無單雙號限制, 則找到第一筆資料, 就是了
    if v_Noe = 0 then
      v_Flag := 1;
      exit;
    end if;

    -- 檢查是否符合單雙號條件
    if v_Alley2 is not null then	-- 若衖有值, 則檢查其是否衖內單雙號
      if p_Noe4 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    elsif v_Alley is not null then	-- 若弄有值, 則檢查其是否弄內單雙號
      if p_Noe3 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    elsif v_Lane is not null then	-- 若巷有值, 則檢查其是否巷內單雙號
      if p_Noe2 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    else				-- 若巷弄衖皆無值, 則檢查其是否街內單雙號
      if p_Noe1 = v_Noe then
	v_Flag := 1;
	exit;
      end if;
    end if;

  end loop;
  close cc1;

  if v_Flag = 1 then
    -- 若找到, 則取該筆應相關欄位資料
    select ZipCode, CircuitNo, AreaCode, AreaName, ServCode, ServName,
	ClctAreaCode, ClctAreaName, ClctEn, ClctName, BTCode, BTName, 
	MduId, MDUName
      into v_ZipCode, v_CircuitNo, v_AreaCode, v_AreaName, v_ServCode, v_ServName, 
	v_ClctAreaCode, v_ClctAreaName, v_ClctEn, v_ClctName, v_BTCode, v_BTName, 
	v_MduId, v_MduName
      from SO016 where StrtCode=v_StrtCode and Ord=v_Ord;

    p_ZipCode   := v_ZipCode;
    p_CircuitNo := v_CircuitNo;
    p_AreaCode  := v_AreaCode;
    p_AreaName  := v_AreaName;
    p_ServCode  := v_ServCode;
    p_ServName  := v_ServName;
    p_ClctAreaCode := v_ClctAreaCode;
    p_ClctAreaName := v_ClctAreaName;
    p_ClctEn    := v_ClctEn;
    p_ClctName  := v_ClctName;
    p_BTCode    := v_BTCode;
    p_BTName    := v_BTName;
    p_MduId     := v_MduId;
    p_MduName   := v_MduName;
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