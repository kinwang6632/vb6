 
CREATE OR REPLACE FUNCTION SF_MduToMain
  (p_MduId VARCHAR2, p_ClassSQL VARCHAR2, p_StatusSQL VARCHAR2, p_CitemCode NUMBER,
   p_CitemName VARCHAR2, p_Period NUMBER, p_Amount NUMBER, p_ClctDate VARCHAR2,
   p_MainCustId NUMBER, p_Option1 NUMBER, p_FinTime VARCHAR2, p_GroupCode VARCHAR2, 
   p_GroupName VARCHAR2, p_WorkerEn1 VARCHAR2, p_WorkerName1 VARCHAR2, 
   p_WorkerEn2 VARCHAR2, p_WorkerName2 VARCHAR2, p_SignEn VARCHAR2, 
   p_SignName VARCHAR2, p_Option2 NUMBER, p_UserId VARCHAR2, p_RetMsg out VARCHAR2) 
  RETURN number
  AS
  e_NoPR	exception;
 
  v_Cursor	number;			-- ��dynamic SQL
  v_RetCode	number;
  v_Command	VARCHAR2(2000);
  v_SubQuery	varchar2(2000) := '';
  v_Now		date;
 
  v_CustId	number;
  v_CustStatus	number;
  v_OldStatusCode number;
  v_Count	number := 0;
  v_Tmp		number;
  v_UpdTime	varchar2(19);
  v_UserName	varchar2(20);
  v_SNo		varchar2(9);
  v_InstCode	number;
  v_InstName	varchar2(12);
  v_NormalName  varchar2(12);
  v_PrCode1	number;
  v_PrName1	varchar2(12);
  v_PrCode2	number;
  v_PrName2	varchar2(12);
  v_ReasonCode	number;
  v_ReasonName	varchar2(12);
 
  v_CustName	varchar2(30);
  v_Tel1	varchar2(20);
  v_InstAddrNo	number;
  v_InstAddress varchar2(60);
  v_ServCode	varchar2(3);
  v_StrtCode	number;
 
begin
 
  -- �ˬd�Ѽ�
  if (p_MduId is null) or (p_CitemCode is null) or (p_CitemName is null) or 
   (p_Period is null) or (p_Amount is null) or (p_ClctDate is null) or 
   (p_MainCustId is null) or (p_UserId is null) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  if (p_Option1!=0 and p_Option1!=1) or (p_Option1=1 and p_FinTime is null) or 
    (p_Option2!=0 and p_Option2!=1) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- �ھڶǨӰѼơA�զ�����D�j�Ӥl��һݪ�SQL���O�G<SubQuery>
  v_SubQuery := 'select CustId, CustStatusCode, InstAddrNo from SO001 where MduId='
	||chr(39)||p_MduId||chr(39)||' and ChargeType!=3';
  if p_ClassSQL is not null then
    v_SubQuery := v_SubQuery || ' and ClassCode1 ' || p_ClassSQL;
  end if;
  if p_StatusSQL is not null then
    v_SubQuery := v_SubQuery || ' and CustStatusCode ' || p_StatusSQL;
  end if;
 
  -- �ھ�<p_UserId>��<���ʤH���W��>, �P<���ʮɶ�>
  select sysdate into v_Now from dual;
  v_UpdTime := GiPackage.GetDTString2(v_Now);
  begin
    select UserName into v_UserName from SO026 where UserId=p_UserId; 
  exception
    when no_data_found then
      p_RetMsg := 'No such user id.'||chr(39)||p_UserId||chr(39);
      return -6;
  end;
 
  /*
     Loop�C�@���D�j�Ӥl�᪺�Ȥ���:
     (1)�R���ӫȤ᪺�����g���ʶ��ظ��
     (2)��s�ӫȤ��Ȥ����ɤ������O�ݩʦ�"�j�ӲΦ�"�A�βM���򥻥x������
     (3)�Y p_Option1=1 �B�Y�Ӥᬰ����Ȥ�, �h:
        a.�X�۰ʵ��׸˾���: �Y�s�W�@���w��������줺�e���s�Ȥ�˾����
	b.�ର���`��: �N�ӵ��Ȥ��ƪ��Ȥ᪬�A�ର���`, �B��s�˾�����P�����u�渹�X
	c.��D�w�ˤ��+1, �j�Ӧ��Ĥ��+1, �j�Ӱ���/�ݩ���-1
     (4)�Y p_Option2=1 �B�Y�Ӥ�ݩ�Ȥ�, �h:
	a.�۰ʱN�����h��
	b.���ର���`��,
	c.��D�w�ˤ��+1, �j�Ӧ��Ĥ��+1, �j�Ӱ���/�ݩ���-1
  */
  if p_Option1 = 1 or p_Option2 = 1 then	  -- ��'Active'�Ȥ᪬�A�W��
    select Description into v_NormalName from CD035 where CodeNo=1; 
  end if;
  if p_Option1 = 1 then		-- �X�۰ʵ��׸˾���(����_����), ���������N�X�P�W�٭�
    v_InstCode := SF_GetCodeByRef('CD005', 5, v_InstName); -- 5=����_��
    if v_InstCode is null then
      p_RetMsg := 'No RefNo=5 record in CD005       ';
      return -2;
    end if;
  end if;
  if p_Option2 = 1 then		-- �۰ʱN�����h��, ���������N�X�P�W�٭�
    v_PrCode1 := SF_GetCodeByRef('CD007', 2, v_PrName1); -- 2=���
    v_PrCode2 := SF_GetCodeByRef('CD007', 5, v_PrName2); -- 5=��ú�O���
    if v_PrCode1 is null or v_PrCode2 is null then
      p_RetMsg := 'No RefNo=2 or 5 record in CD007             ';
      return -3;
    end if;
 
    v_ReasonCode := SF_GetCodeByRef('CD015', 2, v_ReasonName); -- 2=��Ƶ��P
    if v_ReasonCode is null then
      p_RetMsg := 'No RefNo=2 record in CD015       ';
      return -4;
    end if;
 
  end if;
 
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_SubQuery, DBMS_SQL.V7);
    DBMS_SQL.Define_Column(v_Cursor, 1, v_CustId);	-- �w�q���X�����
    DBMS_SQL.Define_Column(v_Cursor, 2, v_CustStatus);
    DBMS_SQL.Define_Column(v_Cursor, 3, v_InstAddrNo);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      p_RetMsg := 'SQL���O���~: ' || v_SubQuery;
      return -7;
  end;
 
  loop
    if DBMS_SQL.Fetch_rows(v_Cursor) > 0 then		-- Fetch 1 row
      DBMS_SQL.Column_Value(v_Cursor, 1, v_CustId);	-- ��CustId����
      DBMS_SQL.Column_Value(v_Cursor, 2, v_CustStatus);	-- ��CustStatusCode����
      DBMS_SQL.Column_Value(v_Cursor, 3, v_InstAddrNo); -- ��InstAddrNo����
 
      delete from SO003 where CustId=v_CustId and CitemCode=p_CitemCode;
      update SO001 set ChargeType=3, NextBCDate=null, UpdTime=v_UpdTime, UpdEn=v_UserName
	where CustId=v_CustId;
 
      -- �Y�Ӥᬰ����Ȥ�, �h�X�۰ʵ��׸˾���, ���ର���`��
      if p_Option1 = 1 and v_CustStatus = 3 then
	-- �����@�i�γ渹
	select S_SO007_SNo.NextVal into v_Tmp from dual;
	v_SNo := 'I'||ltrim(to_char(v_Tmp,'09999999'));	-- format: 'I99999999'
	-- ���Ȥ�����򥻸��
        select CustName, Tel1, InstAddress, ServCode 
	  into v_CustName, v_Tel1, v_InstAddress, v_ServCode
	  from SO001 where CustId=v_CustId;
        select StrtCode into v_StrtCode from SO014 where AddrNo=v_InstAddrNo;
 
	-- �s�W�@���۰ʵ��׸˾���
	insert into SO007 (CustId, SNo, AddrNo, Address, InstCode, InstName, ResvTime, 
		FinTime, AcceptTime, WorkerEn1, WorkerName1, WorkerEn2, WorkerName2, 
		SignEn, SignName, SignDate, AcceptEn, AcceptName, 
		ServCode, StrtCode, UpdTime, UpdEn, CustName, Tel1, 
		NextBCPeriod, NextBCAmount, GroupCode, GroupName, Note) 
	  values (v_CustId, v_SNo, v_InstAddrNo, v_InstAddress, v_InstCode, v_InstName, null,
		to_date(p_FinTime,'YYYYMMDDHH24MI'), to_date(to_char(v_Now,'YYYYMMDDHH24MI'),'YYYYMMDDHH24MI'),
		p_WorkerEn1, p_WorkerName1, p_WorkerEn2, p_WorkerName2, p_SignEn, p_SignName,
		to_date(to_char(v_Now,'YYYYMMDD'),'YYYYMMDD'), p_UserId, v_UserName, v_ServCode, v_StrtCode, 
		v_UpdTime, v_UserName, v_CustName, v_Tel1, p_Period, p_Amount, 
		p_GroupCode, p_GroupName, '�j�ӭӦ����Φ���, �t�Φ۰ʷs�W�����׸˾���');
 
	-- �N�ӵ��Ȥ��ƪ��Ȥ᪬�A�ର���`, �B��s�˾�����P�����u�渹�X
	update SO001 set CustStatusCode=1, CustStatusName=v_NormalName, 
	  InstTime=to_date(p_FinTime,'YYYYMMDDHH24MI'), InstSNo=v_SNo, OldStatusCode=3
	  where CustId=v_CustId;
 
	-- ��D�w�ˤ��+1, �j�Ӧ��Ĥ��+1, �j�Ӱ���/�ݩ���-1
	update CD017 set InstCnt=nvl(InstCnt,0)+1 where CodeNo=v_StrtCode;
	update SO017 set InstCnt=nvl(InstCnt,0)+1, UnInstCnt=nvl(UnInstCnt,0)-1 
	  where MduId=p_MduId;
      end if;
--DBMS_OUTPUT.PUT_LINE(v_CustId||' '||v_CustStatus);
      -- �Y�Ӥ�ݩ�Ȥ�, �h�۰ʱN�����h��, ���ର���`��
      if p_Option2 = 1 and v_CustStatus = 6 then
	-- ���ӱi����渹
	v_SNo := null;
	--select max(SNo) into v_SNo from SO009 where CustId=v_CustId and 
	--  FinTime is null and SignDate is null and (PrCode=2 or PrCode=5);
	select max(SNo) into v_SNo from SO009 where CustId=v_CustId and 
	  FinTime is null and SignDate is null and (PrCode=v_PrCode1 or PrCode=v_PrCode2);
 
	if v_SNo is not null then
	  select StrtCode into v_StrtCode from SO014 where AddrNo=v_InstAddrNo;
 
	  -- �N�����h��
	  update SO009 set ReasonCode=v_ReasonCode, ReasonName=v_ReasonName,
	    SignEn=p_SignEn, SignName=p_SignName, 
	    SignDate=to_date(to_char(v_Now,'YYYYMMDD'),'YYYYMMDD'), 
	    UpdEn=v_UserName, UpdTime=v_UpdTime
	    where SNo=v_SNo;
--DBMS_OUTPUT.PUT_LINE(v_CustId||', �N�����h��: '||v_SNo);
	  -- �R������������ulog�ɤ��ӵ����
	  delete from SO073 where CustId=v_CustId and SNo=v_SNo;
 
	  -- �N�ӵ��Ȥ��ƪ��Ȥ᪬�A�ର���`
	  update SO001 set CustStatusCode=1, CustStatusName=v_NormalName, OldStatusCode=6
	    where CustId=v_CustId;
	  -- �A��s���u���A�T
	  v_Tmp := SF_NewWip3(v_CustId);
 
	  -- ��D�w�ˤ��+1, �j�Ӧ��Ĥ��+1, �j�Ӱ���/�ݩ���-1
	  update CD017 set InstCnt=nvl(InstCnt,0)+1 where CodeNo=v_StrtCode;
	  update SO017 set InstCnt=nvl(InstCnt,0)+1, UnInstCnt=nvl(UnInstCnt,0)-1 
	    where MduId=p_MduId;
 
	else				-- �ݩ��o�L�����u��
	  raise e_NoPR;
	end if;
      end if;
 
      v_Count := v_Count + 1;
    else					-- No more rows to fetch
      exit;
    end if;
 
  end loop;
 
  DBMS_SQL.Close_cursor(v_Cursor);
 
  /*
    ��s�Φ��᪺������ơG
      �E�R���Φ��᪺�Ӷ��g���ʶ��ظ��
      �E�s�W�@���Ӷg���ʶ��ظ��
      �E��s�Φ����Ȥ����ɤ������O�ݩʦ�"�j�ӲΦ�"�ΰ򥻥x�����馨<p_ClctDate>
  */
  delete from SO003 where CustId=p_MainCustId and CitemCode=p_CitemCode;
  insert into SO003 (CustId, CitemCode, CitemName, Period, Amount, ClctDate, UpdTime, UpdEn)
	values (p_MainCustId, p_CitemCode, p_CitemName, p_Period, p_Amount, 
	to_date(p_ClctDate, 'YYYYMMDD'), v_UpdTime, v_UserName);
  update SO001 set ChargeType=3, NextBCDate=to_date(p_ClctDate, 'YYYYMMDD'),
	UpdEn=v_UserName, UpdTime=v_UpdTime where CustId=p_MainCustId;
 
 
  -- �վ�Ӥj�Ӱ򥻸�ƪ����e
  select CustName into v_CustName from SO001 where CustId=p_MainCustId;
  update SO017 set MainCustId=p_MainCustId, MainCustName=v_CustName, ClctMethod=1, 
    BCAmount=p_Amount, BCPeriod=p_Period, UpdTime=v_UpdTime, UpdEn=v_UserName 
    where MduId=p_MduId;
 
  commit;
  p_RetMsg := 'Execution OK, Cust. count='||to_char(v_Count);
  return 0;
 
exception
  when e_NoPR then
    DBMS_SQL.Close_cursor(v_Cursor);
    p_RetMsg := '�ݩ��o�L�����u��, �Ȥ�s��='||to_char(v_CustId);
    rollback;
    return -5;
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    p_RetMsg := 'Other error.';
    rollback;
    return -99;
end;
/
