
CREATE OR REPLACE FUNCTION SF_GetDataArea
  (p_UserID IN VARCHAR2) RETURN VARCHAR2
  AS
  v_TotalAreaCnt number;
  v_UserAreaCnt number;
  v_Tmp varchar2(100) := '';
  v_Flag boolean;

  -- �Ӿާ@�̥i�ϥΪ���ư�
  cursor CC1 is 
      select ServCode from SO028 where UserId=p_UserID;

begin
  -- �Ѽ��ˬd
  if p_UserId is null then
    return '������~';
  end if;

  -- ���ӭ��i�ϥΰϤ���Ƶ���
  select count(*) into v_UserAreaCnt from SO028 where UserId=p_UserId;

  -- �Y�ӭ��i�ϥΰϤ���Ƶ���=0, �h�i�ϥΩҦ���ư�
  if v_UserAreaCnt = 0 then
    return '';
  end if;

  -- �Y�ӭ��i�ϥΰϤ���Ƶ���=1, �h�i�ϥάY�@��
  if v_UserAreaCnt = 1 then
    select ServCode into v_Tmp from SO028 where UserId=p_userId;
    return 'ServCode=' || chr(39) || v_Tmp || chr(39);
  end if;

  -- ���Ҧ���ưϤ���Ƶ���
  select count(*) into v_TotalAreaCnt from CD002;

  -- �Y�ӭ��i�ϥΰϤ���Ƶ��Ƥ�Ҧ���ưϵ��Ʈt1��, �h���i�ϥάY�@��
  if (v_TotalAreaCnt - v_UserAreaCnt) = 1 then
    select CodeNo into v_Tmp from CD002 where CodeNo not in
      (select ServCode from SO028 where UserId=p_UserId);
    return 'ServCode!=' || chr(39) || v_Tmp || chr(39);
  end if;

  -- �i�ϥάY�X��
  v_Tmp := 'ServCode in (';
  for cr1 in CC1 loop
    v_Tmp := v_Tmp || chr(39) || cr1.ServCode || chr(39) || ',';
  end loop;
  v_Tmp := rtrim(v_Tmp,',') || ')';
  return v_Tmp;

exception
  when others then
    return '��ưϿ��~';
end;
/