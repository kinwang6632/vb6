/*
    �ɥR��h���Ʀ�q���W�u�һݥN�X, ���ܦUSO��ưϪ������榹�ɮקY�i
    Date: 2003.01.16
    By: Pierre
*/

-- �����q�O
variable v_CompCode number
begin
    select CodeNo into :v_CompCode from CD039 where rownum<=1;
end;
/

-----------------------------------------------------------------------
prompt CD061 STB���M��]�N�X
delete from CD061;
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 1, '�h�a���ݤF', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 2, '�`�ذ��D', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 3, '���Q��', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 4, '��L', 0);


-----------------------------------------------------------------------
prompt CD060 �W�D���M��]�N�X
delete from CD060;
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 1, '�`�ذ��D', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 2, '�O�ΰ��D', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 3, '���Q��', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 4, '��L', 0);


-----------------------------------------------------------------------
prompt CD028 STB�t��N�X��
delete from CD028;
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 1, '���z�d', 800, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 2, '������', 500, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 3, '�q���u', 200, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 4, 'AV�u', 200, 0);


-----------------------------------------------------------------------
prompt CD057 �q�����O�N�X��
delete from CD057;
insert into CD057 (CodeNo, Description, RefNo, StopFlag)
    values (1, '�Ӹ�eBox', 10, 0);
insert into CD057 (CodeNo, Description, RefNo, StopFlag)
    values (2, '�I�O�W�D����', 9, 0);


-----------------------------------------------------------------------
prompt CD051 �@�o��]�N�X��
declare
  v_Cnt number := 0;
  v_MaxCodeNo number := 0;
begin
  select count(*) into v_Cnt from CD051 where RefNo=1;
  select max(nvl(CodeNo,0)) into v_MaxCodeNo from CD051;

  if v_Cnt =0 then
    insert into CD051 (CompCode, CodeNo, Description, RefNo, StopFlag)
	values (:v_CompCode, v_MaxCodeNo+1, '�W�D���M', 1, 0);
  end if;
end;
/

