/*
  ADD_SEQ.SQL: Create sequence objects
  Naming rule: S_table_column
  �Y��������, �h�i�ק糧��, �ϧǸ��q 1 �}�l
  Date: 2004.03.08
*/
set heading off
variable str varchar2(80)
exec :str := '<< �إ�Sequence objects >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\SGS\Script\Add_Seq.log
print str

prompt S_SGSMSG_SeqNo
-- SMS���Ͱߤ@�������s�X (format: N(10))
drop sequence S_SGSMSG_SeqNo;
create sequence S_SGSMSG_SeqNo
       MINVALUE 1
       MAXVALUE 9999999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       NOCYCLE;


-- ���~�q�ʸ�T�d�ߤ�,�����q�ʪ��W�D�����y�q��s���z (format: N(6))
drop sequence S_SGSORDER_SeqNo;
create sequence S_SGSORDER_SeqNo
       MINVALUE 1
       MAXVALUE 999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       CYCLE;

spool off
set heading on
