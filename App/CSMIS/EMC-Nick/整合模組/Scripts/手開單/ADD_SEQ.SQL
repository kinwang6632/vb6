/*
  ADD_SEQ.SQL: Create sequence objects
  Naming rule: S_table_column
  若為正式版, 則可修改本檔, 使序號從 1 開始
  Date: 2002.09.03
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立Sequence objects >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool p:\gird\Log\Add_Seq.log
print str



prompt S_So126_SeqNo
-- 手開單之領單資料檔  So126.Seq (format: N(10))
drop sequence S_SO126_Seq;
create sequence S_SO126_Seq
       MINVALUE 1
       MAXVALUE 9999999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       NOCYCLE;
spool off
set heading on
