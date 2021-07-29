/*
  ADD_SEQ.SQL: Create sequence objects
  Naming rule: S_table_column
  若為正式版, 則可修改本檔, 使序號從 1 開始
  Date: 2004.03.08
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立Sequence objects >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\SGS\Script\Add_Seq.log
print str

prompt S_SGSMSG_SeqNo
-- SMS產生唯一的消息編碼 (format: N(10))
drop sequence S_SGSMSG_SeqNo;
create sequence S_SGSMSG_SeqNo
       MINVALUE 1
       MAXVALUE 9999999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       NOCYCLE;


-- 產品訂購資訊查詢中,取消訂購的頻道中的『訂單編號』 (format: N(6))
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
