prompt S_SEND_NDS_SeqNo
-- CA Command Gateway���������O�����.SeqNo SEND_NDS.SEQNo (format: 99999999)
-- SELECT MAX(SEQNO)+1 FROM SEND_NDS;
drop sequence S_SEND_NDS_SeqNo;
create sequence S_SEND_NDS_SeqNo
       MINVALUE 1
       MAXVALUE 99999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       CYCLE;


