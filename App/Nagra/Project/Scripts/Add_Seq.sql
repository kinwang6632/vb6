prompt S_SendNagra_SeqNo
-- CA Command Gateway之高階指令資料檔.SeqNo Send_Nagra.SEQNo (format: 99999999)
-- SELECT MAX(SEQNO)+1 FROM Send_Nagra;
drop sequence S_SendNagra_SeqNo;
create sequence S_SendNagra_SeqNo
       MINVALUE 1
       MAXVALUE 999999
       INCREMENT BY 1
       START WITH 1
       NOCACHE
       CYCLE;


