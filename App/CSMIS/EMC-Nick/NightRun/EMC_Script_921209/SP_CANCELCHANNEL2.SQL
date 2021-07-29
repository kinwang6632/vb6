CREATE OR REPLACE PROCEDURE SP_CANCELCHANNEL2 (p_RetCode Out Number, p_RetMsg Out Varchar2)IS


/*

  --@C:\SP_CANCELCHANNEL2.sql
    variable p_RetCode number
    variable p_RetMsg varchar2(100)
    exec  SP_CANCELCHANNEL2(:p_RetCode,:p_RetMsg);
    print P_Retcode
    print p_RetMsg

  目的: 將已經超過應收日期數日卻仍未繳費的付費頻道關閉
  檔名: SP_ChannelClose.sql



  OUT p_RetMsg  varchar2: 處理結果自串

      p_RetCode number: 處理結果代碼, 對應訊息存於p_RetMsg中

              0: p_RetMsg='欠費自動關閉執行完畢'
             -1: p_RetMsg='INSERT 時發生錯誤
             -2: P_RetMsg='Delete時發生錯誤
             -3: p_RetMsg='Insert至SOAC0202時發生錯誤
            -99: p_RetMsg='其他錯誤'

  流程:
	Loop所有使用中的STB

		        新增一筆關頻道指令(B2)至共用資料區的Send_Nagra中, 欄位如下:
			High_Level_Cmd_ID: 'B2'
			ICC_NO: 該設備之ICC卡號
			Notes: 待關頻道之Product id. list, 以','區隔
                        CMD_STATUS:'W'
                        OPERATOR:'批次作業'
                        UPDTIME:SYSDATE
                        SEQNO:從 sequence object S_SendNagra_SeqNo 中select 出NextValue
                        COMPCODE:該客戶所屬之公司別代碼

                        刪除客戶頻道資料檔(SO005)被關頻道的資料

                        紀錄一筆資料至STB/ICC設定記錄檔(SOAC0202),紀錄客戶被關頻道的資料



  定義:
	. 所有使用中的STB: 從SO033(正式應收檔)找出有收看付費頻道的客戶,未收原因代碼不是空的,
          CancelFlag(作廢)=0,而該客戶的應付費日期加上SO043(收費參數檔)的欠費日小於系統日,
          SO002(客戶各服務基本資料檔)CustStatusCode in (1，2，5，6) 者,CD019(收費項目代碼檔)
          的參考號(RefNo)為2者,SO004(客戶設備資料檔)FaciSNO(物品序號)不為空者且該設備的拆除
          日期為空者(PRDate is null)

  Date: 2003/02/12
  By:Morris
*/


V_Compcode                   Number(3);
V_CustID                     Number(8);
V_CvtsNo                     Varchar2(14);
V_Chcode                     Varchar2(3);
V_ChName                     Varchar2(26);
V_Smartcardno                Varchar2(15);
V_ChannelID                  Varchar2(12);
V_Smartcardno1               Varchar2(15);
V_ChannelID1                 Varchar2(1000);
V_TIME                       Date;
s_SQL                        Varchar2(4000);
V_Rowcount                   Number(3):=0;
V_HIGHT_LEVEL_CMD_ID         Varchar2(5);
V_CMD_STATUS                 Varchar2(2);
V_OPERATOR                   Varchar2(15);
V_count                      Number(3):=0;
V_count1                     Number(3):=0;
V_count2                     Number(3):=0;
V_ModeType                   Varchar2(5);
V_UPDEN                      Varchar2(20);

Cursor n_close is
  Select distinct(f.channelid)ChannelID,a.compcode CompCode,a.custid CustID,
                  h.smartcardno Smartcardno,h.cvtsno CvtsNo,f.codeno Chcode,
                  f.description ChName
     from so033 a,so002 b,cd019 c,cd019a e,cd024 f,so004 g,so005 h,so043 i
    where (a.shoulddate+i.para21)<sysdate
      and a.custid=b.custid
      and a.compcode=b.compcode
      and a.custid=g.custid
      and a.custid=h.custid
      and b.custstatuscode in (1,2,5,6)
      and a.citemcode=c.codeno
      and c.refno='2'
      and a.uccode is not null
      and a.cancelflag='0'
      and c.codeno=e.citemcode
      and f.codeno=e.codeno
      and g.prdate is null
      and g.facisno is not null
      and g.smartcardno is not null
      and h.cvtsno=g.facisno
      and h.smartcardno=g.smartcardno
      order by h.smartcardno,f.channelid;

BEGIN

  V_HIGHT_LEVEL_CMD_ID:='B2';
  V_CMD_STATUS:='W';
  V_OPERATOR:='批次作業';
  V_ModeType:='B2';
  V_TIME:=SYSDATE;
  V_UPDEN:='批次作業';
  

  
  for n_record in n_close loop 

    V_CompCode:=n_record.CompCode;  
    V_CustID:=n_record.CustID;
    V_Cvtsno:=n_record.Cvtsno;
    V_ChName:=n_record.ChName;  
  
    
    if V_Smartcardno1=n_record.Smartcardno then
       V_ChannelID1:=to_char(to_number(V_ChannelID1))||','||to_char(to_number(n_record.ChannelID));
    else
      begin      
       if n_close%rowcount!=1 then
          Insert into Com.Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,NOTES,CMD_STATUS,
                 OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES)
                          Values(V_HIGHT_LEVEL_CMD_ID,V_Smartcardno1,V_ChannelID1,V_CMD_STATUS,
                 V_OPERATOR,V_TIME,Com.S_SENDNAGRA_SEQNO.NEXTVAL,n_record.CompCode,'');

                 V_count:= V_count+1;
       end if;

                 V_Smartcardno1:=n_record.Smartcardno;
                 V_ChannelID1:=to_char(to_number(n_record.ChannelID));

          exception
            when others then
             p_RetMsg := 'INSERT時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
             p_RetCode := -1;
         end;
     
    end if;
     
      begin
        Delete SO005 Where CustID=n_record.CustID
                       and Smartcardno=V_Smartcardno1
                       and Chcode=n_record.Chcode;
       exception
            when others then
             p_RetMsg := 'Delete時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
             p_RetCode := -2;
         end;
        
       begin
         Insert into SOAC0202(CompCode,CustId,STBSNo,
                             SmartCardNo,ModeType,ChCode,ChName,
                             UpdTime,UpdEn)
                     Values(n_record.Compcode,n_record.CustID,n_record.Cvtsno,
                            V_Smartcardno1,V_ModeType,n_record.Chcode,n_record.ChName,
                            V_Time,V_UPDEN);               
          exception
            when others then
             p_RetMsg := 'INSERT至SOAC0202時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
             p_RetCode := -3;
         end;

    end loop;
      
      begin    
        Insert into Com.Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,NOTES,CMD_STATUS,
           OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES)
                       Values(V_HIGHT_LEVEL_CMD_ID,V_Smartcardno1,V_ChannelID1,V_CMD_STATUS,
          V_OPERATOR,V_TIME,Com.S_SENDNAGRA_SEQNO.NEXTVAL,V_CompCode,'');

          V_count1:=V_count1+1;

        exception
         when others then
          p_RetMsg := 'INSERT時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
          p_RetCode := -1;
       end;

     begin
        Delete SO005 Where CustID=V_CustID
                       and Smartcardno=V_Smartcardno1
                       and Chcode=V_Chcode;
        exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'Delete時發生錯誤: '||SQLERRM;
           p_RetCode:= -2;                                             
       end;
        
     begin
         Insert into SOAC0202(CompCode,CustId,STBSNo,
                             SmartCardNo,ModeType,ChCode,ChName,
                             UpdTime,UpdEn)
                     Values(V_Compcode,V_CustID,V_CvtsNo,
                            V_Smartcardno1,V_ModeType,V_Chcode,V_ChName,
                            V_Time,V_UPDEN);               
          exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'INSERT至SOAC0202時發生錯誤: '||SQLERRM;
           p_RetCode:= -3;                                             
        end; 

             V_count2:=V_count+V_count1;

        commit;

      p_RetMsg:='欠費自動關閉執行完成,共產生 ' || V_count2 ||' 筆待關頻道之資料 ';
      p_RetCode := 0;

 EXCEPTION
    WHEN others THEN
     rollback ;
     p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||s_SQL;
     P_RetCode:=-99;

END ; -- Procedure CHANNELCLOSE
--Version1.1
--Modify by Morris 2003/02/12
