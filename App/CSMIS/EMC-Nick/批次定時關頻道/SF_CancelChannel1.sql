CREATE OR REPLACE FUNCTION SF_CANCELCHANNEL1(p_CompCode number, p_ProdIdList varchar2,
  p_ChcodeList varchar2,p_ExecTime varchar2, p_RetMsg out varchar2) return number
  as

/*
--@c:\SF_CancelChannel1
variable nn number
variable msg varchar2(100)
exec :nn := SF_CancelChannel1(7, '12,13', '200304010010', :msg);
print nn
print msg

  批次作業: 新增一批關頻道指令, 以做到於某一時間到時, 可關閉這些付費頻道
  目的: 可針對所有使用中的STB, 關閉試看至某一日期的頻道
  檔名: SF_CancelChannel1.sql

  注意: 此函式目的設計成當92.03.31, EMC會要求我們將所有STB試看至該日的情色頻道組全部關掉,
	但我們的高階功能並未提供試看至某日的功能(只有試看固定天數的功能), 做的時候, EMC
	全部系統台都要執行一次.  此函式也可以讓我們刻意關掉某些頻道.

  IN p_CompCode number: 公司別 
     p_ProdIdList varchar2: Nagra product id. list, 以','區隔, 不要有空白
     p_ExecTime varchar2: 關頻道時間, 西曆時間10碼字串, 'YYYYMMDDHH24MI', 不含秒
  OUT p_RetMsg varchar2: 處理結果自串
  Return NUMBER：處理結果代碼
	>=0:執行完畢
	   -1:INSERT時發生錯誤
                     -2:Delete時發生錯誤
                     -3:INSERT至SOAC0202時發生錯誤　　　　
                   -99:其他錯誤

  流程:
	Loop所有使用中的STB
		    新增一筆關頻道指令(B2)至共用資料區的Send_Nagra中, 欄位如下:
				High_Level_Cmd_ID: 'B2'
			                   ICC_NO: 該設備之ICC卡號
			                   Notes: 待關頻道之Product id. list, 以','區隔
                                                                           CMD_STATUS:'W'
                                                                           OPERATOR:'NIGHT-RUN'
                                                                           UPDTIME:SYSDATE
                                                                           SEQNO:從 sequence object S_SendNagra_SeqNo 中select 出NextValue
                                                                           COMPCODE:該客戶所屬之公司別代碼
                                                                           ProcessingDate:處理日期

                                         刪除客戶頻道資料檔(SO005)被關頻道的資料

                                         紀錄一筆資料至STB/ICC設定記錄檔(SOAC0202),紀錄客戶被關頻道的資料

  定義: 
	. 所有使用中的STB: 設備資料檔(SO004)中其品名代碼(FaciCode), 對應至品名代碼
	  檔(CD022)的參考號(RefNo)為3(即STB)者, 且該設備的拆除日期為空者(PRDate is null)

  Date:2003/01/23
  By: Morris
*/ 
 
I		                number;
v_Index	                                   number;
v_ChcodeList                             varchar2(30);
v_ChcodeCnt	                 number;
v_FirstChcode                            varchar2(20);
s_SQL                                          Varchar2(4000);
V_HIGHT_LEVEL_CMD_ID    Varchar2(5);
V_CMD_STATUS                     Varchar2(2);
V_OPERATOR                           Varchar2(15);
V_count                                       Number(3):=0;
V_ModeType                              Varchar2(5);
V_Time                                         Date;
V_ChName                                  Varchar2(20);
TYPE Varchar2Ary IS TABLE OF varchar2(20) INDEX BY BINARY_INTEGER;
ChcodeName Varchar2Ary;	-- array for Chcode

Cursor n_cancel is
 select distinct (a.smartcardno) smartcardno,a.compcode compcode,
                  a.custid CustID,b.cvtsno Cvtsno
  from so004 a,so005 b,cd022 c,cd024 d,so002 e
  where a.custid=b.custid
    and a.smartcardno=b.smartcardno   
    and a.facicode=c.codeno
    and c.refno='3'
    and a.prdate is null
    and b.chcode=d.codeno
    and a.compcode=p_CompCode
    and a.custid=e.custid
    and e.custstatuscode in (1,2,5,6)
 order by a.smartcardno;
        
BEGIN
  
  V_HIGHT_LEVEL_CMD_ID:='B2';
  V_CMD_STATUS:='W';
  V_OPERATOR:='NIGHT_RUN';
  V_ModeType:='B2';
  V_Time:=sysdate;  
        
    --切割Chcode字串至Chcode陣列中
     v_ChcodeCnt:= 0;
  if p_ChcodeList is not null then
     v_ChcodeList := p_ChcodeList;
     v_Index := 1;
     I := 1;
    while v_ChcodeList is not null loop
     v_Index := instr(v_ChcodeList, ',');
      if v_Index > 0 then
	begin
      ChcodeName(I) := ltrim(rtrim(substr(v_ChcodeList, 1, v_Index-1)));
	  v_ChcodeList := substr(v_ChcodeList, v_Index+1);
      I:=I+1;     
    exception
	  when others then
	    v_ChcodeList := null;
	end;   
      else
        ChcodeName(I) := rtrim(ltrim(v_ChcodeList));
        v_ChcodeList := null;     
      end if; 
           
    end loop;
    
    v_ChcodeCnt := I;
    v_FirstChcode := ChcodeName(1); -- 記住第一個Chcode
    
  end if; 
       
    for n_record in n_cancel loop
    
      begin               
        Insert into Com. Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,NOTES,CMD_STATUS,
               OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES,ProcessingDate)
                        Values(V_HIGHT_LEVEL_CMD_ID,n_record.smartcardno,n_record.Cvtsno,
                               p_ProdIdList,V_CMD_STATUS,V_OPERATOR,V_Time,Com.S_SENDNAGRA_SEQNO.NEXTVAL,n_record.compcode,'',to_date(p_ExecTime,'YYYYMMDDHH24MI'));
         
                    V_count:= V_count+1; 

        exception
          When OTHERS Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'INSERT時發生錯誤: '||SQLERRM;
          return -1;                                
        end;
   
      for I in 1..v_ChcodeCnt loop 
      
      Select Description into V_ChName  from Com.CD024 Where CodeNo=ChcodeName(I);
  
     begin
        Delete  Com.SO005 Where CustID=n_record.CustID
                                                and Smartcardno=n_record.smartcardno 
                                                and Chcode=ChcodeName(I);
        exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'Delete時發生錯誤: '||SQLERRM;
           return -2;                                             
        end;
      
      begin
             Insert into  Com.SOAC0202(CompCode,CustId,STBSNo,
                            SmartCardNo,ModeType,ChCode,ChName,
                           UpdTime,UpdEn)
                     Values(n_record.Compcode,n_record.CustID,n_record.Cvtsno,
                            n_record.Smartcardno,V_ModeType,ChcodeName(i),V_ChName,
                        V_Time,'系統自動');               
          exception
          When Others Then
           DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
           p_RetMsg := 'INSERT至SOAC0202時發生錯誤: '||SQLERRM;
           return -3;                                             
        end; 
              
      end loop;
      
    end loop;
          
    
     commit;
      
      p_RetMsg:='自動關閉頻道執行完成,共產生 ' || V_count ||' 筆待關頻道之資料 ';
      return 0;

 EXCEPTION
  WHEN OTHERS THEN
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := '其他錯誤: '||SQLERRM;
    rollback;
    return -99;

END ; 

/
--modify by morris 2003/01/30
--Version 1.1

