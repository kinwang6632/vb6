CREATE OR REPLACE FUNCTION SF_PRODPURCHASEQUERY
  (p_CompCode varchar2,p_SrcCode varchar2,
   p_DstCode varchar2,p_DstMsgID varchar2,p_CasID varchar2,p_InstQuery varchar2,
   p_FirstDate varchar2,p_HandleEDateTime varchar2,p_CancelChStr varchar2,
   p_SrcMsgID OUT varchar2,p_ErrCode OUT varchar2,p_ErrMsg OUT varchar2,
   p_TotalExecCounts OUT number,p_FirstFlag OUT varchar2)
  RETURN  number AS
/*
--@D:\App\SGS\Script\SF_ProdPurchaseQuery;
VARIABLE ReturnNo NUMBER
VARIABLE ErrCode NUMBER
VARIABLE ErrMsg VARCHAR2(1000)
VARIABLE SrcMsgID VARCHAR2(30)
EXEC :ReturnNo := SF_ProdPurchaseQuery('7','110000S01','110000G01','','0x4a01','0','2003/03/24','2003/03/25 23:59:59','''B1'',''B2''',:SrcMsgID,:ErrCode,:ErrMsg)
PRINT ReturnNo
PRINT ErrCode
PRINT ErrMsg
PRINT SrcMsgID


  產品定購資訊查詢應答介面用於服務平臺將數位電視用戶購買數位電視?品的變化資訊返回給監管平臺
  檔名: SF_ProdPurchaseQuery.sql

  說明:


  IN   p_CompCode         :公司別
       p_SrcCode          :SMS編碼
       p_DstCode          :監管中心/分中心編碼
       p_DstMsgID         :應答資料對應的監管請求的MSGID
       p_CasID            :CAS的ID
       p_InstQuery        :是否即時查詢的返回結果(0：非即時查詢返回結果  1：即時查詢結果)
       p_FirstDate        :第一次產生完整資料的日期 YYYY/MM/DD
       p_HandleEDateTime  :要產生資料的結束時間 YYYY/MM/DD HH24:MI:SS
       p_CancelChStr      :要關閉的頻道指令代碼(目前訂購頻道查詢是不包含免費頻道及試看頻道)



  OUT  p_ErrCode         number:   結果訊息
       p_ErrMsg          varchar2: 結果訊息

       Return            number:   處理結果代碼, 對應訊息存於p_RetMsg中
                             =0: p_RetMsg='',傳回空值,代表正常執行完畢
                             -1: p_RetMsg=代表下SQL指令時有Exception
                             -2: p_RetMsg=代表
                            -99: p_RetMsg='其他錯誤'

    By: Jackal
    Date: 2004.03.12
    2004.03.18 修改 :查詢Log,改由前端處理完成後Log至SO460
    2004.03.19 修改 :查詢Log,改由前端處理完成後Log至SO459
    2004.04.05 修改 :SaveDataToDB有Exception時回傳相關欄位
*/
  v_SQL              	varchar2(4000);		-- must long enough

  v_CurrDateTime      Date;
  v_FirstDate         varchar2(20);


  v_Counts            Number := 0;
  v_TotalExecCounts   Number := 0;


  v_FirstFlag         Number;
  v_ModeType          Number;
  v_RealDateTime      varchar2(20);
  v_ExecDateTime      varchar2(20);
  v_SmsMsgSeq         Number;
  v_SmsOrderSeq       Number;
  v_ProductCode       varchar2(20);
  v_PurchaseID        varchar2(20);
  v_ICCardNo          varchar2(20);
  v_ICCardID          varchar2(20);
  v_SetTime           varchar2(20);
  v_DueDate           varchar2(20);
  v_Status            Number;
  v_OrderTitle        varchar2(20);
  v_CancelChStr       varchar2(50);



  v_StartDateTime     varchar2(20);--非第一次查詢的某日開始時間
  v_StopDateTime      varchar2(20);--非第一次查詢的某日結束時間

  v_Table             varchar2(10);

  v_ErrMsg            varchar2(2000);
  
  type CancelCh is ref cursor;	--自訂cursor型態
  cDynamic CancelCh;          	--供dynamic SQL



  cursor ccFirstData is--第一次查詢完整資料
    Select A.OrderNo,A.SmartCardNo,A.ChCode,C.ChannelID AS ProductCode,A.SetTime,
           A.DueDate,B.AcceptTime from SO005 A,SO105 B,CD024 C
    Where A.CompCode=p_CompCode And A.ChCode=C.CodeNo
    And A.OrderNo=B.OrderNo And C.PayFlag<>0;



  cursor ccDiffDataA is--某日新訂購的頻道
    Select A.OrderNo,A.SmartCardNo,A.ChCode,C.ChannelID AS ProductCode,A.SetTime,
           A.DueDate,B.AcceptTime from SO005 A,SO105 B,CD024 C
    Where A.CompCode=p_CompCode And A.ChCode=C.CodeNo
    And A.OrderNo=B.OrderNo And C.PayFlag<>0 And B.AcceptTime Between
    To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')AND
    To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');



--將資料儲存置SO455或SO456
  FUNCTION SaveDataToDB(I_Table varchar2,I_CompCode Number,I_HandleEDateTime varchar2,
                        I_RealDateTime varchar2,I_ExecDateTime varchar2,
                        I_Status NUMBER,I_PurchaseID varchar2,I_ICCardNo varchar2,
                        I_ICCardID varchar2,I_CasID varchar2,I_ProductCode varchar2,
                        I_StartTime varchar2,I_EndTime varchar2)
  RETURN varchar2
  AS
    L_SQL varchar2(4000);
    L_ErrorMsg varchar2(1000);
  BEGIN
    BEGIN


      L_SQL := 'Insert Into '||I_Table||' (CompCode,HandleEDateTime,RealDateTime,'||
               'ExecDateTime,Status,PurchaseID,ICCardNo,ICCardID,CasID,ProductCode,'||
               'StartTime,EndTime) '||
               'Values('||I_CompCode||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_RealDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_Status||','''||I_PurchaseID||''','''||I_ICCardNo||''','''||
               I_ICCardID||''','''||I_CasID||''','''||I_ProductCode||''','||
               'TO_DATE('''||I_StartTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_EndTime||''',''YYYY/MM/DD HH24:MI:SS''))';
               
               
               
      L_ErrorMsg := '  ErrorMsg:Table='||I_Table||
                    '  CompCode='||I_CompCode||
                    '  HandleEDateTime='||I_HandleEDateTime||
                    '  RealDateTime='||I_RealDateTime||
                    '  ExecDateTime='||I_ExecDateTime||
                    '  Status='||I_Status||
                    '  PurchaseID='||I_PurchaseID||
                    '  ICCardNo='||I_ICCardNo||
                    '  ICCardID='||I_ICCardID||
                    '  CasID='||I_CasID||
                    '  ProductCode='||I_ProductCode||
                    '  StartTime='||I_StartTime||
                    '  EndTime='||I_EndTime;
                    
                    
      EXECUTE IMMEDIATE L_SQL;


      RETURN Null;
    EXCEPTION
      WHEN others THEN
        ROLLBACK;--之前做的先回復

        --RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
        RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||L_ErrorMsg;
    END;
  END;




--將資料LOG至SO459
  PROCEDURE SaveSO459Log(I_CompCode Number,I_ModeType Number,I_FirstFlag Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --查詢Log改由前端處理完成後Log
    L_SQL := ' ';
/*
    IF I_ErrorMsg IS NULL THEN
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''))';
    ELSE
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime,ErrorCode,ErrorMsg) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'''||
               I_ErrorCode||''','''||I_ErrorMsg||''')';
    END IF;
    EXECUTE IMMEDIATE L_SQL;
    COMMIT;
*/
  END;





--將資料LOG至SO460
  PROCEDURE SaveSO460Log(I_CompCode Number,I_ModeType Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_InstQuery Number,I_SrcCode varchar2,I_SrcMsgID varchar2,
                        I_DstCode varchar2,I_DstMsgID varchar2,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --查詢Log改由前端處理完成後Log
    L_SQL := ' ';
/*
    IF I_ErrorMsg IS NULL THEN
      L_SQL := 'Insert Into SO460 (CompCode,ModeType,HandleEDateTime,'||
               'ExecDateTime,InstQuery,SrcCode,SrcMsgID,DstCode,DstMsgID) '||
               'Values('||I_CompCode||','||I_ModeType||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_InstQuery||','''||I_SrcCode||''','''||I_SrcMsgID||''','''||
               I_DstCode||''','''||I_DstMsgID||''')';
    ELSE
      L_SQL := 'Insert Into SO460 (CompCode,ModeType,HandleEDateTime,'||
               'ExecDateTime,InstQuery,SrcCode,SrcMsgID,DstCode,DstMsgID,ErrorCode,ErrorMsg) '||
               'Values('||I_CompCode||','||I_ModeType||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_InstQuery||','''||I_SrcCode||''','''||I_SrcMsgID||''','''||
               I_DstCode||''','''||I_DstMsgID||''','''||
               I_ErrorCode||''','''||I_ErrorMsg||''')';
    END IF;
    EXECUTE IMMEDIATE L_SQL;
    COMMIT;
*/
  END;
  
  
--拆解CancelChannel
  FUNCTION ChangeCancelChStr(I_CancelChStr varchar2)
  RETURN varchar2
  AS
    v_TempCancelChStr varchar2(50);
    v_CancelChStr varchar2(50);
    v_Temp varchar2(10);
    v_Index number;
    I number;
    TYPE Varchar2Ary IS TABLE OF varchar2(50) INDEX BY BINARY_INTEGER;
    CancelChStr Varchar2Ary;
  BEGIN

    I := 0;
    v_TempCancelChStr := I_CancelChStr;
    v_CancelChStr := ' ';
    
    WHILE v_TempCancelChStr IS NOT NULL LOOP
      v_Index := instr(v_TempCancelChStr, ',');
      IF v_Index > 0 THEN
        v_Temp := ltrim(rtrim(substr(v_TempCancelChStr, 1, v_Index-1)));
        
        IF v_CancelChStr = ' ' THEN
          v_CancelChStr := ''''||v_Temp||'''';
        ELSE
          v_CancelChStr := v_CancelChStr||','''||v_Temp||'''';
        END IF;             	
        
        v_TempCancelChStr := substr(v_TempCancelChStr, v_Index+1);
      ELSE
        v_Temp := rtrim(ltrim(v_TempCancelChStr));

        IF v_CancelChStr = ' ' THEN
          v_CancelChStr := ''''||v_Temp||'''';
        ELSE
          v_CancelChStr := v_CancelChStr||','''||v_Temp||'''';
        END IF;

        v_TempCancelChStr := NULL;        	
      END IF;
  	END LOOP;
  	
  	
   RETURN v_CancelChStr;
  END;


BEGIN
  --取出系統時間
  --SELECT To_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') INTO v_CurrDateTime FROM dual;
  v_CurrDateTime := sysdate;
  v_ExecDateTime := To_Char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');
  v_FirstDate := p_FirstDate;
  p_SrcMsgID := ' ';
  p_FirstFlag := ' ';

  --ModeType=1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
  v_ModeType := 3;


  --判別第一次完整資料(CAProductQuery)是否已經做過
  SELECT COUNT(*) INTO v_Counts FROM SO459
    WHERE CompCode=p_CompCode And FirstFlag='1' And ErrorCode IS NULL And ModeType=v_ModeType;



  IF v_Counts=0 THEN
--**********************************************************************
--* * * * * * 歷史資料(第一次資料)查出未拆機的所有IC卡資料 * * * * * * *
--**********************************************************************
    v_FirstFlag := 1;--執行第一次資料
    p_FirstFlag := 1;


    IF To_Date(To_Char(v_CurrDateTime,'YYYY/MM/DD'),'YYYY/MM/DD') > To_Date(v_FirstDate,'YYYY/MM/DD') THEN
      --檢核處理日期是否大於第一天
      IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') = To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN

        FOR crFirstData in ccFirstData loop
          --第一次執行皆用新訂購的頻道
          v_RealDateTime := To_Char(crFirstData.SetTime,'YYYY/MM/DD HH24:MI:SS');
          v_Status := 0;--第一次歷史資料皆為New(0:Modify  1:Delete)
          v_PurchaseID := crFirstData.OrderNo;
          v_ICCardNo := crFirstData.SmartCardNo;
          v_ICCardID := '';--目前沒有
          v_ProductCode := crFirstData.ProductCode;
          v_SetTime :=To_Char(crFirstData.SetTime,'YYYY/MM/DD HH24:MI:SS');
          v_DueDate :=To_Char(crFirstData.DueDate,'YYYY/MM/DD HH24:MI:SS');


          v_ErrMsg := SaveDataToDB('SO455',p_CompCode,p_HandleEDateTime,
                        v_RealDateTime,v_ExecDateTime,
                        v_Status,v_PurchaseID,v_ICCardNo,
                        v_ICCardID,p_CasID,v_ProductCode,
                        v_SetTime,v_DueDate);


          IF v_ErrMsg Is Not Null THEN
            p_ErrCode := -1;
            p_ErrMsg := SubStr(v_ErrMsg,1,2000);
            p_TotalExecCounts := 0;


            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

            RETURN -1;
            
            
          ELSE
            v_TotalExecCounts := v_TotalExecCounts + 1;
          END IF;

        END LOOP;



      --NightRun正常執行完畢Log
      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,NULL,NULL);
      ELSE
        p_ErrCode := -2;
        p_ErrMsg := 'Return_Msg_104_Content';--尚未執行第一次資料
        p_TotalExecCounts := 0;

        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,p_ErrCode,p_ErrMsg);

        RETURN -2;
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_102_Content';--未到執行第一次時間
      p_TotalExecCounts := 0;

      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
             v_ExecDateTime,p_ErrCode,p_ErrMsg);

      RETURN -2;

    END IF;

  ELSE
--**********************************************************************
--* * * * * *  歷史資料(非第一次資料)及即時查詢當日異動資料* * * * * * *
--**********************************************************************
--歷史資料(非第一次資料)及即時查詢當日異動資料
    v_FirstFlag := 0;--非執行第一次資料
    p_FirstFlag := 0;

    v_StartDateTime := SubStr(p_HandleEDateTime,1,10)||' 00:00:00';
    v_StopDateTime := p_HandleEDateTime;
    
    
    --檢核處理日期是否大於第一天
    IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') > To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN
      --處理前先把原來的資料清空
      --歷史查詢:SO455   立即查詢:SO456
      IF p_InstQuery = '0' THEN--歷史資料
        v_Table := 'SO455';

        Delete SO455 Where CompCode=p_CompCode And
               To_Char(HandleEDateTime,'YYYY/MM/DD')=SubStr(p_HandleEDateTime,1,10);
      ELSE
        v_Table := 'SO456';

        Delete SO456;

        --即時查詢必須產生唯一的一組SMS消息編碼,歷史查詢則由監管平台要資料時產生
        select S_SGSMSG_SeqNo.NextVal into v_SmsMsgSeq from dual;
        p_SrcMsgID := v_SmsMsgSeq;
      END IF;




--***************************************
--當日新訂購的頻道
      FOR crDiffDataA in ccDiffDataA loop
        --皆用新訂購的頻道SetTime
        v_RealDateTime := To_Char(crDiffDataA.SetTime,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 0;--新訂購的頻道資料皆為New(0:Modify  1:Delete)
        v_PurchaseID := crDiffDataA.OrderNo;
        v_ICCardNo := crDiffDataA.SmartCardNo;
        v_ICCardID := '';--目前沒有
        v_ProductCode := crDiffDataA.ProductCode;
        v_SetTime :=To_Char(crDiffDataA.SetTime,'YYYY/MM/DD HH24:MI:SS');
        v_DueDate :=To_Char(crDiffDataA.DueDate,'YYYY/MM/DD HH24:MI:SS');


        v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                      v_RealDateTime,v_ExecDateTime,
                      v_Status,v_PurchaseID,v_ICCardNo,
                      v_ICCardID,p_CasID,v_ProductCode,
                      v_SetTime,v_DueDate);
                      

        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--即時查詢
            SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                         v_ExecDateTime,p_InstQuery,p_SrcCode,
                         v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                         p_ErrCode,p_ErrMsg);
          END IF;
          RETURN -1;
          
        ELSE
          v_TotalExecCounts := v_TotalExecCounts + 1;
        END IF;
      END LOOP;
      
      
--***************************************
--當日取消訂購的頻道
      v_CancelChStr := ChangeCancelChStr(p_CancelChStr);
      v_SQL := 'Select A.SmartCardNo,B.ChannelID AS ProductCode,'||
                'To_Char(A.AuthorStartDate,''YYYY/MM/DD HH24:MI:SS''),'||
                'To_Char(A.AuthorStopDate,''YYYY/MM/DD HH24:MI:SS''),'||
                'A.UpdTime from SOAC0202 A,CD024 B '||
                'Where A.CompCode='||p_CompCode||
                ' And A.ChCode=B.CodeNo And B.PayFlag<>0 '||
                'And To_Date(A.UpdTime,''YYYY/MM/DD HH24:MI:SS'')>= To_Date('''||v_StartDateTime||''',''YYYY/MM/DD HH24:MI:SS'') '||
                'And To_Date(A.UpdTime,''YYYY/MM/DD HH24:MI:SS'')<= To_Date('''||v_StopDateTime||''',''YYYY/MM/DD HH24:MI:SS'') '||
                'And A.ModeType IN ('||v_CancelChStr||')';
                
                
                
      --2.開啟Dynamic sql
      BEGIN

        IF NOT cDynamic%ISOPEN then
          OPEN cDynamic FOR v_SQL;
        END IF;

      EXCEPTION
        WHEN OTHERS THEN
          ROLLBACK;
          p_ErrCode := -1;
          p_ErrMsg := 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--即時查詢
            SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                         v_ExecDateTime,p_InstQuery,p_SrcCode,
                         v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                         p_ErrCode,p_ErrMsg);
          END IF;
      END;
      
      
      --3.loop取每一筆資料來處理
      LOOP
        FETCH cDynamic INTO v_ICCardNo,v_ProductCode,v_SetTime,v_DueDate,v_RealDateTime;
        EXIT WHEN cDynamic%NOTFOUND;
        --皆用取消訂購的頻道 AuthorStopDate
        v_Status := 1;--新訂購的頻道資料皆為New(0:Modify  1:Delete)
        v_ICCardID := '';--目前沒有

        --取消訂購的頻道中的『訂單編號』，由監管GateWay自行產生，並僅存於監管GateWay的Table中
        select S_SGSORDER_SeqNo.NextVal into v_SmsOrderSeq from dual;
        v_OrderTitle := SubStr(Trim(Replace(v_SetTime,'/','')),1,8);
--        v_PurchaseID := v_OrderTitle||'S'||trim(to_char(v_SmsOrderSeq,'000000'));
        v_PurchaseID := v_OrderTitle||'0'||trim(to_char(v_SmsOrderSeq,'000000'));
        

        v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                      v_RealDateTime,v_ExecDateTime,
                      v_Status,v_PurchaseID,v_ICCardNo,
                      v_ICCardID,p_CasID,v_ProductCode,
                      v_SetTime,v_DueDate);

        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--即時查詢
            SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                         v_ExecDateTime,p_InstQuery,p_SrcCode,
                         v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                         p_ErrCode,p_ErrMsg);
          END IF;
          RETURN -1;
          
        ELSE
          v_TotalExecCounts := v_TotalExecCounts + 1;
        END IF;
        
      END LOOP;
      
      --4.Close dynamic cursor
      CLOSE cDynamic;



      IF p_InstQuery = '0' THEN--歷史資料
        --NightRun正常執行完畢Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,NULL,NULL);

      ELSE--即時查詢
        SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                     v_ExecDateTime,p_InstQuery,p_SrcCode,
                     v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                     NULL,NULL);
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_103_Content';--查詢時間必須大於第一次時間
      p_TotalExecCounts := 0;


      IF p_InstQuery = '0' THEN--歷史資料
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,p_ErrCode,p_ErrMsg);

      ELSE--即時查詢
        SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                     v_ExecDateTime,p_InstQuery,p_SrcCode,
                     v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                     p_ErrCode,p_ErrMsg);
      END IF;

      RETURN -2;
    END IF;
  END IF;


  p_ErrCode := 0;
  p_ErrMsg := ' ';
  p_TotalExecCounts := v_TotalExecCounts;
  RETURN 0;

EXCEPTION
 WHEN OTHERS THEN
  ROLLBACK;
  p_ErrCode := -1;
  p_ErrMsg := 'Other Error:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
  p_TotalExecCounts := 0;
  
  
  IF p_InstQuery = '0' THEN--歷史資料
    SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
           v_ExecDateTime,p_ErrCode,p_ErrMsg);

  ELSE--即時查詢
    SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                 v_ExecDateTime,p_InstQuery,p_SrcCode,
                 v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                 p_ErrCode,p_ErrMsg);
  END IF;
                 
  RETURN -1;
END; -- Function SF_ProdPurchaseQuery
/