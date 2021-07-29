CREATE OR REPLACE FUNCTION SF_ICCARDQUERY
  (p_CompCode varchar2,p_SrcCode varchar2,
   p_DstCode varchar2,p_DstMsgID varchar2,p_CasID varchar2,
   p_InstQuery varchar2,p_FirstDate varchar2,p_HandleEDateTime varchar2,
   p_SrcMsgID OUT varchar2,p_ErrCode OUT varchar2,p_ErrMsg OUT varchar2,
   p_TotalExecCounts OUT Number,
   p_TotalICCardNum OUT Number,p_TotalSubNum OUT Number,p_FirstFlag OUT varchar2)
  RETURN  number AS
/*
--@D:\App\SGS\Script\SF_ICCardQuery;
VARIABLE ReturnNo NUMBER
VARIABLE ErrCode NUMBER
VARIABLE ErrMsg VARCHAR2(1000)
VARIABLE SrcMsgID VARCHAR2(30)
EXEC :ReturnNo := SF_ICCardQuery('7','110000S01','110000G01','','0x4a01','0','2004/03/09','2004/03/09 23:59:59',:SrcMsgID,:ErrCode,:ErrMsg)
PRINT ReturnNo
PRINT ErrCode
PRINT ErrMsg
PRINT SrcMsgID


  IC卡資訊查詢應答介面用於服務平臺將IC卡基本資訊返回給監管平臺。
  檔名: SF_ICCardQuery.sql

  說明:


  IN   p_CompCode         :公司別
       p_SrcCode          :SMS編碼
       p_DstCode          :監管中心/分中心編碼
       p_DstMsgID         :應答資料對應的監管請求的MSGID
       p_CasID            :CAS的ID
       p_InstQuery        :是否即時查詢的返回結果(0：非即時查詢返回結果  1：即時查詢結果)
       p_FirstDate        :第一次產生完整資料的日期 YYYY/MM/DD
       p_HandleEDateTime  :要產生資料的結束時間 YYYY/MM/DD HH24:MI:SS
       
       

  OUT  p_ErrCode         number:   結果訊息
       p_ErrMsg          varchar2: 結果訊息

       Return            number:   處理結果代碼, 對應訊息存於p_RetMsg中
                             =0: p_RetMsg='',傳回空值,代表正常執行完畢
                             -1: p_RetMsg=代表下SQL指令時有Exception
                             -2: p_RetMsg=代表
                            -99: p_RetMsg='其他錯誤'

    By: Jackal
    Date: 2004.03.08
    2004.03.18 修改 :查詢Log,改由前端處理完成後Log至SO460
    2004.03.19 修改 :查詢Log,改由前端處理完成後Log至SO459
    2004.04.05 修改 :SaveDataToDB有Exception時回傳相關欄位
    2004.04.16 修改 :計算SMS內IC卡總數及SMS內用戶總數多加
                     InstDate日期條件<=HandleEDateTime
                     PRDate日期條件>=HandleEDateTime
*/
  v_SQL              	varchar2(4000);		-- must long enough
  
  v_CurrDateTime      Date;
  v_FirstDate         varchar2(20);
  

  v_Counts            Number := 0;
  v_TotalExecCounts   Number := 0;
  v_TotalICCardNum    Number := 0;
  v_TotalSubNum       Number := 0;
  
  v_ICCardNo          varchar2(20);
  v_ICCardID          varchar2(20);
  v_ICVersion         varchar2(10);
  v_Feature           varchar2(300);
  v_CustName          varchar2(50);
  v_ID                varchar2(15);
  v_SubVocation       varchar2(50);
  v_Status            Number;
  v_SubAge            varchar2(4);
  v_FirstFlag         Number;
  v_ModeType          Number;
  v_RealDateTime      varchar2(20);
  v_ExecDateTime      varchar2(20);
  v_SmsMsgSeq         Number;
  
  
  v_StartDateTime     varchar2(20);--非第一次查詢的某日開始時間
  v_StopDateTime      varchar2(20);--非第一次查詢的某日結束時間
  
  v_Table             varchar2(10);
  
  v_ErrMsg            varchar2(2000);
  
  
  
  
  
  cursor ccFirstData is--第一次查詢完整資料
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate,B.CustName,
      B.ID,B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Null And A.InstDate Is Not Null
    And A.SmartCardNo Is Not Null;


  cursor ccDiffDataA is--某日IC卡裝機資料
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate ,B.CustName,
      B.ID, B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Null And A.InstDate Is Not Null And A.SmartCardNo Is Not Null
    And A.InstDate Between To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    And To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');
    
    
  cursor ccDiffDataB is--某日IC卡拆機資料
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate ,B.CustName,
      B.ID,B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Not Null And A.InstDate Is Not Null
    And A.SmartCardNo Is Not Null
    And A.PRDate Between To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    AND To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');



--將資料儲存置SO451或SO452
  FUNCTION SaveDataToDB(I_Table varchar2,I_CompCode Number,I_HandleEDateTime varchar2,
                        I_RealDateTime varchar2,I_ExecDateTime varchar2,
                        I_Status Number,I_ICCardNo varchar2,I_ICCardID varchar2,
                        I_CasID varchar2,I_Feature varchar2,I_Version varchar2,
                        I_SubName varchar2,I_SubVocation varchar2,
                        I_SubAge varchar2,I_SubIDCardNo varchar2)
  RETURN varchar2
  AS
    L_SQL varchar2(4000);
    L_ErrorMsg varchar2(1000);
  BEGIN
    BEGIN
      

      L_SQL := 'Insert Into '||I_Table||' (CompCode,HandleEDateTime,RealDateTime,'||
               'ExecDateTime,Status,ICCardNo,ICCardID,CasID,Feature,'||
               'Version,SubName,SubVocation,SubAge,SubIDCardNo) '||
               'Values('||I_CompCode||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_RealDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_Status||','''||I_ICCardNo||''','''||I_ICCardID||
               ''','''||I_CasID||''','''||I_Feature||''','''||I_Version||
               ''','''||I_SubName||''','''||I_SubVocation||''','||I_SubAge||
               ','''||I_SubIDCardNo||''')';
               
               
      L_ErrorMsg := '  ErrorMsg:Table='||I_Table||
                    '  CompCode='||I_CompCode||
                    '  HandleEDateTime='||I_HandleEDateTime||
                    '  RealDateTime='||I_RealDateTime||
                    '  ExecDateTime='||I_ExecDateTime||
                    '  Status='||I_Status||
                    '  ICCardNo='||I_ICCardNo||
                    '  ICCardID='||I_ICCardID||
                    '  CasID='||I_CasID||
                    '  Feature='||I_Feature||
                    '  Version='||I_Version||
                    '  SubName='||I_SubName||
                    '  SubVocation='||I_SubVocation||
                    '  SubAge='||I_SubAge||
                    '  SubIDCardNo='||I_SubIDCardNo;
                    
                    
      EXECUTE IMMEDIATE L_SQL;
      
      
      RETURN Null;
    EXCEPTION
      WHEN others THEN
        ROLLBACK;--之前做的先回復

        --RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||'    '||L_SQL;
        RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||L_ErrorMsg;
    END;
  END;
  
  
  
  
  

--將資料LOG至SO459
  PROCEDURE SaveSO459Log(I_CompCode Number,I_ModeType Number,I_FirstFlag Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_TotalICCardNum Number,I_TotalSubNum Number,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --查詢Log改由前端處理完成後Log
    L_SQL := ' ';
/*
    IF I_ErrorMsg IS NULL THEN
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime,TotalICCardNum,TotalSubNum) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_TotalICCardNum||','||I_TotalSubNum||')';
    ELSE
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime,TotalICCardNum,TotalSubNum,ErrorCode,ErrorMsg) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_TotalICCardNum||','||I_TotalSubNum||','''||
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
  

BEGIN

  --取出系統時間
  --SELECT To_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') INTO v_CurrDateTime FROM dual;
  v_CurrDateTime := sysdate;
  v_ExecDateTime := To_Char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');
  v_FirstDate := p_FirstDate;
  p_SrcMsgID := ' ';
  p_FirstFlag := ' ';
  p_TotalICCardNum := 0;
  p_TotalSubNum := 0;
  
  --ModeType=1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
  v_ModeType := 1;


  --判別第一次完整資料(ICCardQuery)是否已經做過
  SELECT COUNT(*) INTO v_Counts FROM SO459
    WHERE CompCode=p_CompCode And FirstFlag='1' And ErrorCode IS NULL And ModeType=v_ModeType;
    --AND To_Char(HandleEDateTime,'YYYY/MM/DD')=p_FirstDate;
    
  
  
  IF v_Counts=0 THEN
--**********************************************************************
--* * * * * * 歷史資料(第一次資料)查出未拆機的所有IC卡資料 * * * * * * *
--**********************************************************************
    v_FirstFlag := 1;--執行第一次資料
    p_FirstFlag := 1;
    
    IF To_Date(To_Char(v_CurrDateTime,'YYYY/MM/DD'),'YYYY/MM/DD') > To_Date(v_FirstDate,'YYYY/MM/DD') THEN
      --檢核處理日期是否大於第一天
      IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') = To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN

        --計算當前整個SMS內IC卡總數
        Select Count(CustID) Into v_TotalICCardNum from SO004
          Where CompCode=p_CompCode
          And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
          And InstDate Is Not Null
          And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
          And SmartCardNo Is Not Null;
          

        --計算當前整個SMS內用戶總數
        Select Count(Distinct CustID) Into v_TotalSubNum from SO004
          Where CompCode=p_CompCode
          And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
          And InstDate Is Not Null
          And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
          And SmartCardNo Is Not Null;
          
          
          
        p_TotalICCardNum := v_TotalICCardNum;
        p_TotalSubNum := v_TotalSubNum;
          



        FOR crFirstData in ccFirstData loop

          v_RealDateTime := To_Char(crFirstData.InstDate,'YYYY/MM/DD HH24:MI:SS');
          v_Status := 0;--第一次歷史資料皆為New(0:New  1:Delete)
          v_ICCardNo := crFirstData.SmartCardNo;
          v_ICCardID := '';--目前沒有
          v_Feature := '';--目前沒有
          v_ICVersion := '';--目前沒有
          v_CustName := crFirstData.CustName;
          v_SubVocation := '';--目前沒有
          v_SubAge := 'null';
          v_ID := crFirstData.ID;



          v_ErrMsg := SaveDataToDB('SO451',p_CompCode,p_HandleEDateTime,
                      v_RealDateTime,v_ExecDateTime,v_Status,v_ICCardNo,
                      v_ICCardID,p_CasID,v_Feature,v_ICVersion,v_CustName,
                      v_SubVocation,v_SubAge,v_ID);


          IF v_ErrMsg Is Not Null THEN
            p_ErrCode := -1;
            p_ErrMsg := SubStr(v_ErrMsg,1,2000);
            p_TotalExecCounts := 0;


            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                         p_ErrCode,p_ErrMsg);

            RETURN -1;
            
          ELSE
            v_TotalExecCounts := v_TotalExecCounts + 1;
          END IF;
          
          
          
        END LOOP;


        --NightRun正常執行完畢Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                     NULL,NULL);
      ELSE
        p_ErrCode := -2;
        p_ErrMsg := 'Return_Msg_104_Content';--尚未執行第一次資料
        p_TotalExecCounts := 0;

        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
               p_ErrCode,p_ErrMsg);

        RETURN -2;
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_102_Content';--未到執行第一次時間
      p_TotalExecCounts := 0;
      
      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
             v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
             p_ErrCode,p_ErrMsg);

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
      --歷史查詢:SO451   立即查詢:SO452
      IF p_InstQuery = '0' THEN--歷史資料
        v_Table := 'SO451';
        
        Delete SO451 Where CompCode=p_CompCode And
               To_Char(HandleEDateTime,'YYYY/MM/DD')=SubStr(p_HandleEDateTime,1,10);
      ELSE
        v_Table := 'SO452';
        
        Delete SO452;
        
        --即時查詢必須產生唯一的一組SMS消息編碼,歷史查詢則由監管平台要資料時產生
        select S_SGSMSG_SeqNo.NextVal into v_SmsMsgSeq from dual;
        p_SrcMsgID := v_SmsMsgSeq;
      END IF;
      
      
      
    
      --計算當前整個SMS內IC卡總數
      Select Count(CustID) Into v_TotalICCardNum from SO004
        Where CompCode=p_CompCode
        And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
        And InstDate Is Not Null
        And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
        And SmartCardNo Is Not Null;


      --計算當前整個SMS內用戶總數
      Select Count(Distinct CustID) Into v_TotalSubNum from SO004
        Where CompCode=p_CompCode
        And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
        And InstDate Is Not Null
        And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
        And SmartCardNo Is Not Null;

        
        
      p_TotalICCardNum := v_TotalICCardNum;
      p_TotalSubNum := v_TotalSubNum;


      
--***************************************
--當日新增資料
      FOR crDiffDataA in ccDiffDataA loop
        v_RealDateTime := To_Char(crDiffDataA.InstDate,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 0;--非第一次歷史資料皆為New(0:New  1:Delete)
        v_ICCardNo := crDiffDataA.SmartCardNo;
        v_ICCardID := '';--目前沒有
        v_Feature := '';--目前沒有
        v_ICVersion := '';--目前沒有
        v_CustName := crDiffDataA.CustName;
        v_SubVocation := '';--目前沒有
        v_SubAge := 'null';
        v_ID := crDiffDataA.ID;
        
        
        v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                    v_RealDateTime,v_ExecDateTime,v_Status,v_ICCardNo,
                    v_ICCardID,p_CasID,v_Feature,v_ICVersion,v_CustName,
                    v_SubVocation,v_SubAge,v_ID);
                    
                    
        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                         p_ErrCode,p_ErrMsg);
                         
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
--當日刪減資料
      FOR crDiffDataB in ccDiffDataB loop
        v_RealDateTime := To_Char(crDiffDataB.PRDate,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 1;--非第一次歷史資料皆為New(0:New  1:Delete)
        v_ICCardNo := crDiffDataB.SmartCardNo;
        v_ICCardID := '';--目前沒有
        v_Feature := '';--目前沒有
        v_ICVersion := '';--目前沒有
        v_CustName := crDiffDataB.CustName;
        v_SubVocation := '';--目前沒有
        v_SubAge := 'null';
        v_ID := crDiffDataB.ID;


        v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                    v_RealDateTime,v_ExecDateTime,v_Status,v_ICCardNo,
                    v_ICCardID,p_CasID,v_Feature,v_ICVersion,v_CustName,
                    v_SubVocation,v_SubAge,v_ID);


        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                         p_ErrCode,p_ErrMsg);

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


      --NightRun正常執行完畢Log
      IF p_InstQuery = '0' THEN--歷史資料
        --NightRun正常執行完畢Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                     NULL,NULL);

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
               v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
               p_ErrCode,p_ErrMsg);

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
           v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
           p_ErrCode,p_ErrMsg);

  ELSE--即時查詢
    SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                 v_ExecDateTime,p_InstQuery,p_SrcCode,
                 v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                 p_ErrCode,p_ErrMsg);
  END IF;
  
  RETURN -1;
END; -- Function SF_ICCardQuery
/