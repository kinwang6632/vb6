CREATE OR REPLACE FUNCTION SF_CAPRODUCTQUERY
  (p_CompCode varchar2,p_SrcCode varchar2,
   p_DstCode varchar2,p_DstMsgID varchar2,p_CasID varchar2,
   p_InstQuery varchar2,p_FirstDate varchar2,p_HandleEDateTime varchar2,
   p_SrcMsgID OUT varchar2,p_ErrCode OUT varchar2,p_ErrMsg OUT varchar2,
   p_TotalExecCounts OUT number,
   p_TotalProductNum OUT Number,p_TotalChannelNum OUT Number,p_FirstFlag OUT varchar2)
  RETURN  number AS
/*
--@D:\App\SGS\Script\SF_CAProductQuery;
VARIABLE ReturnNo NUMBER
VARIABLE ErrCode NUMBER
VARIABLE ErrMsg VARCHAR2(1000)
VARIABLE SrcMsgID VARCHAR2(30)
EXEC :ReturnNo := SF_CAProductQuery('7','110000S01','110000G01','','0x4a01','0','2004/03/09','2004/03/09 23:59:59',:SrcMsgID,:ErrCode,:ErrMsg)
PRINT ReturnNo
PRINT ErrCode
PRINT ErrMsg
PRINT SrcMsgID


  產品設置資訊查詢應答介面用於服務平臺將CA系統中當前全部產品資訊返回給監管平臺
  檔名: SF_CAProductQuery.sql

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
    Date: 2004.03.11
    2004.03.18 修改 :查詢Log,改由前端處理完成後Log至SO460
    2004.03.19 修改 :查詢Log,改由前端處理完成後Log至SO459
    2004.04.05 修改 :SaveDataToDB有Exception時回傳相關欄位
    2004.04.16 修改 :1.去除停播資料
                     2.改開播頻道StartDate<=HandleEDateTime,而不是某日區間
*/
  v_SQL              	varchar2(4000);		-- must long enough

  v_CurrDateTime      Date;
  v_FirstDate         varchar2(20);


  v_Counts            Number := 0;
  v_TotalExecCounts   Number := 0;
  v_TotalProductNum   Number := 0;
  v_TotalChannelNum   Number := 0;


  v_FirstFlag         Number;
  v_ModeType          Number;
  v_RealDateTime      varchar2(20);
  v_ExecDateTime      varchar2(20);
  v_SmsMsgSeq         Number;
  v_ProductCode       varchar2(20);
  v_ProductName       varchar2(100);
  v_Provider          varchar2(40);
  v_ChannelName       varchar2(40);
  


  v_StartDateTime     varchar2(20);--非第一次查詢的某日開始時間
  v_StopDateTime      varchar2(20);--非第一次查詢的某日結束時間

  v_Table             varchar2(10);

  v_ErrMsg            varchar2(2000);


  cursor ccFirstData is--第一次查詢完整資料
    Select A.ChannelID as ProductCode,A.Description as ProductName,B.CodeNo,B.Description as ChannelName,B.ProviderName,
           A.StartDate,A.StopDate from CD024 A,CD024B B
    Where A.CompCode=p_CompCode And A.StopFlag=0
    And A.ChannelID=B.ChannelID;


  cursor ccDiffDataA is--開播頻道
    Select A.ChannelID as ProductCode,A.Description as ProductName,A.StartDate,A.StopDate,B.CodeNo,
           B.Description as ChannelName,B.ProviderName from CD024 A,CD024B B
    Where A.CompCode=p_CompCode And A.StopFlag=0
    And A.ChannelID=B.ChannelID
    And A.StartDate<= To_Date(v_StopDateTime,'YYYY/MM/DD HH24:MI:SS')
    And A.StartDate Is Not Null And A.StopDate Is Null;



/*
  cursor ccDiffDataA is--某日開播頻道
    Select A.ChannelID as ProductCode,A.Description as ProductName,A.StartDate,A.StopDate,B.CodeNo,
           B.Description as ChannelName,B.ProviderName from CD024 A,CD024B B
    Where A.CompCode=p_CompCode And A.StopFlag=0
    And A.ChannelID=B.ChannelID
    And To_Date(To_Char(A.StartDate,'YYYY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')>= To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    And To_Date(To_Char(A.StartDate,'YYYY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')<= To_Date(v_StopDateTime,'YYYY/MM/DD HH24:MI:SS')
    And A.StartDate Is Not Null And A.StopDate Is Null;



  cursor ccDiffDataB is--某日停播頻道
    Select A.ChannelID as ProductCode,A.Description as ProductName,A.StartDate,A.StopDate,B.CodeNo,
           B.Description as ChannelName,B.ProviderName from CD024 A,CD024B B
    Where A.CompCode=p_CompCode And A.StopFlag=0
    And A.ChannelID=B.ChannelID
    And To_Date(To_Char(A.StopDate,'YYYY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')>= To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    And To_Date(To_Char(A.StopDate,'YYYY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')<= To_Date(v_StopDateTime,'YYYY/MM/DD HH24:MI:SS')
    And A.StartDate Is Not Null And A.StopDate Is Not Null;
*/





--將資料儲存置SO453或SO454
  FUNCTION SaveDataToDB(I_Table varchar2,I_CompCode Number,I_HandleEDateTime varchar2,
                        I_RealDateTime varchar2,I_ExecDateTime varchar2,
                        I_ProductCode varchar2,I_ProductName varchar2,I_CasID varchar2,
                        I_Provider varchar2,I_ChannelName varchar2)
  RETURN varchar2
  AS
    L_SQL varchar2(4000);
    L_ErrorMsg varchar2(1000);
  BEGIN
    BEGIN


      L_SQL := 'Insert Into '||I_Table||' (CompCode,HandleEDateTime,RealDateTime,'||
               'ExecDateTime,ProductCode,ProductName,CasID,Provider,ChannelName) '||
               'Values('||I_CompCode||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_RealDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'''||
               I_ProductCode||''','''||I_ProductName||''','''||I_CasID||''','''||
               I_Provider||''','''||I_ChannelName||''')';
               
               
      L_ErrorMsg := '  ErrorMsg:Table='||I_Table||
                    '  CompCode='||I_CompCode||
                    '  HandleEDateTime='||I_HandleEDateTime||
                    '  RealDateTime='||I_RealDateTime||
                    '  ExecDateTime='||I_ExecDateTime||
                    '  ProductCode='||I_ProductCode||
                    '  ProductName='||I_ProductName||
                    '  CasID='||I_CasID||
                    '  Provider='||I_Provider||
                    '  ChannelName='||I_ChannelName;
               
               
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
                        I_TotalProductNum Number,I_TotalChannelNum Number,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --查詢Log改由前端處理完成後Log
    L_SQL := ' ';
/*
    IF I_ErrorMsg IS NULL THEN
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime,TotalProductNum,TotalChannelNum) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_TotalProductNum||','||I_TotalChannelNum||')';
    ELSE
      L_SQL := 'Insert Into SO459 (CompCode,ModeType,FirstFlag,HandleEDateTime,'||
               'ExecDateTime,TotalProductNum,TotalChannelNum,ErrorCode,ErrorMsg) '||
               'Values('||I_CompCode||','||I_ModeType||','||I_FirstFlag||
               ',TO_DATE('''||I_HandleEDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               'TO_DATE('''||I_ExecDateTime||''',''YYYY/MM/DD HH24:MI:SS''),'||
               I_TotalProductNum||','||I_TotalChannelNum||','''||
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
  p_TotalProductNum := 0;
  p_TotalChannelNum := 0;

  --ModeType=1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
  v_ModeType := 2;


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

        --當前整個SMS內CA產品總數
        Select Count(Distinct ChannelID) Into v_TotalProductNum
          from CD024 Where CompCode=p_CompCode
          And StopFlag=0 And ChannelID Is Not Null
          And StartDate<= To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS');
          --And To_Date(To_Char(A.StartDate,'YYYY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')<= To_Date(v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');


        --當前整個SMS內數位電視頻道總數
        Select Count(CodeNo) Into v_TotalChannelNum from CD024A
          Where CompCode=p_CompCode And StopFlag=0;
          
        p_TotalProductNum := v_TotalProductNum;
        p_TotalChannelNum := v_TotalChannelNum;



        FOR crFirstData in ccFirstData loop
          --第一次執行皆用開播時間
          v_RealDateTime := To_Char(crFirstData.StartDate,'YYYY/MM/DD HH24:MI:SS');
          v_ProductCode := crFirstData.ProductCode;
          v_ProductName := crFirstData.ProductName;
          v_Provider := crFirstData.ProviderName;
          v_ChannelName := crFirstData.ChannelName;


          v_ErrMsg := SaveDataToDB('SO453',p_CompCode,p_HandleEDateTime,
                      v_RealDateTime,v_ExecDateTime,v_ProductCode,
                      v_ProductName,p_CasID,v_Provider,v_ChannelName);


          IF v_ErrMsg Is Not Null THEN
            p_ErrCode := -1;
            p_ErrMsg := SubStr(v_ErrMsg,1,2000);
            p_TotalExecCounts := 0;


            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
                         p_ErrCode,p_ErrMsg);

            RETURN -1;
            
          ELSE
            v_TotalExecCounts := v_TotalExecCounts + 1;
          END IF;

        END LOOP;



      --NightRun正常執行完畢Log
      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
                     NULL,NULL);
      ELSE
        p_ErrCode := -2;
        p_ErrMsg := 'Return_Msg_104_Content';--尚未執行第一次資料
        p_TotalExecCounts := 0;

        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
               p_ErrCode,p_ErrMsg);

        RETURN -2;
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_102_Content';--未到執行第一次時間
      p_TotalExecCounts := 0;

      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
             v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
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
      --歷史查詢:SO453   立即查詢:SO454
      IF p_InstQuery = '0' THEN--歷史資料
        v_Table := 'SO453';

        Delete SO453 Where CompCode=p_CompCode And
               To_Char(HandleEDateTime,'YYYY/MM/DD')=SubStr(p_HandleEDateTime,1,10);
      ELSE
        v_Table := 'SO454';

        Delete SO454;

        --即時查詢必須產生唯一的一組SMS消息編碼,歷史查詢則由監管平台要資料時產生
        select S_SGSMSG_SeqNo.NextVal into v_SmsMsgSeq from dual;
        p_SrcMsgID := v_SmsMsgSeq;
      END IF;



      --計算當前整個SMS內CA產品總數
      Select Count(Distinct ChannelID) Into v_TotalProductNum
        from CD024 Where CompCode=p_CompCode
        And StopFlag=0 And ChannelID Is Not Null
        And StartDate<= To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS');

      --計算當前整個SMS內數位電視頻道總數
      Select Count(CodeNo) Into v_TotalChannelNum from CD024A
        Where CompCode=p_CompCode And StopFlag=0;
        
        
      p_TotalProductNum := v_TotalProductNum;
      p_TotalChannelNum := v_TotalChannelNum;



--***************************************
--當日開播資料
      FOR crDiffDataA in ccDiffDataA loop
        --皆用開播時間
        v_RealDateTime := To_Char(crDiffDataA.StartDate,'YYYY/MM/DD HH24:MI:SS');
        v_ProductCode := crDiffDataA.ProductCode;
        v_ProductName := crDiffDataA.ProductName;
        v_Provider := crDiffDataA.ProviderName;
        v_ChannelName := crDiffDataA.ChannelName;


        v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                    v_RealDateTime,v_ExecDateTime,v_ProductCode,
                    v_ProductName,p_CasID,v_Provider,v_ChannelName);




        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
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
/*
--當日刪減資料
      FOR crDiffDataB in ccDiffDataB loop
          --皆用停播時間
          v_RealDateTime := To_Char(crDiffDataB.StopDate,'YYYY/MM/DD HH24:MI:SS');
          v_ProductCode := crDiffDataB.ProductCode;
          v_ProductName := crDiffDataB.ProductName;
          v_Provider := crDiffDataB.ProviderName;
          v_ChannelName := crDiffDataB.ChannelName;


          v_ErrMsg := SaveDataToDB(v_Table,p_CompCode,p_HandleEDateTime,
                      v_RealDateTime,v_ExecDateTime,v_ProductCode,
                      v_ProductName,p_CasID,v_Provider,v_ChannelName);




        IF v_ErrMsg Is Not Null THEN
          p_ErrCode := -1;
          p_ErrMsg := SubStr(v_ErrMsg,1,2000);
          p_TotalExecCounts := 0;

          IF p_InstQuery = '0' THEN--歷史資料
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
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
      
*/


      --NightRun正常執行完畢Log
      IF p_InstQuery = '0' THEN--歷史資料
        --NightRun正常執行完畢Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
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
               v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
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
           v_ExecDateTime,v_TotalProductNum,v_TotalChannelNum,
           p_ErrCode,p_ErrMsg);

  ELSE--即時查詢
    SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                 v_ExecDateTime,p_InstQuery,p_SrcCode,
                 v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                 p_ErrCode,p_ErrMsg);
  END IF;
  
  
  RETURN -1;
END; -- Function SF_CAProductQuery
/