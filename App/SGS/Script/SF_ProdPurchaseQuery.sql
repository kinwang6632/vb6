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


  ���~�w�ʸ�T�d�����������Ω�A�ȥ��O�N�Ʀ�q���Τ��ʶR�Ʀ�q��?�~���ܤƸ�T��^���ʺޥ��O
  �ɦW: SF_ProdPurchaseQuery.sql

  ����:


  IN   p_CompCode         :���q�O
       p_SrcCode          :SMS�s�X
       p_DstCode          :�ʺޤ���/�����߽s�X
       p_DstMsgID         :������ƹ������ʺ޽ШD��MSGID
       p_CasID            :CAS��ID
       p_InstQuery        :�O�_�Y�ɬd�ߪ���^���G(0�G�D�Y�ɬd�ߪ�^���G  1�G�Y�ɬd�ߵ��G)
       p_FirstDate        :�Ĥ@�����ͧ����ƪ���� YYYY/MM/DD
       p_HandleEDateTime  :�n���͸�ƪ������ɶ� YYYY/MM/DD HH24:MI:SS
       p_CancelChStr      :�n�������W�D���O�N�X(�ثe�q���W�D�d�߬O���]�t�K�O�W�D�θլ��W�D)



  OUT  p_ErrCode         number:   ���G�T��
       p_ErrMsg          varchar2: ���G�T��

       Return            number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                             =0: p_RetMsg='',�Ǧ^�ŭ�,�N���`���槹��
                             -1: p_RetMsg=�N��USQL���O�ɦ�Exception
                             -2: p_RetMsg=�N��
                            -99: p_RetMsg='��L���~'

    By: Jackal
    Date: 2004.03.12
    2004.03.18 �ק� :�d��Log,��ѫe�ݳB�z������Log��SO460
    2004.03.19 �ק� :�d��Log,��ѫe�ݳB�z������Log��SO459
    2004.04.05 �ק� :SaveDataToDB��Exception�ɦ^�Ǭ������
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



  v_StartDateTime     varchar2(20);--�D�Ĥ@���d�ߪ��Y��}�l�ɶ�
  v_StopDateTime      varchar2(20);--�D�Ĥ@���d�ߪ��Y�鵲���ɶ�

  v_Table             varchar2(10);

  v_ErrMsg            varchar2(2000);
  
  type CancelCh is ref cursor;	--�ۭqcursor���A
  cDynamic CancelCh;          	--��dynamic SQL



  cursor ccFirstData is--�Ĥ@���d�ߧ�����
    Select A.OrderNo,A.SmartCardNo,A.ChCode,C.ChannelID AS ProductCode,A.SetTime,
           A.DueDate,B.AcceptTime from SO005 A,SO105 B,CD024 C
    Where A.CompCode=p_CompCode And A.ChCode=C.CodeNo
    And A.OrderNo=B.OrderNo And C.PayFlag<>0;



  cursor ccDiffDataA is--�Y��s�q�ʪ��W�D
    Select A.OrderNo,A.SmartCardNo,A.ChCode,C.ChannelID AS ProductCode,A.SetTime,
           A.DueDate,B.AcceptTime from SO005 A,SO105 B,CD024 C
    Where A.CompCode=p_CompCode And A.ChCode=C.CodeNo
    And A.OrderNo=B.OrderNo And C.PayFlag<>0 And B.AcceptTime Between
    To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')AND
    To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');



--�N����x�s�mSO455��SO456
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
        ROLLBACK;--���e�������^�_

        --RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
        RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||L_ErrorMsg;
    END;
  END;




--�N���LOG��SO459
  PROCEDURE SaveSO459Log(I_CompCode Number,I_ModeType Number,I_FirstFlag Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --�d��Log��ѫe�ݳB�z������Log
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





--�N���LOG��SO460
  PROCEDURE SaveSO460Log(I_CompCode Number,I_ModeType Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_InstQuery Number,I_SrcCode varchar2,I_SrcMsgID varchar2,
                        I_DstCode varchar2,I_DstMsgID varchar2,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --�d��Log��ѫe�ݳB�z������Log
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
  
  
--���CancelChannel
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
  --���X�t�ήɶ�
  --SELECT To_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') INTO v_CurrDateTime FROM dual;
  v_CurrDateTime := sysdate;
  v_ExecDateTime := To_Char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');
  v_FirstDate := p_FirstDate;
  p_SrcMsgID := ' ';
  p_FirstFlag := ' ';

  --ModeType=1:IC�d��T2:���~�w�q��T3:���~�w�ʸ�T4:���v��T
  v_ModeType := 3;


  --�P�O�Ĥ@��������(CAProductQuery)�O�_�w�g���L
  SELECT COUNT(*) INTO v_Counts FROM SO459
    WHERE CompCode=p_CompCode And FirstFlag='1' And ErrorCode IS NULL And ModeType=v_ModeType;



  IF v_Counts=0 THEN
--**********************************************************************
--* * * * * * ���v���(�Ĥ@�����)�d�X��������Ҧ�IC�d��� * * * * * * *
--**********************************************************************
    v_FirstFlag := 1;--����Ĥ@�����
    p_FirstFlag := 1;


    IF To_Date(To_Char(v_CurrDateTime,'YYYY/MM/DD'),'YYYY/MM/DD') > To_Date(v_FirstDate,'YYYY/MM/DD') THEN
      --�ˮֳB�z����O�_�j��Ĥ@��
      IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') = To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN

        FOR crFirstData in ccFirstData loop
          --�Ĥ@������ҥηs�q�ʪ��W�D
          v_RealDateTime := To_Char(crFirstData.SetTime,'YYYY/MM/DD HH24:MI:SS');
          v_Status := 0;--�Ĥ@�����v��ƬҬ�New(0:Modify  1:Delete)
          v_PurchaseID := crFirstData.OrderNo;
          v_ICCardNo := crFirstData.SmartCardNo;
          v_ICCardID := '';--�ثe�S��
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



      --NightRun���`���槹��Log
      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,NULL,NULL);
      ELSE
        p_ErrCode := -2;
        p_ErrMsg := 'Return_Msg_104_Content';--�|������Ĥ@�����
        p_TotalExecCounts := 0;

        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,p_ErrCode,p_ErrMsg);

        RETURN -2;
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_102_Content';--�������Ĥ@���ɶ�
      p_TotalExecCounts := 0;

      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
             v_ExecDateTime,p_ErrCode,p_ErrMsg);

      RETURN -2;

    END IF;

  ELSE
--**********************************************************************
--* * * * * *  ���v���(�D�Ĥ@�����)�ΧY�ɬd�߷�鲧�ʸ��* * * * * * *
--**********************************************************************
--���v���(�D�Ĥ@�����)�ΧY�ɬd�߷�鲧�ʸ��
    v_FirstFlag := 0;--�D����Ĥ@�����
    p_FirstFlag := 0;

    v_StartDateTime := SubStr(p_HandleEDateTime,1,10)||' 00:00:00';
    v_StopDateTime := p_HandleEDateTime;
    
    
    --�ˮֳB�z����O�_�j��Ĥ@��
    IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') > To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN
      --�B�z�e�����Ӫ���ƲM��
      --���v�d��:SO455   �ߧY�d��:SO456
      IF p_InstQuery = '0' THEN--���v���
        v_Table := 'SO455';

        Delete SO455 Where CompCode=p_CompCode And
               To_Char(HandleEDateTime,'YYYY/MM/DD')=SubStr(p_HandleEDateTime,1,10);
      ELSE
        v_Table := 'SO456';

        Delete SO456;

        --�Y�ɬd�ߥ������Ͱߤ@���@��SMS�����s�X,���v�d�߫h�Ѻʺޥ��x�n��Ʈɲ���
        select S_SGSMSG_SeqNo.NextVal into v_SmsMsgSeq from dual;
        p_SrcMsgID := v_SmsMsgSeq;
      END IF;




--***************************************
--���s�q�ʪ��W�D
      FOR crDiffDataA in ccDiffDataA loop
        --�ҥηs�q�ʪ��W�DSetTime
        v_RealDateTime := To_Char(crDiffDataA.SetTime,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 0;--�s�q�ʪ��W�D��ƬҬ�New(0:Modify  1:Delete)
        v_PurchaseID := crDiffDataA.OrderNo;
        v_ICCardNo := crDiffDataA.SmartCardNo;
        v_ICCardID := '';--�ثe�S��
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

          IF p_InstQuery = '0' THEN--���v���
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--�Y�ɬd��
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
--�������q�ʪ��W�D
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
                
                
                
      --2.�}��Dynamic sql
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

          IF p_InstQuery = '0' THEN--���v���
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--�Y�ɬd��
            SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                         v_ExecDateTime,p_InstQuery,p_SrcCode,
                         v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                         p_ErrCode,p_ErrMsg);
          END IF;
      END;
      
      
      --3.loop���C�@����ƨӳB�z
      LOOP
        FETCH cDynamic INTO v_ICCardNo,v_ProductCode,v_SetTime,v_DueDate,v_RealDateTime;
        EXIT WHEN cDynamic%NOTFOUND;
        --�ҥΨ����q�ʪ��W�D AuthorStopDate
        v_Status := 1;--�s�q�ʪ��W�D��ƬҬ�New(0:Modify  1:Delete)
        v_ICCardID := '';--�ثe�S��

        --�����q�ʪ��W�D�����y�q��s���z�A�Ѻʺ�GateWay�ۦ沣�͡A�öȦs��ʺ�GateWay��Table��
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

          IF p_InstQuery = '0' THEN--���v���
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,p_ErrCode,p_ErrMsg);

          ELSE--�Y�ɬd��
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



      IF p_InstQuery = '0' THEN--���v���
        --NightRun���`���槹��Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,NULL,NULL);

      ELSE--�Y�ɬd��
        SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                     v_ExecDateTime,p_InstQuery,p_SrcCode,
                     v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                     NULL,NULL);
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_103_Content';--�d�߮ɶ������j��Ĥ@���ɶ�
      p_TotalExecCounts := 0;


      IF p_InstQuery = '0' THEN--���v���
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,p_ErrCode,p_ErrMsg);

      ELSE--�Y�ɬd��
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
  
  
  IF p_InstQuery = '0' THEN--���v���
    SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
           v_ExecDateTime,p_ErrCode,p_ErrMsg);

  ELSE--�Y�ɬd��
    SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                 v_ExecDateTime,p_InstQuery,p_SrcCode,
                 v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                 p_ErrCode,p_ErrMsg);
  END IF;
                 
  RETURN -1;
END; -- Function SF_ProdPurchaseQuery
/