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


  IC�d��T�d�����������Ω�A�ȥ��O�NIC�d�򥻸�T��^���ʺޥ��O�C
  �ɦW: SF_ICCardQuery.sql

  ����:


  IN   p_CompCode         :���q�O
       p_SrcCode          :SMS�s�X
       p_DstCode          :�ʺޤ���/�����߽s�X
       p_DstMsgID         :������ƹ������ʺ޽ШD��MSGID
       p_CasID            :CAS��ID
       p_InstQuery        :�O�_�Y�ɬd�ߪ���^���G(0�G�D�Y�ɬd�ߪ�^���G  1�G�Y�ɬd�ߵ��G)
       p_FirstDate        :�Ĥ@�����ͧ����ƪ���� YYYY/MM/DD
       p_HandleEDateTime  :�n���͸�ƪ������ɶ� YYYY/MM/DD HH24:MI:SS
       
       

  OUT  p_ErrCode         number:   ���G�T��
       p_ErrMsg          varchar2: ���G�T��

       Return            number:   �B�z���G�N�X, �����T���s��p_RetMsg��
                             =0: p_RetMsg='',�Ǧ^�ŭ�,�N���`���槹��
                             -1: p_RetMsg=�N��USQL���O�ɦ�Exception
                             -2: p_RetMsg=�N��
                            -99: p_RetMsg='��L���~'

    By: Jackal
    Date: 2004.03.08
    2004.03.18 �ק� :�d��Log,��ѫe�ݳB�z������Log��SO460
    2004.03.19 �ק� :�d��Log,��ѫe�ݳB�z������Log��SO459
    2004.04.05 �ק� :SaveDataToDB��Exception�ɦ^�Ǭ������
    2004.04.16 �ק� :�p��SMS��IC�d�`�Ƥ�SMS���Τ��`�Ʀh�[
                     InstDate�������<=HandleEDateTime
                     PRDate�������>=HandleEDateTime
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
  
  
  v_StartDateTime     varchar2(20);--�D�Ĥ@���d�ߪ��Y��}�l�ɶ�
  v_StopDateTime      varchar2(20);--�D�Ĥ@���d�ߪ��Y�鵲���ɶ�
  
  v_Table             varchar2(10);
  
  v_ErrMsg            varchar2(2000);
  
  
  
  
  
  cursor ccFirstData is--�Ĥ@���d�ߧ�����
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate,B.CustName,
      B.ID,B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Null And A.InstDate Is Not Null
    And A.SmartCardNo Is Not Null;


  cursor ccDiffDataA is--�Y��IC�d�˾����
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate ,B.CustName,
      B.ID, B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Null And A.InstDate Is Not Null And A.SmartCardNo Is Not Null
    And A.InstDate Between To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    And To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');
    
    
  cursor ccDiffDataB is--�Y��IC�d������
    Select A.CustId,A.SmartCardNo,A.InstDate,A.PRDate ,B.CustName,
      B.ID,B.Brithday from SO004 A,SO001 B
    Where A.CompCode=p_CompCode And A.CustId=B.CustID
    And A.PRDate Is Not Null And A.InstDate Is Not Null
    And A.SmartCardNo Is Not Null
    And A.PRDate Between To_Date(v_StartDateTime,'YYYY/MM/DD HH24:MI:SS')
    AND To_Date (v_StopDateTime,'YYYY/MM/DD HH24:MI:SS');



--�N����x�s�mSO451��SO452
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
        ROLLBACK;--���e�������^�_

        --RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||'    '||L_SQL;
        RETURN 'SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM||L_ErrorMsg;
    END;
  END;
  
  
  
  
  

--�N���LOG��SO459
  PROCEDURE SaveSO459Log(I_CompCode Number,I_ModeType Number,I_FirstFlag Number,
                        I_HandleEDateTime varchar2,I_ExecDateTime varchar2,
                        I_TotalICCardNum Number,I_TotalSubNum Number,
                        I_ErrorCode varchar2,I_ErrorMsg varchar2)
  AS
    L_SQL varchar2(4000);
  BEGIN
    --�d��Log��ѫe�ݳB�z������Log
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
  

BEGIN

  --���X�t�ήɶ�
  --SELECT To_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') INTO v_CurrDateTime FROM dual;
  v_CurrDateTime := sysdate;
  v_ExecDateTime := To_Char(v_CurrDateTime,'YYYY/MM/DD HH24:MI:SS');
  v_FirstDate := p_FirstDate;
  p_SrcMsgID := ' ';
  p_FirstFlag := ' ';
  p_TotalICCardNum := 0;
  p_TotalSubNum := 0;
  
  --ModeType=1:IC�d��T2:���~�w�q��T3:���~�w�ʸ�T4:���v��T
  v_ModeType := 1;


  --�P�O�Ĥ@��������(ICCardQuery)�O�_�w�g���L
  SELECT COUNT(*) INTO v_Counts FROM SO459
    WHERE CompCode=p_CompCode And FirstFlag='1' And ErrorCode IS NULL And ModeType=v_ModeType;
    --AND To_Char(HandleEDateTime,'YYYY/MM/DD')=p_FirstDate;
    
  
  
  IF v_Counts=0 THEN
--**********************************************************************
--* * * * * * ���v���(�Ĥ@�����)�d�X��������Ҧ�IC�d��� * * * * * * *
--**********************************************************************
    v_FirstFlag := 1;--����Ĥ@�����
    p_FirstFlag := 1;
    
    IF To_Date(To_Char(v_CurrDateTime,'YYYY/MM/DD'),'YYYY/MM/DD') > To_Date(v_FirstDate,'YYYY/MM/DD') THEN
      --�ˮֳB�z����O�_�j��Ĥ@��
      IF To_Date(p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS') = To_Date(v_FirstDate||' 23:59:59','YYYY/MM/DD HH24:MI:SS') THEN

        --�p���e���SMS��IC�d�`��
        Select Count(CustID) Into v_TotalICCardNum from SO004
          Where CompCode=p_CompCode
          And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
          And InstDate Is Not Null
          And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
          And SmartCardNo Is Not Null;
          

        --�p���e���SMS���Τ��`��
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
          v_Status := 0;--�Ĥ@�����v��ƬҬ�New(0:New  1:Delete)
          v_ICCardNo := crFirstData.SmartCardNo;
          v_ICCardID := '';--�ثe�S��
          v_Feature := '';--�ثe�S��
          v_ICVersion := '';--�ثe�S��
          v_CustName := crFirstData.CustName;
          v_SubVocation := '';--�ثe�S��
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


        --NightRun���`���槹��Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                     NULL,NULL);
      ELSE
        p_ErrCode := -2;
        p_ErrMsg := 'Return_Msg_104_Content';--�|������Ĥ@�����
        p_TotalExecCounts := 0;

        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
               v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
               p_ErrCode,p_ErrMsg);

        RETURN -2;
      END IF;
    ELSE
      p_ErrCode := -2;
      p_ErrMsg := 'Return_Msg_102_Content';--�������Ĥ@���ɶ�
      p_TotalExecCounts := 0;
      
      SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
             v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
             p_ErrCode,p_ErrMsg);

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
      --���v�d��:SO451   �ߧY�d��:SO452
      IF p_InstQuery = '0' THEN--���v���
        v_Table := 'SO451';
        
        Delete SO451 Where CompCode=p_CompCode And
               To_Char(HandleEDateTime,'YYYY/MM/DD')=SubStr(p_HandleEDateTime,1,10);
      ELSE
        v_Table := 'SO452';
        
        Delete SO452;
        
        --�Y�ɬd�ߥ������Ͱߤ@���@��SMS�����s�X,���v�d�߫h�Ѻʺޥ��x�n��Ʈɲ���
        select S_SGSMSG_SeqNo.NextVal into v_SmsMsgSeq from dual;
        p_SrcMsgID := v_SmsMsgSeq;
      END IF;
      
      
      
    
      --�p���e���SMS��IC�d�`��
      Select Count(CustID) Into v_TotalICCardNum from SO004
        Where CompCode=p_CompCode
        And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
        And InstDate Is Not Null
        And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
        And SmartCardNo Is Not Null;


      --�p���e���SMS���Τ��`��
      Select Count(Distinct CustID) Into v_TotalSubNum from SO004
        Where CompCode=p_CompCode
        And (PRDate Is Null or PRDate>=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS'))
        And InstDate Is Not Null
        And InstDate<=To_Date (p_HandleEDateTime,'YYYY/MM/DD HH24:MI:SS')
        And SmartCardNo Is Not Null;

        
        
      p_TotalICCardNum := v_TotalICCardNum;
      p_TotalSubNum := v_TotalSubNum;


      
--***************************************
--���s�W���
      FOR crDiffDataA in ccDiffDataA loop
        v_RealDateTime := To_Char(crDiffDataA.InstDate,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 0;--�D�Ĥ@�����v��ƬҬ�New(0:New  1:Delete)
        v_ICCardNo := crDiffDataA.SmartCardNo;
        v_ICCardID := '';--�ثe�S��
        v_Feature := '';--�ثe�S��
        v_ICVersion := '';--�ثe�S��
        v_CustName := crDiffDataA.CustName;
        v_SubVocation := '';--�ثe�S��
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

          IF p_InstQuery = '0' THEN--���v���
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                         p_ErrCode,p_ErrMsg);
                         
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
--���R����
      FOR crDiffDataB in ccDiffDataB loop
        v_RealDateTime := To_Char(crDiffDataB.PRDate,'YYYY/MM/DD HH24:MI:SS');
        v_Status := 1;--�D�Ĥ@�����v��ƬҬ�New(0:New  1:Delete)
        v_ICCardNo := crDiffDataB.SmartCardNo;
        v_ICCardID := '';--�ثe�S��
        v_Feature := '';--�ثe�S��
        v_ICVersion := '';--�ثe�S��
        v_CustName := crDiffDataB.CustName;
        v_SubVocation := '';--�ثe�S��
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

          IF p_InstQuery = '0' THEN--���v���
            SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                         v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                         p_ErrCode,p_ErrMsg);

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


      --NightRun���`���槹��Log
      IF p_InstQuery = '0' THEN--���v���
        --NightRun���`���槹��Log
        SaveSO459Log(p_CompCode,v_ModeType,v_FirstFlag,p_HandleEDateTime,
                     v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
                     NULL,NULL);

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
               v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
               p_ErrCode,p_ErrMsg);

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
           v_ExecDateTime,v_TotalICCardNum,v_TotalSubNum,
           p_ErrCode,p_ErrMsg);

  ELSE--�Y�ɬd��
    SaveSO460Log(p_CompCode,v_ModeType,p_HandleEDateTime,
                 v_ExecDateTime,p_InstQuery,p_SrcCode,
                 v_SmsMsgSeq,p_DstCode,p_DstMsgID,
                 p_ErrCode,p_ErrMsg);
  END IF;
  
  RETURN -1;
END; -- Function SF_ICCardQuery
/