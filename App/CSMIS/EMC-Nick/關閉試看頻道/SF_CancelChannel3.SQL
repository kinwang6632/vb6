CREATE OR REPLACE FUNCTION SF_CANCELCHANNEL3 (p_CompCode number,p_Option number,p_EndData varchar2,p_ProcessingDate varchar2,p_RetCode Out Number, p_RetMsg Out Varchar2) return number AS
/*
  --@D:\SF_CancelChannel3.sql
  VARIABLE ReturnNo NUMBER
  VARIABLE p_RetCode NUMBER
  VARIABLE p_RetMsg VARCHAR2(1000)
  EXEC :ReturnNo := SF_CancelChannel3(9,0,'20031001','',:p_RetCode,:p_RetMsg);
  PRINT ReturnNo
  PRINT p_RetCode
  PRINT p_RetMsg

  目的: 將免費的付費頻道關閉
  檔名: SF_CancelChannel3.sql

  IN   p_CompCode           number: 公司別代碼
       p_Option             number: 程式流程選項
          p_Option <>1     則僅將要關閉頻道資料LOG至SO080
          p_Option = 1     將關頻道資料正式存入資料庫,
       p_EndData           頻道設定截止日SO005.SETTIME
       p_ProcessingDate    執行時間


  OUT p_RetMsg  varchar2: 處理結果自串
      p_RetCode number: 處理結果代碼, 對應訊息存於p_RetMsg中

              0: p_RetMsg='關閉頻道執行完成,共產生 ??? 筆待關頻道之資料'
             -1: p_RetMsg='INSERT 時發生錯誤
             -2: P_RetMsg='Delete時發生錯誤
             -3: p_RetMsg='Insert至SOAC0202時發生錯誤
             -4: p_RetMsg='參數錯誤:...
            -98: p_RetMsg='DELETE時發生錯誤:...
            -99: p_RetMsg='其他錯誤'

  流程:
  ..將要關閉的頻道存至陣列中(請於執行前確認是否有異動)

  ..刪除上次 Log 的SO080資料

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

                        紀錄一筆資料至SO080(客戶狀態重整log檔)
                          ..其中SO080.Type有下列三種狀態
                            Type :=  88;  --資料正常存至 Send_Nagra 的 Type
                            Type := -88;  --資料存至 Send_Nagra 出現 Exception 的 Type
                            Type := -99;  --處理其他資料出現 Exception 的 Type


  定義:
	..將客戶的業務活動為： 1-無、 201-ebox員工專案、 203-TV plus舊換新的資料篩選出來，
    將資料依下列條件分成A、B大類，
    ..A類-->維持不變，
    ..B類-->將HAPPY、HOT、PLAY BOY、YOYO-A、YOYO-B、YOYO-C、東森高爾夫、
            東森醫療健康、東森戲曲教學等共10個應付費的頻道關閉。

  PS.：A類--為依客戶類別 --> 902-公關戶(無CATV) 、 191-eBox學校免費、 192-eBox關企免費
            或依收費項目 --> 有掛短收原因 97 - 專案應收款-BOX
       B類--A類以外，全計為B類。

  Date: 2003/09/17  By:Jackal
  Modify:2003/09/24  By:Jackal  SO080.Content加串中文
  Modify:2003/09/30  By:Jackal  多傳一個參數p_EndData,用以控制資料截止日期
  Modify:2003/10/01  By:Jackal  多傳一個參廚p_ProcessingDate,用以控制執行時間
*/
  V_SQL                        Varchar2(4000);
  V_Rowcount                   Number:=0;

  V_CustID                     Number;
  V_TempCustID                 Number;
  V_SNo                        Varchar2(20);
  V_TempSNo                    Varchar2(20);
  V_Cvtsno                     Varchar2(15);
  V_SmartCardNo                Varchar2(20);
  V_Chcode                     Varchar2(3);
  V_TempChcode                 Varchar2(3);
  V_ChName                     Varchar2(20);
  V_TempChName                 Varchar2(20);
  V_ChannelID                  Varchar2(12);
  V_PromCode                   Number;
  V_TempPromCode               Number;
  V_PromName                   Varchar2(20);
  V_TempPromName               Varchar2(20);
  V_ClassCode1                 Number;
  V_TempClassCode1             Number;
  V_ClassName1                 Varchar2(20);
  V_TempClassName1             Varchar2(20);
  V_Compcode                   Number;
  V_TempCompcode               Number;
  V_ChannelIDStrList           Varchar2(1000);
  V_ChCodeStrList              Varchar2(1000);
  V_ServiceType                Varchar2(12);
  V_TempServiceType            Varchar2(12);

  V_TempChannelID              Varchar2(12);
  V_TempSmartCardNo            Varchar2(20);
  V_Content                    Varchar2(1000);
  V_Index                      Number := 0;
  V_CancelChannelIDCnt         Number := 0;
  I                            Number := 1;
  V_IsClose                    Char(1) := 'N';

  V_TIME                       Date;
  V_HIGHT_LEVEL_CMD_ID         Varchar2(5);
  V_CMD_STATUS                 Varchar2(2);
  V_OPERATOR                   Varchar2(15);

  V_count                      Number:=0;
  V_ModeType                   Varchar2(5);
  V_CancelChannelType          Number;
  V_ExceptionType1             Number;
  V_ExceptionType2             Number;
  V_Length                     Number := 0;
  V_TempProcessingDate         Date;
  V_TempEndData                Date;



Cursor n_close is
  SELECT X.CustId,X.SNo,X.COMPCODE,
         Y.Cvtsno,Y.SMARTCARDNO,Y.CHCODE,
         V.Description,V.ChannelID,
         Z.PROMCODE,Z.PROMNAME,
         W.CLASSCODE1,W.CLASSNAME1,W.ServiceType
  FROM CD024 V,SO001 W,SO004 X, SO005 Y,SO007 Z
  WHERE X.FACISNO=Y.SMARTCARDNO AND
  X.SNO IN (SELECT A.SNO FROM SO007 A, SO001 B WHERE A.CUSTID=B.CUSTID AND
            A.PROMCODE IN (1,201,203) AND A.RETURNCODE IS NULL AND
            A.INSTCODE = 201 AND B.CLASSCODE1 NOT IN (920,191,192)
    MINUS
      (SELECT DISTINCT(BILLNO) FROM SO033 WHERE STCODE=97
       UNION
       SELECT DISTINCT(BILLNO) FROM SO034 WHERE STCODE=97))
  AND X.FACICODE=202 AND X.PRDATE IS NULL AND Z.SNO=X.SNO
  AND W.CUSTID=Z.CUSTID AND V.CODENO=Y.CHCODE AND X.CUSTID=Y.CUSTID
  AND V.ChannelID IN (43,44,45,48,49,51,50,46,57,87,5,6,7,8,107)
  AND Y.SETTIME<TO_DATE(p_EndData,'YYYYMMDD')
  ORDER BY Y.SMARTCARDNO;



  TYPE Varchar2Ary IS TABLE OF varchar2(12) INDEX BY BINARY_INTEGER;
  CancelChannelIDAry Varchar2Ary;


BEGIN
  if p_CompCode Is Null then
    p_RetMsg := '參數錯誤:公司別不能為空值';
    p_RetCode := -4;
    Return -4;
  end if;


  if p_Option Is Null then
    p_RetMsg := '參數錯誤:p_Option不能為空值';
    p_RetCode := -4;
    Return -4;
  end if;


  if p_EndData Is Null then
    p_RetMsg := '參數錯誤:p_EndData不能為空值';
    p_RetCode := -4;
    Return -4;
  end if;

  V_Length := Length(p_EndData);
  IF V_Length <> 8 THEN
    p_RetMsg := '參數錯誤:p_EndData日期格式錯誤(EX:20030924)';
    p_RetCode := -4;
    Return -4;
  ELSE
    BEGIN
      V_TempEndData := TO_DATE(p_EndData,'YYYYMMDD');
    EXCEPTION
      WHEN others THEN
        p_RetMsg := '參數錯誤:p_EndData日期格式錯誤(EX:20030924)';
        p_RetCode := -4;
        Return -4;
    END;

  END IF;


  V_Length := Length(p_ProcessingDate);
  IF V_Length=14 then
    BEGIN
      V_TempProcessingDate := TO_DATE(p_ProcessingDate,'YYYYMMDDHH24MISS');
    EXCEPTION
      WHEN others THEN
        p_RetMsg := '參數錯誤:p_ProcessingDate日期格式錯誤(EX:20031001001000)';
        p_RetCode := -4;
        Return -4;
    END;
  ELSIF (V_Length=0) or (V_Length IS NULL) THEN
    V_TempProcessingDate := NULL;
  ELSE
    p_RetMsg := '參數錯誤:p_ProcessingDate日期格式錯誤(EX:20031001001000)';
    p_RetCode := -4;
    Return -4;
  END IF;
  



  --將要關閉的頻道存至陣列中
  CancelChannelIDAry(1) := '43';
  CancelChannelIDAry(2) := '44';
  CancelChannelIDAry(3) := '45';
  CancelChannelIDAry(4) := '48';
  CancelChannelIDAry(5) := '49';
  CancelChannelIDAry(6) := '51';
  CancelChannelIDAry(7) := '50';
  CancelChannelIDAry(8) := '46';
  CancelChannelIDAry(9) := '57';
  CancelChannelIDAry(10) := '87';
  CancelChannelIDAry(11) := '5';
  CancelChannelIDAry(12) := '6';
  CancelChannelIDAry(13) := '7';
  CancelChannelIDAry(14) := '8';
  CancelChannelIDAry(15) := '107';
  V_CancelChannelIDCnt := 15;


  V_HIGHT_LEVEL_CMD_ID:='B2';
  V_CMD_STATUS:='W';
  V_OPERATOR:='批次作業';
  V_ModeType:='B2';
  V_TIME:=SYSDATE;

  V_CancelChannelType := 88;--資料正常存至 Send_Nagra 的 Type
  V_ExceptionType1 := -88;  --資料存至 Send_Nagra 出現 Exception 的 Type
  V_ExceptionType2 := -99;  --處理其他資料出現 Exception 的 Type


  V_TempSmartCardNo := '';

  --刪除上次統計的SO080資料
  BEGIN
    DELETE SO080 WHERE TYPE=V_CancelChannelType OR TYPE=V_ExceptionType1 OR TYPE=V_ExceptionType2;
  EXCEPTION
    WHEN others THEN
      p_RetMsg := 'DELETE時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
      p_RetCode := -98;
      RETURN -98;
  END;



  FOR n_record IN n_close LOOP
    V_CustID := n_record.CustID;
    V_SNo := n_record.SNo;
    V_Cvtsno := n_record.Cvtsno;
    V_Smartcardno := n_record.Smartcardno;
    V_Chcode := n_record.Chcode;
    V_ChName := n_record.Description;
    V_ChannelID := n_record.ChannelID;
    V_PromCode := n_record.PromCode;
    V_PromName := n_record.PromName;
    V_ClassCode1 := n_record.ClassCode1;
    V_ClassName1 := n_record.ClassName1;
    V_ServiceType := n_record.ServiceType;
    V_Compcode := n_record.Compcode;

    --檢查是否為要關閉的頻道
    FOR I IN 1 .. V_CancelChannelIDCnt LOOP
      V_TempChannelID := CancelChannelIDAry(I);

      IF V_ChannelID=V_TempChannelID THEN
        V_IsClose := 'Y';
        EXIT;
      ELSE
        V_IsClose := 'N';
      END IF;
    END LOOP;

    IF (V_CustID=1427) or (V_CustID=140326) THEN
      V_CustID := V_CustID;
    END IF;

    IF V_IsClose='Y' THEN--若為要關閉的頻道

      IF V_TempSmartCardNo = V_Smartcardno THEN
         V_ChannelIDStrList := V_ChannelIDStrList||','||V_ChannelID;
         V_ChCodeStrList := V_ChCodeStrList||','||V_Chcode||'-'||V_ChName;
      ELSE
        --IF n_close%rowcount!=1 THEN
        IF (V_ChannelIDStrList <> '') OR (V_ChannelIDStrList <> ' ') OR (V_ChannelIDStrList IS NOT NULL) THEN
          --p_Option = 1則將關頻道資料正式存入資料庫,
          --若p_Option <> 1則僅將要關閉頻道資料LOG至SO080
          IF p_Option='1' THEN
            BEGIN
              Insert into Com.Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,NOTES,CMD_STATUS,
                     OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES,PROCESSINGDATE)
              Values(V_HIGHT_LEVEL_CMD_ID,V_TempSmartCardNo,V_ChannelIDStrList,V_CMD_STATUS,
                     V_OPERATOR,V_TIME,Com.S_SENDNAGRA_SEQNO.NEXTVAL,V_Compcode,'',V_TempProcessingDate);
            EXCEPTION
              WHEN others THEN
                p_RetMsg := V_Content || '---INSERT Send_Nagra時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
                BEGIN
                  Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
                  Values('',V_CustID,V_ExceptionType1,V_SNo,p_RetMsg,V_Compcode,V_ServiceType);
                EXCEPTION
                  WHEN others THEN
                    p_RetCode := -1;
                END;
            END;
          END IF;


          BEGIN
            --LOG資料有……客編、卡號、頻道、工單號碼、業務活動、客戶類別
            --Content= 卡號;業務活動;客戶類別;要關閉的頻道List
            V_Content := V_TempSmartcardno || ';'
                         || V_TempPromCode || '-' || V_TempPromName ||';'
                         || V_TempClassCode1 || '-' || V_TempClassName1 || ';'
                         || V_ChCodeStrList || ';'
                         || V_ChannelIDStrList;

            Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
            Values('',V_TempCustID,V_CancelChannelType,V_TempSNo,V_Content,V_TempCompcode,V_TempServiceType);

            --處理完一個卡號則清空
            V_ChannelIDStrList := '';
            V_ChCodeStrList := '';
          EXCEPTION
            WHEN others THEN
              p_RetMsg := 'INSERT時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||V_SQL;
              p_RetCode := -1;
          END;
          V_count:= V_count+1;
        END IF;
        V_TempSmartCardNo := V_SmartCardNo;
        V_ChannelIDStrList := V_ChannelID;
        V_ChCodeStrList := V_Chcode||'-'||V_ChName;

        --SO080
        V_TempCustID := V_CustID;
        V_TempSNo := V_SNo;
        V_TempCompcode := V_Compcode;
        V_TempServiceType := V_ServiceType;
        V_TempPromCode := V_PromCode;
        V_TempPromName := V_PromName;
        V_TempClassCode1 := V_ClassCode1;
        V_TempClassName1 := V_ClassName1;
        V_TempChcode := V_Chcode;
        V_TempChName := V_ChName;


      END IF;


      --p_Option = 1則將關頻道資料正式存入資料庫,
      --若p_Option <> 1則僅將要關閉頻道資料LOG至SO080
      IF p_Option='1' THEN
        BEGIN
          Delete SO005 Where CustID = V_CustID
                         and Smartcardno = V_TempSmartCardNo
                         and Chcode = V_Chcode;
        EXCEPTION
          WHEN others THEN
            p_RetMsg := 'Smartcardno=' ||V_TempSmartCardNo || ',Chcode=' || V_Chcode || '---' ||
                        'Delete SO005時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
            BEGIN
              Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
              Values('',V_CustID,V_ExceptionType2,V_SNo,p_RetMsg,V_Compcode,V_ServiceType);
            EXCEPTION
              WHEN others THEN
                p_RetCode := -1;
            END;
        END;



        BEGIN
          Insert into SOAC0202(CompCode,CustId,STBSNo,
                             SmartCardNo,ModeType,ChCode,ChName,
                             UpdTime,UpdEn)
                     Values(V_Compcode,V_CustID,V_Cvtsno,
                            V_TempSmartCardNo,V_ModeType,V_Chcode,V_ChName,
                            V_Time,V_OPERATOR);
        EXCEPTION
          WHEN others then
            p_RetMsg := 'INSERT至SOAC0202時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||V_SQL;
            BEGIN
              Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
              Values('',V_CustID,V_ExceptionType2,V_SNo,p_RetMsg,V_Compcode,V_ServiceType);
            EXCEPTION
              WHEN others THEN
                p_RetCode := -1;
            END;
        END;
      END IF;

    END IF;
  END LOOP;



--******************************************************************************
--******************************   處理最後一筆  *******************************
--******************************************************************************
  --IF V_Content <> '' THEN  --如果有要關閉的資料
  IF (V_ChannelIDStrList <> '') OR (V_ChannelIDStrList <> ' ') OR (V_ChannelIDStrList IS NOT NULL) THEN  --如果有要關閉的資料
    --p_Option = 1則將關頻道資料正式存入資料庫,
    --若p_Option <> 1則僅將要關閉頻道資料LOG至SO080
    IF p_Option='1' THEN
      BEGIN
        Insert into Com.Send_Nagra(HIGH_LEVEL_CMD_ID,ICC_NO,NOTES,CMD_STATUS,
               OPERATOR,UPDTIME,SEQNO,COMPCODE,RESENTTIMES,PROCESSINGDATE)
        Values(V_HIGHT_LEVEL_CMD_ID,V_TempSmartCardNo,V_ChannelIDStrList,V_CMD_STATUS,
               V_OPERATOR,V_TIME,Com.S_SENDNAGRA_SEQNO.NEXTVAL,V_Compcode,'',V_TempProcessingDate);
      EXCEPTION
        WHEN others THEN
          p_RetMsg := V_Content || '---INSERT Send_Nagra時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM;
          BEGIN
            Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
            Values('',V_CustID,V_ExceptionType1,V_SNo,p_RetMsg,V_Compcode,V_ServiceType);
          EXCEPTION
            WHEN others THEN
              p_RetCode := -1;
          END;
      END;
    END IF;



    BEGIN
      --LOG資料有……客編、卡號、頻道、工單號碼、業務活動、客戶類別
      --Content= 卡號;業務活動;客戶類別;要關閉的頻道List
      V_Content := V_TempSmartcardno || ';'
                   || V_TempPromCode || '-' || V_TempPromName ||';'
                   || V_TempClassCode1 || '-' || V_TempClassName1 || ';'
                   || V_ChCodeStrList || ';'
                   || V_ChannelIDStrList;

      Insert into SO080(RcdNo,CustId,Type,SNo,Content,CompCode,ServiceType)
      Values('',V_TempCustID,V_CancelChannelType,V_TempSNo,V_Content,V_TempCompcode,V_TempServiceType);
    EXCEPTION
      WHEN others THEN
        p_RetMsg := 'INSERT時發生錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||V_SQL;
        p_RetCode := -1;
    END;
    V_count:= V_count+1;
  END IF;

  COMMIT;

  p_RetMsg:='關閉頻道執行完成,共產生 ' || V_count ||' 筆待關頻道之資料 ';
  p_RetCode := 0;
  RETURN 0;

 EXCEPTION
    WHEN others THEN
     rollback ;
     p_RetMsg := '其他錯誤:  '||' SQLCODE='||SQLCODE||'   '||'   SQLERRM='||SQLERRM|| '      SQL='||V_SQL;
     P_RetCode:=-99;
     RETURN -99;
END ; -- Procedure SF_CancelChannel3
/