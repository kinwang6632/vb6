CREATE OR REPLACE PROCEDURE SP_TRANSFORM_1
  AS
/*
    東森代碼轉檔
    檔名: SP_TRANSFORM_1.sql
    By: Jackal   Date: 2003.12.27

*/
--declare
  --系統參數(常數)
  C_CodeMapErrorType number :=1;          --對照表有誤
  C_UpdDataErrorType number :=2;          --UPDATE資料檔有誤
  C_OtherErrorType   number :=9;          --其他異常錯誤
  C_StopFlag         number :=0;          --是否停用,預設否
  C_ServiceType      varchar2(3) := 'C';  --業者服務種類,預設為'C'
  C_PrintWONow       number :=0;          --是否立即列印工單,預設否

  v_SQL1           varchar2(1000);
  v_HaveRBS27      number :=0;   --0有RBS27   1沒有
  v_ErrDesc        varchar2(1000);


  c39		char(1) := chr(39);
  v_OldCode         number;
  v_OldName         varchar2(20);
  v_NewCode         number;
  v_NewName         varchar2(20);
  v_Mark            number;
  v_MaxOldCode      number;
  v_OKCounts        number := 0;
  v_ErrorCounts     number := 0;
  v_StartTime       Date;
  v_TempCounts      number;
  v_AllTiem         varchar2(100);
  v_TotalStartTime  Date;
  v_TotalAllTiem    varchar2(100);
  
  v_Note            varchar2(500);



  v_NewCode1        number;
  v_NewCode2        number;
  v_NewCode3        number;
  v_NewCode4        number;
  v_NewCode5        number;
  v_NewCode6        number;
  v_NewCode7        number;
  v_NewCode8        number;
  v_NewCode9        number;
  v_NewCode10       number;


  v_NewName1        varchar2(20);
  v_NewName2        varchar2(20);
  v_NewName3        varchar2(20);
  v_NewName4        varchar2(20);
  v_NewName5        varchar2(20);
  v_NewName6        varchar2(20);
  v_NewName7        varchar2(20);
  v_NewName8        varchar2(20);
  v_NewName9        varchar2(20);
  v_NewName10       varchar2(20);


--資料檔轉換用
  cursor ccCD008C is
    select Rowid,CodeNo,ServiceCode from CD008C;
    
  cursor ccCD057 is
    select Rowid,CodeNo,ReturnCode,ReturnName from CD057;
    
  cursor ccSO001 is
    select Rowid,CustId,CompCode,PwCode,PwName,OrgCode,OrgName from SO001;
    
  cursor ccSO002 is
    select Rowid,CustId,CompCode,ServiceType,MediaCode,MediaName,
           DelCode,DelName,PRCode,PRName,StopCode,StopName from SO002;
    
  cursor ccSO004 is
    select Rowid,SEQNo,MediaCode,MediaName from SO004;
    
  cursor ccSO006 is
    select Rowid,SEQNo,ServiceCode,ServiceName,ProcResultNo,ProcResultName,
           ServDescCode,ServDescName,MediaCode,MediaName,
           BulletinCode,BulletinName from SO006;
           
  cursor ccSO007 is
    select Rowid,SNo,MediaCode,MediaName,ReturnCode,ReturnName,
           SatiCode,SatiName,BulletinCode,BulletinName from SO007;
           
  cursor ccSO008 is
    select Rowid,SNo,ReturnCode,ReturnName,SatiCode,SatiName from SO008;
           
  cursor ccSO009 is
    select Rowid,SNo,ReasonCode,ReasonName,
           ReturnCode,ReturnName,SatiCode,SatiName from SO009;
           
  cursor ccSO013 is
    select Rowid,IntroId,CompCode,MediaCode,MediaName from SO013;
    
  cursor ccSO014 is
    select Rowid,CompCode,AddrNo,BTCode,BTName from SO014;
    
  cursor ccSO016 is
    select Rowid,StrtCode,Ord,BTCode,BTName from SO016;
    
  cursor ccSO032 is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO032;
           
  cursor ccSO033 is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO033;
           
  cursor ccSO033DEBT is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO033DEBT;
           
           
  cursor ccSO034 is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO034;
           
  cursor ccSO035 is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO035;
           
  cursor ccSO036 is
    select Rowid,BillNo,Item,UCCode,UCName,STCode,STName,
           CancelCode,CancelName from SO036;
           
  cursor ccSO044 is
    select Rowid,CompCode,ServiceType,UCCode,PWCode from SO044;
    
  cursor ccSO050 is
    select Rowid,UCCode,UCName,NSTCode,NSTName,STCode,STName,
           CancelCode,CancelName from SO050;
           
  cursor ccSO051 is
    select Rowid,UCCode,UCName,UCCodeB,UCNameB,STCode,STName,
                 STCodeB,STNameB from SO051;
                 
                 
  cursor ccSO055 is
    select Rowid,MediaCode,MediaName,MediaCodeB,MediaNameB from SO055;
    
    
  cursor ccSO067 is
    select Rowid,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO067;
    
  cursor ccSO074 is
    select Rowid,BillNo,Item,STCode,STName,CancelCode,CancelName from SO074;
    
  cursor ccSO074A is
    select Rowid,BillNo,Item,STCode,STName,CancelCode,CancelName from SO074A;
    
  cursor ccSO077 is
    select Rowid,BillNo,Item,STCode,STName,CancelCode,CancelName from SO077;
    
  cursor ccSO095 is
    select Rowid,BulletinCode,BulletinName from SO095;
    
  cursor ccSO096 is
    select Rowid,BulletinCode,BulletinName from SO096;
    
  cursor ccSO098 is
    select Rowid,MediaCode,MediaName,BulletinCode,BulletinName from SO098;
    
  cursor ccSO101 is
    select Rowid,DelCode,DelName from SO101;
    
  cursor ccSO105 is
    select Rowid,OrderNo,ServiceCode,ServDescCode,MediaCode,MediaName,
           ReturnCode,ReturnName,CATVReturnCode,CATVReturnName,
           DVSReturnCode,DVSReturnName,BulletinCode from SO105;
           
  cursor ccSO105T is
    select Rowid,OrderNo,CompCode,MediaCode,MediaName,ReturnCode,ReturnName,
           CATVReturnCode,CATVReturnName,CMReturnCode,CMReturnName,
           DVSReturnCode,DVSReturnName,BulletinCode,BulletinName from SO105T;
           
  cursor ccSO106 is
    select Rowid,MediaCode,MediaName from SO106;
    
  cursor ccSO121 is
    select Rowid,MediaCode,MediaName from SO121;
    
  cursor ccSO122 is
    select Rowid,MediaCode,MediaName from SO122;
    
  cursor ccSO131 is
    select Rowid,MediaCode,MediaName from SO131;
    
  cursor ccSO134 is
    select Rowid,MediaCode,MediaName from SO134;
    
  cursor ccSO503 is
    select Rowid,MediaCode,MediaName from SO503;
    
    
    
--0000



--代碼檔轉換用
/*
  cursor ccCD005 is
    select Rowid,CodeNo from CD005 ORDER BY CodeNo;
  cursor ccCD005_MAP is
    select NewCode,NewName from CD005_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD006 is
    select Rowid,CodeNo from CD006 ORDER BY CodeNo;
  cursor ccCD006_MAP is
    select NewCode,NewName from CD006_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD007 is
    select Rowid,CodeNo from CD007 ORDER BY CodeNo;
  cursor ccCD007_MAP is
    select NewCode,NewName from CD007_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD008 is
    select Rowid,CodeNo from CD008 ORDER BY CodeNo;
  cursor ccCD008_MAP is
    select NewCode,NewName from CD008_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD008A is
    select Rowid,CodeNo from CD008A ORDER BY CodeNo;
  cursor ccCD008A_MAP is
    select NewCode,NewName from CD008A_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD008C is
    select Rowid,CodeNo from CD008C ORDER BY CodeNo;
  cursor ccCD008C_MAP is
    select NewCode,NewName from CD008C_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD009 is
    select Rowid,CodeNo from CD009 ORDER BY CodeNo;
  cursor ccCD009_MAP is
    select NewCode,NewName from CD009_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD012 is
    select Rowid,CodeNo from CD012 ORDER BY CodeNo;
  cursor ccCD012_MAP is
    select NewCode,NewName from CD012_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD013 is
    select Rowid,CodeNo from CD013 ORDER BY CodeNo;
  cursor ccCD013_MAP is
    select NewCode,NewName from CD013_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD014 is
    select Rowid,CodeNo from CD014 ORDER BY CodeNo;
  cursor ccCD014_MAP is
    select NewCode,NewName from CD014_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD015 is
    select Rowid,CodeNo from CD015 ORDER BY CodeNo;
  cursor ccCD015_MAP is
    select NewCode,NewName from CD015_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD016 is
    select Rowid,CodeNo from CD016 ORDER BY CodeNo;
  cursor ccCD016_MAP is
    select NewCode,NewName from CD016_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD020 is
    select Rowid,CodeNo from CD020 ORDER BY CodeNo;
  cursor ccCD020_MAP is
    select NewCode,NewName from CD020_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD021 is
    select Rowid,CodeNo from CD021 ORDER BY CodeNo;
  cursor ccCD021_MAP is
    select NewCode,NewName from CD021_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD025 is
    select Rowid,CodeNo from CD025 ORDER BY CodeNo;
  cursor ccCD025_MAP is
    select NewCode,NewName from CD025_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD026 is
    select Rowid,CodeNo from CD026 ORDER BY CodeNo;
  cursor ccCD026_MAP is
    select NewCode,NewName from CD026_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD049 is
    select Rowid,CodeNo from CD049 ORDER BY CodeNo;
  cursor ccCD049_MAP is
    select NewCode,NewName from CD049_MAP WHERE OldCode IS NULL ORDER BY NewCode;


  cursor ccCD051 is
    select Rowid,CodeNo from CD051 ORDER BY CodeNo;
  cursor ccCD051_MAP is
    select NewCode,NewName from CD051_MAP WHERE OldCode IS NULL ORDER BY NewCode;
*/


--宣告各代碼陣列
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  TYPE Varchar2Ary IS TABLE OF varchar2(50) INDEX BY BINARY_INTEGER;

  CD008CodeMap NumberAry;
  CD008Mark NumberAry;
  CD008NameMap Varchar2Ary;

  CD008ACodeMap NumberAry;
  CD008AMark NumberAry;
  CD008ANameMap Varchar2Ary;

  CD008CCodeMap NumberAry;
  CD008CMark NumberAry;
  CD008CNameMap Varchar2Ary;

  CD009CodeMap NumberAry;
  CD009Mark NumberAry;
  CD009NameMap Varchar2Ary;

  CD012CodeMap NumberAry;
  CD012Mark NumberAry;
  CD012NameMap Varchar2Ary;

  CD013CodeMap NumberAry;
  CD013Mark NumberAry;
  CD013NameMap Varchar2Ary;

  CD014CodeMap NumberAry;
  CD014Mark NumberAry;
  CD014NameMap Varchar2Ary;

  CD015CodeMap NumberAry;
  CD015Mark NumberAry;
  CD015NameMap Varchar2Ary;

  CD016CodeMap NumberAry;
  CD016Mark NumberAry;
  CD016NameMap Varchar2Ary;

  CD020CodeMap NumberAry;
  CD020Mark NumberAry;
  CD020NameMap Varchar2Ary;

  CD021CodeMap NumberAry;
  CD021Mark NumberAry;
  CD021NameMap Varchar2Ary;

  CD025CodeMap NumberAry;
  CD025Mark NumberAry;
  CD025NameMap Varchar2Ary;

  CD026CodeMap NumberAry;
  CD026Mark NumberAry;
  CD026NameMap Varchar2Ary;

  CD049CodeMap NumberAry;
  CD049Mark NumberAry;
  CD049NameMap Varchar2Ary;

  CD051CodeMap NumberAry;
  CD051Mark NumberAry;
  CD051NameMap Varchar2Ary;

  TYPE CurTyp IS REF CURSOR;  --自訂cursor型態, 供新版dynamic SQL
  v_DyCursor1 CurTyp;


--對於傳入特殊的Index做處理
  FUNCTION GetNewCode(p_OldCode Number,p_CodeMapArray IN NumberAry)
  RETURN number
  AS
    L_Result Number;
  BEGIN
    IF p_OldCode IS NULL THEN
      RETURN NULL;
    ELSIF p_OldCode<=0 THEN
      RETURN p_OldCode;
    ELSE
      BEGIN
        L_Result := p_CodeMapArray(p_OldCode);
        RETURN L_Result;
      EXCEPTION
        WHEN OTHERS then
          RETURN p_OldCode;
      END;
    END IF;
  END;


  --對於傳入特殊的Index做處理
  FUNCTION GetNewName(p_OldCode Number,p_OldName varchar2,p_NameMapArray IN Varchar2Ary)
  RETURN varchar2
  AS
    L_Result Varchar2(100);
  BEGIN
    IF p_OldCode IS NULL THEN
      RETURN NULL;
    ELSIF p_OldCode<=0 THEN
      RETURN NULL;
    ELSE
      BEGIN
        L_Result := p_NameMapArray(p_OldCode);
        IF L_Result IS NULL THEN
          RETURN p_OldName;--若沒有對應的資料,回傳舊名稱
        ELSE
          RETURN L_Result;
        END IF;
      EXCEPTION
        WHEN OTHERS then
          RETURN p_OldName;
      END;
    END IF;
  END;


  --對於傳入特殊的Index做處理
  FUNCTION GetMark(p_OldCode Number,p_MarkArray IN NumberAry)
  RETURN varchar2
  AS
  BEGIN
    IF p_OldCode IS NULL THEN
      RETURN NULL;
    ELSIF p_OldCode<=0 THEN
      RETURN NULL;
    ELSE
      RETURN p_MarkArray(p_OldCode);
    END IF;
  END;


  --計算兩時間差
  FUNCTION GetDiffTime(p_StartDate Date,p_StopDate Date)
  RETURN varchar2
  AS
    L_DiffTime  Number;
    L_HH        Number;
    L_MM        Number;
    L_SS        Number;
    L_SS1       Varchar2(100);
    L_MM1       Varchar2(100);
    L_HH1       Varchar2(100);
  BEGIN
    L_DiffTime := (p_StopDate-p_StartDate)*86400;
    L_HH := ROUND(trunc(L_DiffTime/3600),1);
    L_HH1 := TO_CHAR(L_HH);
    L_SS := ROUND(mod(L_DiffTime, 60),1);
    L_SS1 := TO_CHAR(L_SS);
    L_MM := ROUND((L_DiffTime - L_HH*3600 - L_SS) / 60,1);
    L_MM1 := TO_CHAR(L_MM);
    RETURN '合計執行時間: ' || L_HH1 || '時' || L_MM1 || '分' || L_SS1 || '秒';
  END;


  --對於傳入特殊的Index做處理
  FUNCTION SetSegment(p_HaveRBS27 number)
  RETURN NUMBER
  AS
  BEGIN
    IF p_HaveRBS27=1 THEN
      BEGIN
        DBMS_UTILITY.EXEC_DDL_STATEMENT('set transaction use rollback segment rbS27');
        RETURN 1;
      EXCEPTION
        WHEN OTHERS then
          RETURN 0;
      END;
    ELSE
      RETURN 0;
    END IF;
  END;

BEGIN
  v_TotalStartTime := SysDate;
  DBMS_OUTPUT.PUT_LINE('開始執行:' || to_char(v_TotalStartTime,'YYYY/MM/DD HH24:MI:SS'));

  --判定是ORACLE有無RBS27
  v_SQL1 := 'select COUNT(segment_name) from dba_rollback_segs WHERE segment_name=''RBS27''';
  BEGIN
    OPEN v_DyCursor1 FOR v_SQL1;
  EXCEPTION
    WHEN others THEN
      DBMS_OUTPUT.PUT_LINE('dba_rollback_segs 沒開放權限');
      RETURN;
  END;


  LOOP
    FETCH v_DyCursor1 INTO v_HaveRBS27;
    EXIT WHEN v_DyCursor1%NOTFOUND;

    IF v_HaveRBS27=1 THEN--8i將"RBS27" ONLINE
      BEGIN
        DBMS_UTILITY.EXEC_DDL_STATEMENT('ALTER ROLLBACK SEGMENT "RBS27" ONLINE');
      EXCEPTION
        WHEN others THEN
          NULL;
      END;
    END IF;
  END LOOP;

  CLOSE v_DyCursor1;



--將各代碼陣列填值
  -- 找尋 CD008 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD008;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD008_MAP WHERE OldCode=I;

        --將對照表CD008_MAP轉成陣列
        CD008CodeMap(I) := v_NewCode;
        CD008NameMap(I) := v_NewName;
        CD008Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD008_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD008_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD008CodeMap(I) := I;
          CD008NameMap(I) := null;
          CD008Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD008A 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD008A;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD008A_MAP WHERE OldCode=I;

        --將對照表CD008A_MAP轉成陣列
        CD008ACodeMap(I) := v_NewCode;
        CD008ANameMap(I) := v_NewName;
        CD008AMark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD008A_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD008A_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD008ACodeMap(I) := I;
          CD008ANameMap(I) := null;
          CD008AMark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;
  
  
  -- 找尋 CD008C 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD008C;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD008C_MAP WHERE OldCode=I;

        --將對照表CD008C_MAP轉成陣列
        CD008CCodeMap(I) := v_NewCode;
        CD008CNameMap(I) := v_NewName;
        CD008CMark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD008C_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD008C_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD008CCodeMap(I) := I;
          CD008CNameMap(I) := null;
          CD008CMark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- 找尋 CD009 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD009;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD009_MAP WHERE OldCode=I;

        --將對照表CD009_MAP轉成陣列
        CD009CodeMap(I) := v_NewCode;
        CD009NameMap(I) := v_NewName;
        CD009Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD009_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD009_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD009CodeMap(I) := I;
          CD009NameMap(I) := null;
          CD009Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD012 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD012;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD012_MAP WHERE OldCode=I;

        --將對照表CD012_MAP轉成陣列
        CD012CodeMap(I) := v_NewCode;
        CD012NameMap(I) := v_NewName;
        CD012Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD012_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD012_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD012CodeMap(I) := I;
          CD012NameMap(I) := null;
          CD012Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD013 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD013;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD013_MAP WHERE OldCode=I;

        --將對照表CD013_MAP轉成陣列
        CD013CodeMap(I) := v_NewCode;
        CD013NameMap(I) := v_NewName;
        CD013Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD013_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD013_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD013CodeMap(I) := I;
          CD013NameMap(I) := null;
          CD013Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD014 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD014;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD014_MAP WHERE OldCode=I;

        --將對照表CD014_MAP轉成陣列
        CD014CodeMap(I) := v_NewCode;
        CD014NameMap(I) := v_NewName;
        CD014Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD014_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD014_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD014CodeMap(I) := I;
          CD014NameMap(I) := null;
          CD014Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD015 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD015;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD015_MAP WHERE OldCode=I;

        --將對照表CD015_MAP轉成陣列
        CD015CodeMap(I) := v_NewCode;
        CD015NameMap(I) := v_NewName;
        CD015Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD015_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD015_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD015CodeMap(I) := I;
          CD015NameMap(I) := null;
          CD015Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD016 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD016;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD016_MAP WHERE OldCode=I;

        --將對照表CD016_MAP轉成陣列
        CD016CodeMap(I) := v_NewCode;
        CD016NameMap(I) := v_NewName;
        CD016Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD016_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD016_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD016CodeMap(I) := I;
          CD016NameMap(I) := null;
          CD016Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD020 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD020;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD020_MAP WHERE OldCode=I;

        --將對照表CD020_MAP轉成陣列
        CD020CodeMap(I) := v_NewCode;
        CD020NameMap(I) := v_NewName;
        CD020Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD020_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD020_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD020CodeMap(I) := I;
          CD020NameMap(I) := null;
          CD020Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD021 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD021;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD021_MAP WHERE OldCode=I;

        --將對照表CD021_MAP轉成陣列
        CD021CodeMap(I) := v_NewCode;
        CD021NameMap(I) := v_NewName;
        CD021Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD021_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD021_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD021CodeMap(I) := I;
          CD021NameMap(I) := null;
          CD021Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;
  
  
  -- 找尋 CD025 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD025;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD025_MAP WHERE OldCode=I;

        --將對照表CD025_MAP轉成陣列
        CD025CodeMap(I) := v_NewCode;
        CD025NameMap(I) := v_NewName;
        CD025Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD025_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD025_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD025CodeMap(I) := I;
          CD025NameMap(I) := null;
          CD025Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- 找尋 CD026 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD026;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD026_MAP WHERE OldCode=I;

        --將對照表CD026_MAP轉成陣列
        CD026CodeMap(I) := v_NewCode;
        CD026NameMap(I) := v_NewName;
        CD026Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD026_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD026_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD026CodeMap(I) := I;
          CD026NameMap(I) := null;
          CD026Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- 找尋 CD049 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD049;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD049_MAP WHERE OldCode=I;

        --將對照表CD049_MAP轉成陣列
        CD049CodeMap(I) := v_NewCode;
        CD049NameMap(I) := v_NewName;
        CD049Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD049_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD049_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD049CodeMap(I) := I;
          CD049NameMap(I) := null;
          CD049Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- 找尋 CD051 最大的舊代碼值,要作為陣列的最大Index
  Select Max(CodeNo) into v_MaxOldCode from CD051;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD051_MAP WHERE OldCode=I;

        --將對照表CD051_MAP轉成陣列
        CD051CodeMap(I) := v_NewCode;
        CD051NameMap(I) := v_NewName;
        CD051Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'舊代碼[ '||I||' ]在對照檔中有兩筆以上','CD051_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('舊代碼[ '||I||' ]在CD051_MAP對照檔中有兩筆以上');
          RETURN;
        WHEN OTHERS then
          CD051CodeMap(I) := I;
          CD051NameMap(I) := null;
          CD051Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



--*****************************************************************
--*******************   更  新  資  料  檔   **********************
--*****************************************************************
  --更新CD008C資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD008C';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD008C in ccCD008C loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD008C.ServiceCode,CD008CodeMap);


        UPDATE CD008C set ServiceCode=v_NewCode1
                     WHERE Rowid=crCD008C.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔CD008C有誤  ServiceCode: '||
                   crCD008C.ServiceCode||'  CodeNo='||crCD008C.CodeNo,'CD008C');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD008C','資料檔CD008C代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  
  --更新CD057資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD057';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD057 in ccCD057 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD057.ReturnCode,CD015CodeMap);
        v_NewName1 := GetNewName(crCD057.ReturnCode,crCD057.ReturnName,CD015NameMap);


        UPDATE CD057 set ReturnCode=v_NewCode1,ReturnName=v_NewName1
                     WHERE Rowid=crCD057.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔CD057有誤  CodeNo: '||crCD057.CodeNo,'CD057');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD057','資料檔CD057代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  
  
  --更新SO001資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO001';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO001 in ccSO001 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO001.PwCode,CD020CodeMap);
        v_NewName1 := GetNewName(crSO001.PwCode,crSO001.PwName,CD020NameMap);

        v_NewCode2 := GetNewCode(crSO001.OrgCode,CD025CodeMap);
        v_NewName2 := GetNewName(crSO001.OrgCode,crSO001.OrgName,CD025NameMap);


        UPDATE SO001 set PwCode=v_NewCode1,PwName=v_NewName1,
                         OrgCode=v_NewCode2,OrgName=v_NewName2
                     WHERE Rowid=crSO001.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO001有誤  CustId: '||crSO001.CustId||' CompCode:'||crSO001.CompCode,'SO001');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO001','資料檔SO001代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  
  --更新SO002資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO002';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO002 in ccSO002 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO002.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO002.MediaCode,crSO002.MediaName,CD009NameMap);

        v_NewCode2 := GetNewCode(crSO002.DelCode,CD012CodeMap);
        v_NewName2 := GetNewName(crSO002.DelCode,crSO002.DelName,CD012NameMap);

        v_NewCode3 := GetNewCode(crSO002.PRCode,CD014CodeMap);
        v_NewName3 := GetNewName(crSO002.PRCode,crSO002.PRName,CD014NameMap);

        v_NewCode4 := GetNewCode(crSO002.StopCode,CD014CodeMap);
        v_NewName4 := GetNewName(crSO002.StopCode,crSO002.StopName,CD014NameMap);


        UPDATE SO002 set MediaCode=v_NewCode1,MediaName=v_NewName1,
                         DelCode=v_NewCode2,DelName=v_NewName2,
                         PRCode=v_NewCode3,PRName=v_NewName3,
                         StopCode=v_NewCode4,StopName=v_NewName4
                     WHERE Rowid=crSO002.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO002有誤  CustId: '||crSO002.CustId||' CompCode:'||crSO002.CompCode||'  ServiceType:'||crSO002.ServiceType,'SO002');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO002','資料檔SO002代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  
  --更新SO004資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO004';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO004 in ccSO004 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO004.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO004.MediaCode,crSO004.MediaName,CD009NameMap);


        UPDATE SO004 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO004.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO004有誤  SEQNo: '||crSO004.SEQNo,'SO004');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO004','資料檔SO004代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  
  --更新SO006資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO006';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO006 in ccSO006 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO006.ServiceCode,CD008CodeMap);
        v_NewName1 := GetNewName(crSO006.ServiceCode,crSO006.ServiceName,CD008NameMap);

        v_NewCode2 := GetNewCode(crSO006.ProcResultNo,CD008ACodeMap);
        v_NewName2 := GetNewName(crSO006.ProcResultNo,crSO006.ProcResultName,CD008ANameMap);

        v_NewCode3 := GetNewCode(crSO006.ServDescCode,CD008ACodeMap);
        v_NewName3 := GetNewName(crSO006.ServDescCode,crSO006.ServDescName,CD008ANameMap);

        v_NewCode4 := GetNewCode(crSO006.MediaCode,CD009CodeMap);
        v_NewName4 := GetNewName(crSO006.MediaCode,crSO006.MediaName,CD009NameMap);

        v_NewCode5 := GetNewCode(crSO006.BulletinCode,CD049CodeMap);
        v_NewName5 := GetNewName(crSO006.BulletinCode,crSO006.BulletinName,CD049NameMap);


        UPDATE SO006 set ServiceCode=v_NewCode1,ServiceName=v_NewName1,
                         ProcResultNo=v_NewCode2,ProcResultName=v_NewName2,
                         ServDescCode=v_NewCode3,ServDescName=v_NewName3,
                         MediaCode=v_NewCode4,MediaName=v_NewName4,
                         BulletinCode=v_NewCode5,BulletinName=v_NewName5
                     WHERE Rowid=crSO006.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO006有誤  SEQNo: '||crSO006.SEQNo,'SO006');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO006','資料檔SO006代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO007資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO007';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO007 in ccSO007 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO007.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO007.MediaCode,crSO007.MediaName,CD009NameMap);

        v_NewCode2 := GetNewCode(crSO007.ReturnCode,CD015CodeMap);
        v_NewName2 := GetNewName(crSO007.ReturnCode,crSO007.ReturnName,CD015NameMap);

        v_NewCode3 := GetNewCode(crSO007.SatiCode,CD026CodeMap);
        v_NewName3 := GetNewName(crSO007.SatiCode,crSO007.SatiName,CD026NameMap);

        v_NewCode4 := GetNewCode(crSO007.BulletinCode,CD049CodeMap);
        v_NewName4 := GetNewName(crSO007.BulletinCode,crSO007.BulletinName,CD049NameMap);


        UPDATE SO007 set MediaCode=v_NewCode1,MediaName=v_NewName1,
                         ReturnCode=v_NewCode2,ReturnName=v_NewName2,
                         SatiCode=v_NewCode3,SatiName=v_NewName3,
                         BulletinCode=v_NewCode4,BulletinName=v_NewName4
                     WHERE Rowid=crSO007.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO007有誤  SNo: '||crSO007.SNo,'SO007');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO007','資料檔SO007代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  

  --更新SO008資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO008';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO008 in ccSO008 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO008.ReturnCode,CD015CodeMap);
        v_NewName1 := GetNewName(crSO008.ReturnCode,crSO008.ReturnName,CD015NameMap);

        v_NewCode2 := GetNewCode(crSO008.SatiCode,CD026CodeMap);
        v_NewName2 := GetNewName(crSO008.SatiCode,crSO008.SatiName,CD026NameMap);


        UPDATE SO008 set ReturnCode=v_NewCode1,ReturnName=v_NewName1,
                         SatiCode=v_NewCode2,SatiName=v_NewName2
                     WHERE Rowid=crSO008.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO008有誤  SNo: '||crSO008.SNo,'SO008');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO008','資料檔SO008代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO009資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO009';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO009 in ccSO009 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO009.ReasonCode,CD014CodeMap);
        v_NewName1 := GetNewName(crSO009.ReasonCode,crSO009.ReasonName,CD014NameMap);

        v_NewCode2 := GetNewCode(crSO009.ReturnCode,CD015CodeMap);
        v_NewName2 := GetNewName(crSO009.ReturnCode,crSO009.ReturnName,CD015NameMap);

        v_NewCode3 := GetNewCode(crSO009.SatiCode,CD026CodeMap);
        v_NewName3 := GetNewName(crSO009.SatiCode,crSO009.SatiName,CD026NameMap);


        UPDATE SO009 set ReasonCode=v_NewCode1,ReasonName=v_NewName1,
                         ReturnCode=v_NewCode2,ReturnName=v_NewName2,
                         SatiCode=v_NewCode3,SatiName=v_NewName3
                     WHERE Rowid=crSO009.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO009有誤  SNo: '||crSO009.SNo,'SO009');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO009','資料檔SO009代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO013資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO013';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO013 in ccSO013 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO013.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO013.MediaCode,crSO013.MediaName,CD009NameMap);


        UPDATE SO013 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO013.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO013有誤  IntroId: '||crSO013.IntroId||'  CompCode:'||crSO013.CompCode,'SO013');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO013','資料檔SO013代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO014資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO014';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO014 in ccSO014 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO014.BTCode,CD021CodeMap);
        v_NewName1 := GetNewName(crSO014.BTCode,crSO014.BTName,CD021NameMap);


        UPDATE SO014 set BTCode=v_NewCode1,BTName=v_NewName1
                     WHERE Rowid=crSO014.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO014有誤  AddrNo: '||crSO014.AddrNo||'  CompCode:'||crSO014.CompCode,'SO014');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO014','資料檔SO014代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO016資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO016';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO016 in ccSO016 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO016.BTCode,CD021CodeMap);
        v_NewName1 := GetNewName(crSO016.BTCode,crSO016.BTName,CD021NameMap);


        UPDATE SO016 set BTCode=v_NewCode1,BTName=v_NewName1
                     WHERE Rowid=crSO016.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO016有誤  StrtCode: '||crSO016.StrtCode||'  Ord:'||crSO016.Ord,'SO016');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO016','資料檔SO016代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO032資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO032';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO032 in ccSO032 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO032.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO032.UCCode,crSO032.UCName,CD013NameMap);

        v_NewCode2 := GetNewCode(crSO032.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO032.STCode,crSO032.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO032.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO032.CancelCode,crSO032.CancelName,CD051NameMap);


        UPDATE SO032 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO032.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO032有誤  BillNo: '||crSO032.BillNo||'  Item:'||crSO032.Item,'SO032');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO032','資料檔SO032代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO033資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO033';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO033 in ccSO033 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO033.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO033.UCCode,crSO033.UCName,CD013NameMap);
        v_Note := '{舊的未收原因='||crSO033.UCName||'('||crSO033.UCCode||')}';

        v_NewCode2 := GetNewCode(crSO033.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO033.STCode,crSO033.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO033.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO033.CancelCode,crSO033.CancelName,CD051NameMap);


        UPDATE SO033 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3,
                         Note=Note||v_Note
                     WHERE Rowid=crSO033.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO033有誤  BillNo: '||crSO033.BillNo||'  Item:'||crSO033.Item,'SO033');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO033','資料檔SO033代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO033DEBT資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO033DEBT';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO033DEBT in ccSO033DEBT loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO033DEBT.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO033DEBT.UCCode,crSO033DEBT.UCName,CD013NameMap);

        v_NewCode2 := GetNewCode(crSO033DEBT.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO033DEBT.STCode,crSO033DEBT.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO033DEBT.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO033DEBT.CancelCode,crSO033DEBT.CancelName,CD051NameMap);


        UPDATE SO033DEBT set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO033DEBT.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO033DEBT有誤  BillNo: '||crSO033DEBT.BillNo||'  Item:'||crSO033DEBT.Item,'SO033DEBT');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO033DEBT','資料檔SO033DEBT代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO034資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO034';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO034 in ccSO034 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO034.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO034.UCCode,crSO034.UCName,CD013NameMap);
        v_Note := '{舊的未收原因='||crSO034.UCName||'('||crSO034.UCCode||')}';

        v_NewCode2 := GetNewCode(crSO034.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO034.STCode,crSO034.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO034.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO034.CancelCode,crSO034.CancelName,CD051NameMap);


        UPDATE SO034 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3,
                         Note=Note||v_Note
                     WHERE Rowid=crSO034.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO034有誤  BillNo: '||crSO034.BillNo||'  Item:'||crSO034.Item,'SO034');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO034','資料檔SO034代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO035資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO035';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO035 in ccSO035 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO035.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO035.UCCode,crSO035.UCName,CD013NameMap);
        v_Note := '{舊的未收原因='||crSO035.UCName||'('||crSO035.UCCode||')}';

        v_NewCode2 := GetNewCode(crSO035.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO035.STCode,crSO035.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO035.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO035.CancelCode,crSO035.CancelName,CD051NameMap);


        UPDATE SO035 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3,
                         Note=Note||v_Note
                     WHERE Rowid=crSO035.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO035有誤  BillNo: '||crSO035.BillNo||'  Item:'||crSO035.Item,'SO035');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO035','資料檔SO035代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO036資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO036';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO036 in ccSO036 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO036.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO036.UCCode,crSO036.UCName,CD013NameMap);
        v_Note := '{舊的未收原因='||crSO036.UCName||'('||crSO036.UCCode||')}';

        v_NewCode2 := GetNewCode(crSO036.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO036.STCode,crSO036.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO036.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO036.CancelCode,crSO036.CancelName,CD051NameMap);


        UPDATE SO036 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3,
                         Note=Note||v_Note
                     WHERE Rowid=crSO036.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO036有誤  BillNo: '||crSO036.BillNo||'  Item:'||crSO036.Item,'SO036');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO036','資料檔SO036代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO044資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO044';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO044 in ccSO044 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO044.UCCode,CD013CodeMap);

        v_NewCode2 := GetNewCode(crSO044.PWCode,CD020CodeMap);


        UPDATE SO044 set UCCode=v_NewCode1,
                         PWCode=v_NewCode2
                     WHERE Rowid=crSO044.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO044有誤  CompCode: '||crSO044.CompCode||'  ServiceType:'||crSO044.ServiceType,'SO044');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO044','資料檔SO044代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO050資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO050';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO050 in ccSO050 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO050.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO050.UCCode,crSO050.UCName,CD013NameMap);

        v_NewCode2 := GetNewCode(crSO050.NSTCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO050.NSTCode,crSO050.NSTName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO050.STCode,CD016CodeMap);
        v_NewName3 := GetNewName(crSO050.STCode,crSO050.STName,CD016NameMap);

        v_NewCode4 := GetNewCode(crSO050.CancelCode,CD051CodeMap);
        v_NewName4 := GetNewName(crSO050.CancelCode,crSO050.CancelName,CD051NameMap);


        UPDATE SO050 set UCCode=v_NewCode1,UCName=v_NewName1,
                         NSTCode=v_NewCode2,NSTName=v_NewName2,
                         STCode=v_NewCode3,STName=v_NewName3,
                         CancelCode=v_NewCode4,CancelName=v_NewName4
                     WHERE Rowid=crSO050.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO050有誤  UCCode: '||crSO050.UCCode||' NSTCode:'||crSO050.NSTCode||'  STCode:'||crSO050.STCode||'  CancelCode:'||crSO050.CancelCode,'SO050');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO050','資料檔SO050代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO051資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO051';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO051 in ccSO051 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO051.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO051.UCCode,crSO051.UCName,CD013NameMap);

        v_NewCode2 := GetNewCode(crSO051.UCCodeB,CD013CodeMap);
        v_NewName2 := GetNewName(crSO051.UCCodeB,crSO051.UCNameB,CD013NameMap);

        v_NewCode3 := GetNewCode(crSO051.STCode,CD016CodeMap);
        v_NewName3 := GetNewName(crSO051.STCode,crSO051.STName,CD016NameMap);

        v_NewCode4 := GetNewCode(crSO051.STCodeB,CD016CodeMap);
        v_NewName4 := GetNewName(crSO051.STCodeB,crSO051.STNameB,CD016NameMap);


        UPDATE SO051 set UCCode=v_NewCode1,UCName=v_NewName1,
                         UCCodeB=v_NewCode2,UCNameB=v_NewName2,
                         STCode=v_NewCode3,STName=v_NewName3,
                         STCodeB=v_NewCode4,STNameB=v_NewName4
                     WHERE Rowid=crSO051.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO051有誤  UCCode: '||crSO051.UCCode||' UCCodeB:'||crSO051.UCCodeB||'  STCode:'||crSO051.STCode||'  STCodeB:'||crSO051.STCodeB,'SO051');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO051','資料檔SO051代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO055資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO055';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO055 in ccSO055 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO055.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO055.MediaCode,crSO055.MediaName,CD009NameMap);

        v_NewCode2 := GetNewCode(crSO055.MediaCodeB,CD009CodeMap);
        v_NewName2 := GetNewName(crSO055.MediaCodeB,crSO055.MediaNameB,CD009NameMap);


        UPDATE SO055 set MediaCode=v_NewCode1,MediaName=v_NewName1,
                         MediaCodeB=v_NewCode2,MediaNameB=v_NewName2
                     WHERE Rowid=crSO055.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO055有誤  MediaCode: '||crSO055.MediaCode||'  MediaCodeB:'||crSO055.MediaCodeB,'SO055');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO055','資料檔SO055代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO067資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO067';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO067 in ccSO067 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO067.UCCode,CD013CodeMap);
        v_NewName1 := GetNewName(crSO067.UCCode,crSO067.UCName,CD013NameMap);

        v_NewCode2 := GetNewCode(crSO067.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO067.STCode,crSO067.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO067.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO067.CancelCode,crSO067.CancelName,CD051NameMap);


        UPDATE SO067 set UCCode=v_NewCode1,UCName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO067.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO067有誤  UCCode: '||crSO067.UCCode||'  STCode:'||crSO067.STCode||'  CancelCode:'||crSO067.CancelCode,'SO067');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO067','資料檔SO067代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



  --更新SO074資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO074';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO074 in ccSO074 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO074.STCode,CD016CodeMap);
        v_NewName1 := GetNewName(crSO074.STCode,crSO074.STName,CD016NameMap);

        v_NewCode2 := GetNewCode(crSO074.CancelCode,CD051CodeMap);
        v_NewName2 := GetNewName(crSO074.CancelCode,crSO074.CancelName,CD051NameMap);


        UPDATE SO074 set STCode=v_NewCode1,STName=v_NewName1,
                         CancelCode=v_NewCode2,CancelName=v_NewName2
                     WHERE Rowid=crSO074.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO074有誤  BillNo: '||crSO074.BillNo||'  Item:'||crSO074.Item,'SO074');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO074','資料檔SO074代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO074A資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO074A';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO074A in ccSO074A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO074A.STCode,CD016CodeMap);
        v_NewName1 := GetNewName(crSO074A.STCode,crSO074A.STName,CD016NameMap);

        v_NewCode2 := GetNewCode(crSO074A.CancelCode,CD051CodeMap);
        v_NewName2 := GetNewName(crSO074A.CancelCode,crSO074A.CancelName,CD051NameMap);


        UPDATE SO074A set STCode=v_NewCode1,STName=v_NewName1,
                         CancelCode=v_NewCode2,CancelName=v_NewName2
                     WHERE Rowid=crSO074A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO074A有誤  BillNo: '||crSO074A.BillNo||'  Item:'||crSO074A.Item,'SO074A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO074A','資料檔SO074A代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



  --更新SO077資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO077';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO077 in ccSO077 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO077.STCode,CD016CodeMap);
        v_NewName1 := GetNewName(crSO077.STCode,crSO077.STName,CD016NameMap);

        v_NewCode2 := GetNewCode(crSO077.CancelCode,CD051CodeMap);
        v_NewName2 := GetNewName(crSO077.CancelCode,crSO077.CancelName,CD051NameMap);


        UPDATE SO077 set STCode=v_NewCode1,STName=v_NewName1,
                         CancelCode=v_NewCode2,CancelName=v_NewName2
                     WHERE Rowid=crSO077.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO077有誤  BillNo: '||crSO077.BillNo||'  Item:'||crSO077.Item,'SO077');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO077','資料檔SO077代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO095資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO095';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO095 in ccSO095 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO095.BulletinCode,CD049CodeMap);
        v_NewName1 := GetNewName(crSO095.BulletinCode,crSO095.BulletinName,CD049NameMap);


        UPDATE SO095 set BulletinCode=v_NewCode1,BulletinName=v_NewName1
                     WHERE Rowid=crSO095.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO095有誤  BulletinCode: '||crSO095.BulletinCode,'SO095');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO095','資料檔SO095代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO096資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO096';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO096 in ccSO096 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO096.BulletinCode,CD049CodeMap);
        v_NewName1 := GetNewName(crSO096.BulletinCode,crSO096.BulletinName,CD049NameMap);


        UPDATE SO096 set BulletinCode=v_NewCode1,BulletinName=v_NewName1
                     WHERE Rowid=crSO096.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO096有誤  BulletinCode: '||crSO096.BulletinCode,'SO096');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO096','資料檔SO096代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO098資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO098';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO098 in ccSO098 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO098.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO098.MediaCode,crSO098.MediaName,CD009NameMap);

        v_NewCode2 := GetNewCode(crSO098.BulletinCode,CD049CodeMap);
        v_NewName2 := GetNewName(crSO098.BulletinCode,crSO098.BulletinName,CD049NameMap);


        UPDATE SO098 set MediaCode=v_NewCode1,MediaName=v_NewName1,
                         BulletinCode=v_NewCode2,BulletinName=v_NewName2
                     WHERE Rowid=crSO098.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO098有誤  MediaCode: '||crSO098.MediaCode||'  BulletinCode:'||crSO098.BulletinCode,'SO098');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO098','資料檔SO098代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO101資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO101';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO101 in ccSO101 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO101.DelCode,CD012CodeMap);
        v_NewName1 := GetNewName(crSO101.DelCode,crSO101.DelName,CD012NameMap);


        UPDATE SO101 set DelCode=v_NewCode1,DelName=v_NewName1
                     WHERE Rowid=crSO101.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO101有誤  DelCode: '||crSO101.DelCode,'SO101');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO101','資料檔SO101代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



  --更新SO105資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105 in ccSO105 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105.ServiceCode,CD008CodeMap);

        v_NewCode2 := GetNewCode(crSO105.ServDescCode,CD008ACodeMap);

        v_NewCode3 := GetNewCode(crSO105.MediaCode,CD009CodeMap);
        v_NewName3 := GetNewName(crSO105.MediaCode,crSO105.MediaName,CD009NameMap);

        v_NewCode4 := GetNewCode(crSO105.ReturnCode,CD015CodeMap);
        v_NewName4 := GetNewName(crSO105.ReturnCode,crSO105.ReturnName,CD015NameMap);

        v_NewCode5 := GetNewCode(crSO105.CATVReturnCode,CD015CodeMap);
        v_NewName5 := GetNewName(crSO105.CATVReturnCode,crSO105.CATVReturnName,CD015NameMap);

        v_NewCode6 := GetNewCode(crSO105.DVSReturnCode,CD015CodeMap);
        v_NewName6 := GetNewName(crSO105.DVSReturnCode,crSO105.DVSReturnName,CD015NameMap);

        v_NewCode7 := GetNewCode(crSO105.BulletinCode,CD049CodeMap);


        UPDATE SO105 set ServiceCode=v_NewCode1,
                         ServDescCode=v_NewCode2,
                         MediaCode=v_NewCode3,MediaName=v_NewName3,
                         ReturnCode=v_NewCode4,ReturnName=v_NewName4,
                         CATVReturnCode=v_NewCode5,CATVReturnName=v_NewName5,
                         DVSReturnCode=v_NewCode6,DVSReturnName=v_NewName6,
                         BulletinCode=v_NewCode7
                     WHERE Rowid=crSO105.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO105有誤  OrderNo: '||crSO105.OrderNo,'SO105');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105','資料檔SO105代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO105T資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105T';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105T in ccSO105T loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105T.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO105T.MediaCode,crSO105T.MediaName,CD009NameMap);

        v_NewCode2 := GetNewCode(crSO105T.ReturnCode,CD015CodeMap);
        v_NewName2 := GetNewName(crSO105T.ReturnCode,crSO105T.ReturnName,CD015NameMap);

        v_NewCode3 := GetNewCode(crSO105T.CATVReturnCode,CD015CodeMap);
        v_NewName3 := GetNewName(crSO105T.CATVReturnCode,crSO105T.CATVReturnName,CD015NameMap);

        v_NewCode4 := GetNewCode(crSO105T.CMReturnCode,CD015CodeMap);
        v_NewName4 := GetNewName(crSO105T.CMReturnCode,crSO105T.CMReturnName,CD015NameMap);

        v_NewCode5 := GetNewCode(crSO105T.DVSReturnCode,CD015CodeMap);
        v_NewName5 := GetNewName(crSO105T.DVSReturnCode,crSO105T.DVSReturnName,CD015NameMap);

        v_NewCode6 := GetNewCode(crSO105T.BulletinCode,CD049CodeMap);
        v_NewName6 := GetNewName(crSO105T.BulletinCode,crSO105T.BulletinName,CD049NameMap);


        UPDATE SO105T set MediaCode=v_NewCode1,MediaName=v_NewName1,
                          ReturnCode=v_NewCode2,ReturnName=v_NewName2,
                          CATVReturnCode=v_NewCode3,CATVReturnName=v_NewName3,
                          CMReturnCode=v_NewCode4,CMReturnName=v_NewName4,
                          DVSReturnCode=v_NewCode5,DVSReturnName=v_NewName5,
                          BulletinCode=v_NewCode6,BulletinName=v_NewName6
                     WHERE Rowid=crSO105T.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO105T有誤  OrderNo: '||crSO105T.OrderNo||'  CompCode'||crSO105T.CompCode,'SO105T');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105T','資料檔SO105T代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO106資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO106';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO106 in ccSO106 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO106.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO106.MediaCode,crSO106.MediaName,CD009NameMap);


        UPDATE SO106 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO106.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO106有誤  MediaCode: '||crSO106.MediaCode,'SO106');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO106','資料檔SO106代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO121資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO121';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO121 in ccSO121 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO121.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO121.MediaCode,crSO121.MediaName,CD009NameMap);


        UPDATE SO121 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO121.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO121有誤  MediaCode: '||crSO121.MediaCode,'SO121');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO121','資料檔SO121代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO122資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO122';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO122 in ccSO122 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO122.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO122.MediaCode,crSO122.MediaName,CD009NameMap);


        UPDATE SO122 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO122.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO122有誤  MediaCode: '||crSO122.MediaCode,'SO122');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO122','資料檔SO122代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO131資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO131';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO131 in ccSO131 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO131.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO131.MediaCode,crSO131.MediaName,CD009NameMap);


        UPDATE SO131 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO131.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO131有誤  MediaCode: '||crSO131.MediaCode,'SO131');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO131','資料檔SO131代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO134資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO134';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO134 in ccSO134 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO134.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO134.MediaCode,crSO134.MediaName,CD009NameMap);


        UPDATE SO134 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO134.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO134有誤  MediaCode: '||crSO134.MediaCode,'SO134');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO134','資料檔SO134代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --更新SO503資料
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO503';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO503 in ccSO503 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO503.MediaCode,CD009CodeMap);
        v_NewName1 := GetNewName(crSO503.MediaCode,crSO503.MediaName,CD009NameMap);


        UPDATE SO503 set MediaCode=v_NewCode1,MediaName=v_NewName1
                     WHERE Rowid=crSO503.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 資料檔SO503有誤  MediaCode: '||crSO503.MediaCode,'SO503');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO503','資料檔SO503代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

--1111

--*****************************************************************
--*******************   更  新  代  碼  檔   **********************
--*****************************************************************
/*
  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD006';

  IF v_TempCounts=0 THEN --轉代碼前先確定沒有轉過
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    --置換舊代碼(或刪除轉相同新代碼的舊代碼)
    FOR crCD006 in ccCD006 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD006.CodeNo,CD006CodeMap);
        v_NewName1 := GetNewName(crCD006.CodeNo,CD006NameMap);
        v_Mark := GetMark(crCD006.CodeNo,CD006Mark);

        IF v_Mark=-1 THEN--多個舊代碼要轉一個新代碼,只能留一個
          DELETE CD006 WHERE Rowid=crCD006.Rowid;
        ELSE
          UPDATE CD006 set CODENO=v_NewCode1,Description=v_NewName1
                 WHERE Rowid=crCD006.Rowid;

          v_OKCounts := v_OKCounts + 1;
        END IF;

      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE 代碼檔CD006有誤  CodeNo: '||crCD006.CodeNo,'CD006');

      END;
    END LOOP;


    --新增新代碼
    FOR crCD006_MAP in ccCD006_MAP loop
      v_TempCounts := 0;
      SELECT COUNT(*) INTO v_TempCounts FROM CD006 WHERE CodeNo=crCD006_MAP.NewCode;
      IF v_TempCounts=0 THEN --避免新增相同的代碼
        BEGIN
          INSERT INTO CD006(CodeNo,Description,StopFlag,ServiceType,PrintWONow)
            VALUES(crCD006_MAP.NewCode,crCD006_MAP.NewName,C_StopFlag,C_ServiceType,C_PrintWONow);

          v_OKCounts := v_OKCounts + 1;
        EXCEPTION
          WHEN others THEN
            v_ErrorCounts := v_ErrorCounts + 1;

            INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
              VALUES(C_UpdDataErrorType,'INSERT 代碼檔CD006有誤  NewCodeNo: '||crCD006_MAP.NewCode,'CD006');

        END;
      ELSE
        v_ErrorCounts := v_ErrorCounts + 1;

        INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
          VALUES(C_UpdDataErrorType,'INSERT 代碼檔CD006有重複  NewCodeNo: '||crCD006_MAP.NewCode,'CD006');

      END IF;
    END LOOP;


    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD006','代碼檔CD006代碼轉換完成',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
*/
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  v_TotalAllTiem := GetDiffTime(v_TotalStartTime,SysDate);
  DBMS_OUTPUT.PUT_LINE('執行完畢:' || v_TotalAllTiem);
exception
  WHEN OTHERS THEN
    v_ErrDesc := 'Error Code ='||SQLCODE||' '||'Error message='||SQLERRM;
    DBMS_OUTPUT.PUT_LINE(v_ErrDesc);

    INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
      VALUES(C_OtherErrorType,v_ErrDesc,'');
end;
/