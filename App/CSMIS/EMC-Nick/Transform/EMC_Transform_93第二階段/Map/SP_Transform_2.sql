CREATE OR REPLACE PROCEDURE SP_TRANSFORM_2
  AS
/*
    �F�˥N�X����
    �ɦW: SP_Transform_2.sql
    �ഫ�����CD019A~SO034
    
    By: Jackal   Date: 2004.01.30
    
  	2004.02.17 ��: 1.CD005,CD006,CD007,CD008C,CD011A,CD019CD004,CD022,CD022CD019��CD059
                     ���N�X��,�Y���ݭn�Q�ഫ���D���n���,�@�֧�ѷs�N�XLoad�覡�s�W.
                   2.�]CD041�N�X�ɨM�w����,�G�����SO020�礣�θm��
                   3.CD019A,CD019B,CD019SO017,CD022A,CD027A,CD027B,CD027C��CD027D���N�X�ɷ�@������ഫ
*/

--declare
  --�t�ΰѼ�(�`��)
  C_CodeMapErrorType number :=1;          --��Ӫ��~
  C_UpdDataErrorType number :=2;          --UPDATE����ɦ��~
  C_OtherErrorType   number :=9;          --��L���`���~
  C_StopFlag         number :=0;          --�O�_����,�w�]�_
  C_ServiceType      varchar2(3) := 'C';  --�~�̪A�Ⱥ���,�w�]��'C'
  C_PrintWONow       number :=0;          --�O�_�ߧY�C�L�u��,�w�]�_

  v_SQL1           varchar2(1000);
  v_HaveRBS27      number :=0;   --0��RBS27   1�S��
  v_ErrDesc        varchar2(1000);


  c39		char(1) := chr(39);
  v_OldCode         number;
  v_OldName         varchar2(50);
  v_NewCode         number;
  v_NewName         varchar2(50);
  v_Mark            number;
  v_MaxOldCode      number;
  v_OKCounts        number := 0;
  v_ErrorCounts     number := 0;
  v_StartTime       Date;
  v_TempCounts      number;
  v_AllTiem         varchar2(100);
  v_TotalStartTime  Date;
  v_TotalAllTiem    varchar2(100);

  v_Note1           varchar2(500);
  v_Note2           varchar2(500);
  v_AllNote         varchar2(750);
  V_CitemCodeStr    varchar2(750);

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

  v_NewName1        varchar2(50);
  v_NewName2        varchar2(50);
  v_NewName3        varchar2(50);
  v_NewName4        varchar2(50);
  v_NewName5        varchar2(50);
  v_NewName6        varchar2(50);
  v_NewName7        varchar2(50);
  v_NewName8        varchar2(50);
  v_NewName9        varchar2(50);
  v_NewName10       varchar2(50);


--������ഫ��
/*
  cursor ccCD005 is
    select Rowid,CodeNo,CitemCode1,CitemName1,CitemCode2,CitemName2,
           CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5,
           FacilCode1,FacilName1,FacilCode2,FacilName2,FacilCode3,FacilName3 from CD005;

  cursor ccCD006 is
    select Rowid,CodeNo,CitemCode1,CitemName1,CitemCode2,CitemName2,
           CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5,
           FacilCode1,FacilName1,FacilCode2,FacilName2,FacilCode3,FacilName3 from CD006;

  cursor ccCD007 is
    select Rowid,CodeNo,CitemCode1,CitemName1,CitemCode2,CitemName2,
           CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5,
           FacilCode1,FacilName1,FacilCode2,FacilName2,FacilCode3,FacilName3 from CD007;

  cursor ccCD022 is
    select Rowid,CodeNo,DefCitemCode,DefCitemName from CD022;
    
  cursor ccCD059 is
    select Rowid,CodeNo,CompCode,DefCitemCode,DefCitemName from CD059;
*/
--0000
  cursor ccCD019A is
    select Rowid,CodeNo,CitemCode from CD019A;
    
  cursor ccCD019B is
    select Rowid,CitemCode,CitemName from CD019B;
    
  cursor ccCD019SO017 is
    select Rowid,CitemCode from CD019SO017;
    
  cursor ccCD022A is
    select Rowid,CitemCode,CitemName,FaciCode from CD022A;
    
  cursor ccCD024 is
    select Rowid,CodeNo,CitemCode,CitemName from CD024;
    
  cursor ccCD027A is
    select Rowid,PackageNo,CitemCode,Description from CD027A;
    
  cursor ccCD027B is
    select Rowid,PackageNo,CodeNo,Description from CD027B;
    
  cursor ccCD027C is
    select Rowid,PackageNo,CodeNo,Description from CD027C;
    
  cursor ccCD027D is
    select Rowid,PackageNo,CitemCode,Description from CD027D;

  cursor ccCD048A is
    select Rowid,CitemCodeStr from CD048A where CitemCodeStr is not null;
    
  cursor ccCD057 is
    select Rowid,CodeNo,ReturnCode,ReturnName from CD057;

  cursor ccCD057A is
    select Rowid,CitemCode1,CitemName1,CitemCode2,CitemName2,
           CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5,
           FacilCode1,FacilName1,FacilCode2,FacilName2,FacilCode3,FacilName3,
           InstCode,InstName from CD057A;


    
  cursor ccCD068 is
    select Rowid,CitemCodeStr from CD068 where CitemCodeStr is not null;

  cursor ccSO001 is
    select Rowid,CustId,CompCode,ClassCode1,ClassName1,ClassCode2,ClassName2,
           ClassCode3,ClassName3,OrgCode,OrgName,PwCode,PwName from SO001;

  cursor ccSO002 is
    select Rowid,CustId,CompCode,PRCode,PRName,StopCode,StopName,
           BankCode,BankName,ServiceType,MediaCode,MediaName,DelCode,DelName from SO002;

  cursor ccSO002A is
    select Rowid,CustId,AccountNo,CompCode,BankCode,BankName from SO002A;

  cursor ccSO003 is
    select Rowid,CustId,CompCode,SeqNo,BankCode,BankName,
           CitemCode,CitemName from SO003;

  cursor ccSO004 is
    select Rowid,SEQNo,FaciCode,FaciName,MediaCode,MediaName from SO004;

  cursor ccSO005 is
    select Rowid,CitemCode from SO005;

  cursor ccSO006 is
    select Rowid,SEQNo,ServiceCode,ServiceName,
           ServDescCode,ServDescName,BulletinCode,BulletinName,
           MediaCode,MediaName from SO006;

  cursor ccSO007 is
    select Rowid,SNo,InstCode,InstName,BulletinCode,BulletinName,
           MediaCode,MediaName,ReturnCode,ReturnName,
           SatiCode,SatiName from SO007;

  cursor ccSO008 is
    select Rowid,SNo,ServiceCode,ServiceName,MFCode1,MFName1,MFCode2,MFName2,
           ReturnCode,ReturnName,SatiCode,SatiName from SO008;

  cursor ccSO009 is
    select Rowid,SNo,PRCode,PRName,ReasonCode,ReasonName,ReturnCode,ReturnName,
           SatiCode,SatiName from SO009;

  cursor ccSO013 is
    select Rowid,IntroId,CompCode,MediaCode,MediaName from SO013;

  cursor ccSO014 is
    select Rowid,CompCode,AddrNo,PipelineCode,PipelineName,BTCode,BTName from SO014;

  cursor ccSO016 is
    select Rowid,StrtCode,Ord,BTCode,BTName from SO016;

  cursor ccSO017 is
    select Rowid,MduId,CompCode,PipelineCode,PipelineName from SO017;

  cursor ccSO019 is
    select Rowid,PowerNo,CompCode,ClassCode,ClassName,FaciCode1,
           FaciName1,FaciCode2,FaciName2 from SO019;

  cursor ccSO022 is
    select Rowid,Sno,MFCode,MFName,MFCode2,MFName2 from SO022;

  cursor ccSO032 is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO032;

  cursor ccSO033 is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO033;

  cursor ccSO033DEBT is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO033DEBT;

  cursor ccSO034 is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO034;

--0000

--�N�X���ഫ��
/*
  cursor ccCD005 is
    select Rowid,CodeNo from CD005 ORDER BY CodeNo;
  cursor ccCD005_MAP is
    select NewCode,NewName from CD005_MAP WHERE OldCode IS NULL ORDER BY NewCode;

*/


--�ŧi�U�N�X�}�C
  TYPE NumberAry IS TABLE OF number INDEX BY BINARY_INTEGER;
  TYPE Varchar2Ary IS TABLE OF varchar2(50) INDEX BY BINARY_INTEGER;

  CD004CodeMap NumberAry;
  CD004Mark NumberAry;
  CD004NameMap Varchar2Ary;

  CD005CodeMap NumberAry;
  CD005Mark NumberAry;
  CD005NameMap Varchar2Ary;

  CD006CodeMap NumberAry;
  CD006Mark NumberAry;
  CD006NameMap Varchar2Ary;

  CD007CodeMap NumberAry;
  CD007Mark NumberAry;
  CD007NameMap Varchar2Ary;

  CD008CodeMap NumberAry;
  CD008Mark NumberAry;
  CD008NameMap Varchar2Ary;

  CD008ACodeMap NumberAry;
  CD008AMark NumberAry;
  CD008ANameMap Varchar2Ary;

  CD009CodeMap NumberAry;
  CD009Mark NumberAry;
  CD009NameMap Varchar2Ary;

  CD011CodeMap NumberAry;
  CD011Mark NumberAry;
  CD011NameMap Varchar2Ary;

  CD011BCodeMap NumberAry;
  CD011BMark NumberAry;
  CD011BNameMap Varchar2Ary;

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

  CD018CodeMap NumberAry;
  CD018Mark NumberAry;
  CD018NameMap Varchar2Ary;

  CD019CodeMap NumberAry;
  CD019Mark NumberAry;
  CD019NameMap Varchar2Ary;

  CD020CodeMap NumberAry;
  CD020Mark NumberAry;
  CD020NameMap Varchar2Ary;

  CD021CodeMap NumberAry;
  CD021Mark NumberAry;
  CD021NameMap Varchar2Ary;

  CD022CodeMap NumberAry;
  CD022Mark NumberAry;
  CD022NameMap Varchar2Ary;

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

  CD055CodeMap NumberAry;
  CD055Mark NumberAry;
  CD055NameMap Varchar2Ary;

  CD059CodeMap NumberAry;
  CD059Mark NumberAry;
  CD059NameMap Varchar2Ary;

  TYPE CurTyp IS REF CURSOR;  --�ۭqcursor���A, �ѷs��dynamic SQL
  v_DyCursor1 CurTyp;


--���ǤJ�S��Index���B�z
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


  --���ǤJ�S��Index���B�z
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
          RETURN p_OldName;--�Y�S�����������,�^���¦W��
        ELSE
          RETURN L_Result;
        END IF;
      EXCEPTION
        WHEN OTHERS then
          RETURN p_OldName;
      END;
    END IF;
  END;


  --���ǤJ�S��Index���B�z
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


  --�p���ɶ��t
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
    RETURN '�X�p����ɶ�: ' || L_HH1 || '��' || L_MM1 || '��' || L_SS1 || '��';
  END;


  --SetSegment
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


  --
  FUNCTION GetNote(p_OldCode Number,p_OldName varchar2,p_ColumnName varchar2)
  RETURN varchar2
  AS
  BEGIN
    IF p_OldCode IS NOT NULL THEN
      RETURN '{'||p_ColumnName||'='||p_OldName||'('||p_OldCode||')}';
    ELSE
      RETURN '';
    END IF;
  END;


--�ഫCitemCodeStr
  FUNCTION ChangeCitemCodeStr(p_CitemCodeStr varchar2,p_CodeMapArray IN NumberAry)
  RETURN varchar2
  AS
    L_OldCitemCode varchar2(10);
    L_OldCitemCodeStr varchar2(1000);
    L_TempOldCitemCodeStr varchar2(1000);
    
    L_NewCitemCode varchar2(10);
    L_NewCitemCodeStr varchar2(1000);
    L_Index Number;
    L_MarkIndex Number;
    L_HaveMark char(1):='Y';
  BEGIN

    L_OldCitemCodeStr := p_CitemCodeStr;
    L_MarkIndex := instr(L_OldCitemCodeStr, '''');
    IF L_MarkIndex > 0 THEN
      L_HaveMark := 'Y';
    ELSE
      L_HaveMark := 'N';
    END IF;
    
    L_OldCitemCodeStr := replace(L_OldCitemCodeStr,'''','');

    IF L_OldCitemCodeStr IS NULL THEN
      RETURN NULL;
    ELSIF L_OldCitemCodeStr = '' THEN
      RETURN '';
    ELSE
      L_NewCitemCodeStr := '';
          
      while L_OldCitemCodeStr is not null loop
        L_Index := instr(L_OldCitemCodeStr, ',');
        if L_Index > 0 then
    	    begin
    	
            L_OldCitemCode := ltrim(rtrim(substr(L_OldCitemCodeStr, 1, L_Index-1)));
            
            BEGIN
              L_NewCitemCode := TO_CHAR(p_CodeMapArray(TO_NUMBER(L_OldCitemCode)));

              IF (L_NewCitemCodeStr = '') OR (L_NewCitemCodeStr IS NULL) THEN
                IF L_HaveMark = 'Y' THEN
                  L_NewCitemCodeStr := '''' || L_NewCitemCode || '''';
                ELSE
                  L_NewCitemCodeStr := L_NewCitemCode;
                END IF;
              ELSE
                IF L_HaveMark = 'Y' THEN
                  L_NewCitemCodeStr := L_NewCitemCodeStr || ',''' || L_NewCitemCode || '''';
                ELSE
                  L_NewCitemCodeStr := L_NewCitemCodeStr || ',' || L_NewCitemCode;
                END IF;
              END IF;
              
              L_OldCitemCodeStr := substr(L_OldCitemCodeStr, L_Index+1);
            EXCEPTION
              WHEN OTHERS then
                L_NewCitemCodeStr := p_CitemCodeStr;
                L_OldCitemCodeStr := null;
            END;
  	
          exception
  	      when others then
  	        L_OldCitemCodeStr := null;
  	      end;
        else
          BEGIN
            L_NewCitemCode := TO_CHAR(p_CodeMapArray(TO_NUMBER(L_OldCitemCodeStr)));

            IF (L_NewCitemCodeStr = '') OR (L_NewCitemCodeStr IS NULL) THEN
              IF L_HaveMark = 'Y' THEN
                L_NewCitemCodeStr := '''' || L_NewCitemCode || '''';
              ELSE
                L_NewCitemCodeStr := L_NewCitemCode;
              END IF;
            ELSE
              IF L_HaveMark = 'Y' THEN
                L_NewCitemCodeStr := L_NewCitemCodeStr || ',''' || L_NewCitemCode || '''';
              ELSE
                L_NewCitemCodeStr := L_NewCitemCodeStr || ',' || L_NewCitemCode;
              END IF;
            END IF;
            
          EXCEPTION
            WHEN OTHERS then
              L_NewCitemCodeStr := p_CitemCodeStr;
          END;

          L_OldCitemCodeStr := null;
        end if;
      end loop;
      
      RETURN  L_NewCitemCodeStr;

    END IF;
  END;

BEGIN
  v_TotalStartTime := SysDate;
  DBMS_OUTPUT.PUT_LINE('�}�l����:' || to_char(v_TotalStartTime,'YYYY/MM/DD HH24:MI:SS'));

  --�P�w�OORACLE���LRBS27
  v_SQL1 := 'select COUNT(segment_name) from dba_rollback_segs WHERE segment_name=''RBS27''';
  BEGIN
    OPEN v_DyCursor1 FOR v_SQL1;
  EXCEPTION
    WHEN others THEN
      DBMS_OUTPUT.PUT_LINE('dba_rollback_segs �S�}���v��');
      RETURN;
  END;


  LOOP
    FETCH v_DyCursor1 INTO v_HaveRBS27;
    EXIT WHEN v_DyCursor1%NOTFOUND;

    IF v_HaveRBS27=1 THEN--8i�N"RBS27" ONLINE
      BEGIN
        DBMS_UTILITY.EXEC_DDL_STATEMENT('ALTER ROLLBACK SEGMENT "RBS27" ONLINE');
      EXCEPTION
        WHEN others THEN
          NULL;
      END;
    END IF;
  END LOOP;

  CLOSE v_DyCursor1;




--************************************************************************
  -- ��M CD004 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD004;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD004_MAP WHERE OldCode=I;

        --�N��Ӫ�CD004_MAP�ন�}�C
        CD004CodeMap(I) := v_NewCode;
        CD004NameMap(I) := v_NewName;
        CD004Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD004_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD004_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD004CodeMap(I) := I;
          CD004NameMap(I) := null;
          CD004Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD005 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD005;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD005_MAP WHERE OldCode=I;

        --�N��Ӫ�CD005_MAP�ন�}�C
        CD005CodeMap(I) := v_NewCode;
        CD005NameMap(I) := v_NewName;
        CD005Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD005_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD005_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD005CodeMap(I) := I;
          CD005NameMap(I) := null;
          CD005Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD006 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD006;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD006_MAP WHERE OldCode=I;

        --�N��Ӫ�CD006_MAP�ন�}�C
        CD006CodeMap(I) := v_NewCode;
        CD006NameMap(I) := v_NewName;
        CD006Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD006_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD006_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD006CodeMap(I) := I;
          CD006NameMap(I) := null;
          CD006Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD007 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD007;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD007_MAP WHERE OldCode=I;

        --�N��Ӫ�CD007_MAP�ন�}�C
        CD007CodeMap(I) := v_NewCode;
        CD007NameMap(I) := v_NewName;
        CD007Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD007_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD007_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD007CodeMap(I) := I;
          CD007NameMap(I) := null;
          CD007Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD008 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD008;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD008_MAP WHERE OldCode=I;

        --�N��Ӫ�CD008_MAP�ন�}�C
        CD008CodeMap(I) := v_NewCode;
        CD008NameMap(I) := v_NewName;
        CD008Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD008_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD008_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD008CodeMap(I) := I;
          CD008NameMap(I) := null;
          CD008Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD008A �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD008A;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD008A_MAP WHERE OldCode=I;

        --�N��Ӫ�CD008A_MAP�ন�}�C
        CD008ACodeMap(I) := v_NewCode;
        CD008ANameMap(I) := v_NewName;
        CD008AMark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD008A_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD008A_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD008ACodeMap(I) := I;
          CD008ANameMap(I) := null;
          CD008AMark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- ��M CD009 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD009;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD009_MAP WHERE OldCode=I;

        --�N��Ӫ�CD009_MAP�ন�}�C
        CD009CodeMap(I) := v_NewCode;
        CD009NameMap(I) := v_NewName;
        CD009Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD009_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD009_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD009CodeMap(I) := I;
          CD009NameMap(I) := null;
          CD009Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD011 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD011;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD011_MAP WHERE OldCode=I;

        --�N��Ӫ�CD011_MAP�ন�}�C
        CD011CodeMap(I) := v_NewCode;
        CD011NameMap(I) := v_NewName;
        CD011Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD011_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD011_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD011CodeMap(I) := I;
          CD011NameMap(I) := null;
          CD011Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD011B �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD011B;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD011B_MAP WHERE OldCode=I;

        --�N��Ӫ�CD011B_MAP�ন�}�C
        CD011BCodeMap(I) := v_NewCode;
        CD011BNameMap(I) := v_NewName;
        CD011BMark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD011B_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD011B_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD011BCodeMap(I) := I;
          CD011BNameMap(I) := null;
          CD011BMark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- ��M CD012 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD012;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD012_MAP WHERE OldCode=I;

        --�N��Ӫ�CD012_MAP�ন�}�C
        CD012CodeMap(I) := v_NewCode;
        CD012NameMap(I) := v_NewName;
        CD012Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD012_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD012_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD012CodeMap(I) := I;
          CD012NameMap(I) := null;
          CD012Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD013 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD013;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD013_MAP WHERE OldCode=I;

        --�N��Ӫ�CD013_MAP�ন�}�C
        CD013CodeMap(I) := v_NewCode;
        CD013NameMap(I) := v_NewName;
        CD013Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD013_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD013_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD013CodeMap(I) := I;
          CD013NameMap(I) := null;
          CD013Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD014 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD014;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD014_MAP WHERE OldCode=I;

        --�N��Ӫ�CD014_MAP�ন�}�C
        CD014CodeMap(I) := v_NewCode;
        CD014NameMap(I) := v_NewName;
        CD014Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD014_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD014_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD014CodeMap(I) := I;
          CD014NameMap(I) := null;
          CD014Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD015 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD015;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD015_MAP WHERE OldCode=I;

        --�N��Ӫ�CD015_MAP�ন�}�C
        CD015CodeMap(I) := v_NewCode;
        CD015NameMap(I) := v_NewName;
        CD015Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD015_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD015_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD015CodeMap(I) := I;
          CD015NameMap(I) := null;
          CD015Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD016 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD016;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD016_MAP WHERE OldCode=I;

        --�N��Ӫ�CD016_MAP�ন�}�C
        CD016CodeMap(I) := v_NewCode;
        CD016NameMap(I) := v_NewName;
        CD016Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD016_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD016_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD016CodeMap(I) := I;
          CD016NameMap(I) := null;
          CD016Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD018 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD018;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD018_MAP WHERE OldCode=I;

        --�N��Ӫ�CD018_MAP�ন�}�C
        CD018CodeMap(I) := v_NewCode;
        CD018NameMap(I) := v_NewName;
        CD018Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD018_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD018_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD018CodeMap(I) := I;
          CD018NameMap(I) := null;
          CD018Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD019 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD019;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD019_MAP WHERE OldCode=I;

        --�N��Ӫ�CD019_MAP�ন�}�C
        CD019CodeMap(I) := v_NewCode;
        CD019NameMap(I) := v_NewName;
        CD019Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD019_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD019_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD019CodeMap(I) := I;
          CD019NameMap(I) := null;
          CD019Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD020 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD020;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD020_MAP WHERE OldCode=I;

        --�N��Ӫ�CD020_MAP�ন�}�C
        CD020CodeMap(I) := v_NewCode;
        CD020NameMap(I) := v_NewName;
        CD020Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD020_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD020_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD020CodeMap(I) := I;
          CD020NameMap(I) := null;
          CD020Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD021 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD021;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD021_MAP WHERE OldCode=I;

        --�N��Ӫ�CD021_MAP�ন�}�C
        CD021CodeMap(I) := v_NewCode;
        CD021NameMap(I) := v_NewName;
        CD021Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD021_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD021_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD021CodeMap(I) := I;
          CD021NameMap(I) := null;
          CD021Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD022 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD022;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD022_MAP WHERE OldCode=I;

        --�N��Ӫ�CD022_MAP�ন�}�C
        CD022CodeMap(I) := v_NewCode;
        CD022NameMap(I) := v_NewName;
        CD022Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD022_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD022_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD022CodeMap(I) := I;
          CD022NameMap(I) := null;
          CD022Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD025 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD025;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD025_MAP WHERE OldCode=I;

        --�N��Ӫ�CD025_MAP�ন�}�C
        CD025CodeMap(I) := v_NewCode;
        CD025NameMap(I) := v_NewName;
        CD025Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD025_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD025_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD025CodeMap(I) := I;
          CD025NameMap(I) := null;
          CD025Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD026 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD026;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD026_MAP WHERE OldCode=I;

        --�N��Ӫ�CD026_MAP�ন�}�C
        CD026CodeMap(I) := v_NewCode;
        CD026NameMap(I) := v_NewName;
        CD026Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD026_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD026_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD026CodeMap(I) := I;
          CD026NameMap(I) := null;
          CD026Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;



  -- ��M CD049 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD049;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD049_MAP WHERE OldCode=I;

        --�N��Ӫ�CD049_MAP�ন�}�C
        CD049CodeMap(I) := v_NewCode;
        CD049NameMap(I) := v_NewName;
        CD049Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD049_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD049_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD049CodeMap(I) := I;
          CD049NameMap(I) := null;
          CD049Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD051 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD051;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD051_MAP WHERE OldCode=I;

        --�N��Ӫ�CD051_MAP�ন�}�C
        CD051CodeMap(I) := v_NewCode;
        CD051NameMap(I) := v_NewName;
        CD051Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD051_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD051_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD051CodeMap(I) := I;
          CD051NameMap(I) := null;
          CD051Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD055 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD055;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD055_MAP WHERE OldCode=I;

        --�N��Ӫ�CD055_MAP�ন�}�C
        CD055CodeMap(I) := v_NewCode;
        CD055NameMap(I) := v_NewName;
        CD055Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD055_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD055_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD055CodeMap(I) := I;
          CD055NameMap(I) := null;
          CD055Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;


  -- ��M CD059 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD059;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD059_MAP WHERE OldCode=I;

        --�N��Ӫ�CD059_MAP�ন�}�C
        CD059CodeMap(I) := v_NewCode;
        CD059NameMap(I) := v_NewName;
        CD059Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD059_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD059_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD059CodeMap(I) := I;
          CD059NameMap(I) := null;
          CD059Mark(I) := 0;
          NULL;
      END;
    END LOOP;
  END IF;
--*****************************************************************
--*******************   ��  �s  ��  ��  ��   **********************
--*****************************************************************
/*
  --��sCD005���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD005';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD005 in ccCD005 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD005.CitemCode1,CD019CodeMap);
        v_NewName1 := GetNewName(crCD005.CitemCode1,crCD005.CitemName1,CD019NameMap);

        v_NewCode2 := GetNewCode(crCD005.CitemCode2,CD019CodeMap);
        v_NewName2 := GetNewName(crCD005.CitemCode2,crCD005.CitemName2,CD019NameMap);

        v_NewCode3 := GetNewCode(crCD005.CitemCode3,CD019CodeMap);
        v_NewName3 := GetNewName(crCD005.CitemCode3,crCD005.CitemName3,CD019NameMap);

        v_NewCode4 := GetNewCode(crCD005.CitemCode4,CD019CodeMap);
        v_NewName4 := GetNewName(crCD005.CitemCode4,crCD005.CitemName4,CD019NameMap);

        v_NewCode5 := GetNewCode(crCD005.CitemCode5,CD019CodeMap);
        v_NewName5 := GetNewName(crCD005.CitemCode5,crCD005.CitemName5,CD019NameMap);

        v_NewCode6 := GetNewCode(crCD005.FacilCode1,CD022CodeMap);
        v_NewName6 := GetNewName(crCD005.FacilCode1,crCD005.FacilName1,CD022NameMap);

        v_NewCode7 := GetNewCode(crCD005.FacilCode2,CD022CodeMap);
        v_NewName7 := GetNewName(crCD005.FacilCode2,crCD005.FacilName2,CD022NameMap);

        v_NewCode8 := GetNewCode(crCD005.FacilCode3,CD022CodeMap);
        v_NewName8 := GetNewName(crCD005.FacilCode3,crCD005.FacilName3,CD022NameMap);


        UPDATE CD005 set CitemCode1=v_NewCode1,CitemName1=v_NewName1,
                         CitemCode2=v_NewCode2,CitemName2=v_NewName2,
                         CitemCode3=v_NewCode3,CitemName3=v_NewName3,
                  			 CitemCode4=v_NewCode4,CitemName4=v_NewName4,
                         CitemCode5=v_NewCode5,CitemName5=v_NewName5,
                         FacilCode1=v_NewCode6,FacilName1=v_NewName6,
                         FacilCode2=v_NewCode7,FacilName2=v_NewName7,
                         FacilCode3=v_NewCode8,FacilName3=v_NewName8
                     WHERE Rowid=crCD005.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD005���~  CodeNo: '||crCD005.CodeNo,'CD005');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD005','�����CD005�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD006���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD006';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD006 in ccCD006 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD006.CitemCode1,CD019CodeMap);
        v_NewName1 := GetNewName(crCD006.CitemCode1,crCD006.CitemName1,CD019NameMap);

        v_NewCode2 := GetNewCode(crCD006.CitemCode2,CD019CodeMap);
        v_NewName2 := GetNewName(crCD006.CitemCode2,crCD006.CitemName2,CD019NameMap);

        v_NewCode3 := GetNewCode(crCD006.CitemCode3,CD019CodeMap);
        v_NewName3 := GetNewName(crCD006.CitemCode3,crCD006.CitemName3,CD019NameMap);

        v_NewCode4 := GetNewCode(crCD006.CitemCode4,CD019CodeMap);
        v_NewName4 := GetNewName(crCD006.CitemCode4,crCD006.CitemName4,CD019NameMap);

        v_NewCode5 := GetNewCode(crCD006.CitemCode5,CD019CodeMap);
        v_NewName5 := GetNewName(crCD006.CitemCode5,crCD006.CitemName5,CD019NameMap);

        v_NewCode6 := GetNewCode(crCD006.FacilCode1,CD022CodeMap);
        v_NewName6 := GetNewName(crCD006.FacilCode1,crCD006.FacilName1,CD022NameMap);

        v_NewCode7 := GetNewCode(crCD006.FacilCode2,CD022CodeMap);
        v_NewName7 := GetNewName(crCD006.FacilCode2,crCD006.FacilName2,CD022NameMap);

        v_NewCode8 := GetNewCode(crCD006.FacilCode3,CD022CodeMap);
        v_NewName8 := GetNewName(crCD006.FacilCode3,crCD006.FacilName3,CD022NameMap);


        UPDATE CD006 set CitemCode1=v_NewCode1,CitemName1=v_NewName1,
                         CitemCode2=v_NewCode2,CitemName2=v_NewName2,
                         CitemCode3=v_NewCode3,CitemName3=v_NewName3,
                      	 CitemCode4=v_NewCode4,CitemName4=v_NewName4,
                         CitemCode5=v_NewCode5,CitemName5=v_NewName5,
                         FacilCode1=v_NewCode6,FacilName1=v_NewName6,
                         FacilCode2=v_NewCode7,FacilName2=v_NewName7,
                         FacilCode3=v_NewCode8,FacilName3=v_NewName8
                     WHERE Rowid=crCD006.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD006���~  CodeNo: '||crCD006.CodeNo,'CD006');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD006','�����CD006�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD007���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD007';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD007 in ccCD007 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD007.CitemCode1,CD019CodeMap);
        v_NewName1 := GetNewName(crCD007.CitemCode1,crCD007.CitemName1,CD019NameMap);

        v_NewCode2 := GetNewCode(crCD007.CitemCode2,CD019CodeMap);
        v_NewName2 := GetNewName(crCD007.CitemCode2,crCD007.CitemName2,CD019NameMap);

        v_NewCode3 := GetNewCode(crCD007.CitemCode3,CD019CodeMap);
        v_NewName3 := GetNewName(crCD007.CitemCode3,crCD007.CitemName3,CD019NameMap);

        v_NewCode4 := GetNewCode(crCD007.CitemCode4,CD019CodeMap);
        v_NewName4 := GetNewName(crCD007.CitemCode4,crCD007.CitemName4,CD019NameMap);

        v_NewCode5 := GetNewCode(crCD007.CitemCode5,CD019CodeMap);
        v_NewName5 := GetNewName(crCD007.CitemCode5,crCD007.CitemName5,CD019NameMap);

        v_NewCode6 := GetNewCode(crCD007.FacilCode1,CD022CodeMap);
        v_NewName6 := GetNewName(crCD007.FacilCode1,crCD007.FacilName1,CD022NameMap);

        v_NewCode7 := GetNewCode(crCD007.FacilCode2,CD022CodeMap);
        v_NewName7 := GetNewName(crCD007.FacilCode2,crCD007.FacilName2,CD022NameMap);

        v_NewCode8 := GetNewCode(crCD007.FacilCode3,CD022CodeMap);
        v_NewName8 := GetNewName(crCD007.FacilCode3,crCD007.FacilName3,CD022NameMap);


        UPDATE CD007 set CitemCode1=v_NewCode1,CitemName1=v_NewName1,
                         CitemCode2=v_NewCode2,CitemName2=v_NewName2,
                         CitemCode3=v_NewCode3,CitemName3=v_NewName3,
                      	 CitemCode4=v_NewCode4,CitemName4=v_NewName4,
                         CitemCode5=v_NewCode5,CitemName5=v_NewName5,
                         FacilCode1=v_NewCode6,FacilName1=v_NewName6,
                         FacilCode2=v_NewCode7,FacilName2=v_NewName7,
                         FacilCode3=v_NewCode8,FacilName3=v_NewName8
                     WHERE Rowid=crCD007.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD007���~  CodeNo: '||crCD007.CodeNo,'CD007');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD007','�����CD007�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD022���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD022';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD022 in ccCD022 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD022.DefCitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD022.DefCitemCode,crCD022.DefCitemName,CD019NameMap);


        UPDATE CD022 set DefCitemCode=v_NewCode1,DefCitemName=v_NewName1
                     WHERE Rowid=crCD022.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD022���~  CodeNo: '||crCD022.CodeNo,'CD022');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD022','�����CD022�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD059���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD059';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD059 in ccCD059 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD059.DefCitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD059.DefCitemCode,crCD059.DefCitemName,CD019NameMap);


        UPDATE CD059 set DefCitemCode=v_NewCode1,DefCitemName=v_NewName1
                     WHERE Rowid=crCD059.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD059���~  CodeNo: '||crCD059.CodeNo||' CompCode: '||crCD059.CompCode,'CD059');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD059','�����CD059�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

*/

  --��sCD019A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD019A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD019A in ccCD019A loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD019A.CitemCode,CD019CodeMap);


        UPDATE CD019A set CitemCode=v_NewCode1
                     WHERE Rowid=crCD019A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD019A���~  CodeNo: '||crCD019A.CodeNo||' CitemCode: '||crCD019A.CitemCode,'CD019A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD019A','�����CD019A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD019B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD019B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD019B in ccCD019B loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD019B.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD019B.CitemCode,crCD019B.CitemName,CD019NameMap);


        UPDATE CD019B set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crCD019B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD019B���~  CitemCode: '||crCD019B.CitemCode,'CD019B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD019B','�����CD019B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD019SO017���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD019SO017';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD019SO017 in ccCD019SO017 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD019SO017.CitemCode,CD019CodeMap);


        UPDATE CD019SO017 set CitemCode=v_NewCode1
                     WHERE Rowid=crCD019SO017.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD019SO017���~  CitemCode: '||crCD019SO017.CitemCode,'CD019SO017');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD019SO017','�����CD019SO017�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD022A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD022A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD022A in ccCD022A loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD022A.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD022A.CitemCode,crCD022A.CitemName,CD019NameMap);

      	v_NewCode2 := GetNewCode(crCD022A.FaciCode,CD022CodeMap);


        UPDATE CD022A set CitemCode=v_NewCode1,CitemName=v_NewName1,
                          FaciCode=v_NewCode2
                     WHERE Rowid=crCD022A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD022A���~  CitemCode: '||crCD022A.CitemCode||' CitemCode: '||crCD022A.FaciCode,'CD022A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD022A','�����CD022A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD024���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD024';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD024 in ccCD024 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD024.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD024.CitemCode,crCD024.CitemName,CD019NameMap);


        UPDATE CD024 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crCD024.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD024���~  CodeNo: '||crCD024.CodeNo,'CD024');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD024','�����CD024�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD027A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD027A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD027A in ccCD027A loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD027A.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD027A.CitemCode,crCD027A.Description,CD019NameMap);


        UPDATE CD027A set CitemCode=v_NewCode1,Description=v_NewName1
                     WHERE Rowid=crCD027A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD027A���~  PackageNo: '||crCD027A.PackageNo||' CitemCode: '||crCD027A.CitemCode,'CD027A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD027A','�����CD027A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD027B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD027B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD027B in ccCD027B loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD027B.CodeNo,CD059CodeMap);
        v_NewName1 := GetNewName(crCD027B.CodeNo,crCD027B.Description,CD059NameMap);


        UPDATE CD027B set CodeNo=v_NewCode1,Description=v_NewName1
                     WHERE Rowid=crCD027B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD027B���~  PackageNo: '||crCD027B.PackageNo||' CodeNo: '||crCD027B.CodeNo,'CD027B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD027B','�����CD027B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD027C���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD027C';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD027C in ccCD027C loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD027C.CodeNo,CD022CodeMap);
        v_NewName1 := GetNewName(crCD027C.CodeNo,crCD027C.Description,CD022NameMap);


        UPDATE CD027C set CodeNo=v_NewCode1,Description=v_NewName1
                     WHERE Rowid=crCD027C.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD027C���~  PackageNo: '||crCD027C.PackageNo||' CodeNo: '||crCD027C.CodeNo,'CD027C');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD027C','�����CD027C�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD027D���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD027D';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD027D in ccCD027D loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD027D.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crCD027D.CitemCode,crCD027D.Description,CD019NameMap);


        UPDATE CD027D set CitemCode=v_NewCode1,Description=v_NewName1
                     WHERE Rowid=crCD027D.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD027D���~  PackageNo: '||crCD027D.PackageNo||' CitemCode: '||crCD027D.CitemCode,'CD027D');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD027D','�����CD027D�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

--9999
  --��sCD048A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD048A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD048A in ccCD048A loop
      BEGIN
        V_CitemCodeStr := ChangeCitemCodeStr(crCD048A.CitemCodeStr,CD019CodeMap);

        UPDATE CD048A set CitemCodeStr=V_CitemCodeStr
                     WHERE Rowid=crCD048A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD048A���~  CitemCodeStr: '||crCD048A.CitemCodeStr,'CD048A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD048A','�����CD048A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD057���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD057';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����CD057���~  CodeNo: '||crCD057.CodeNo,'CD057');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD057','�����CD057�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sCD057A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD057A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD057A in ccCD057A loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD057A.CitemCode1,CD019CodeMap);
        v_NewName1 := GetNewName(crCD057A.CitemCode1,crCD057A.CitemName1,CD019NameMap);

        v_NewCode2 := GetNewCode(crCD057A.CitemCode2,CD019CodeMap);
        v_NewName2 := GetNewName(crCD057A.CitemCode2,crCD057A.CitemName2,CD019NameMap);

        v_NewCode3 := GetNewCode(crCD057A.CitemCode3,CD019CodeMap);
        v_NewName3 := GetNewName(crCD057A.CitemCode3,crCD057A.CitemName3,CD019NameMap);

        v_NewCode4 := GetNewCode(crCD057A.CitemCode4,CD019CodeMap);
        v_NewName4 := GetNewName(crCD057A.CitemCode4,crCD057A.CitemName4,CD019NameMap);

        v_NewCode5 := GetNewCode(crCD057A.CitemCode5,CD019CodeMap);
        v_NewName5 := GetNewName(crCD057A.CitemCode5,crCD057A.CitemName5,CD019NameMap);

        v_NewCode6 := GetNewCode(crCD057A.FacilCode1,CD022CodeMap);
        v_NewName6 := GetNewName(crCD057A.FacilCode1,crCD057A.FacilName1,CD022NameMap);

        v_NewCode7 := GetNewCode(crCD057A.FacilCode2,CD022CodeMap);
        v_NewName7 := GetNewName(crCD057A.FacilCode2,crCD057A.FacilName2,CD022NameMap);

        v_NewCode8 := GetNewCode(crCD057A.FacilCode3,CD022CodeMap);
        v_NewName8 := GetNewName(crCD057A.FacilCode3,crCD057A.FacilName3,CD022NameMap);

        v_NewCode9 := GetNewCode(crCD057A.InstCode,CD005CodeMap);
        v_NewName9 := GetNewName(crCD057A.InstCode,crCD057A.InstName,CD005NameMap);


        UPDATE CD057A set CitemCode1=v_NewCode1,CitemName1=v_NewName1,
                         CitemCode2=v_NewCode2,CitemName2=v_NewName2,
                         CitemCode3=v_NewCode3,CitemName3=v_NewName3,
                      	 CitemCode4=v_NewCode4,CitemName4=v_NewName4,
                         CitemCode5=v_NewCode5,CitemName5=v_NewName5,
                         FacilCode1=v_NewCode6,FacilName1=v_NewName6,
                         FacilCode2=v_NewCode7,FacilName2=v_NewName7,
                         FacilCode3=v_NewCode8,FacilName3=v_NewName8,
                         InstCode=v_NewCode9,InstName=v_NewName9
                     WHERE Rowid=crCD057A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD057A���~  CitemCode1: '||crCD057A.CitemCode1||
                                                               'CitemCode2: '||crCD057A.CitemCode2||
                                                               'CitemCode3: '||crCD057A.CitemCode3||
                                                               'CitemCode4: '||crCD057A.CitemCode4||
                                                               'CitemCode5: '||crCD057A.CitemCode5||
                                                               'FacilCode1: '||crCD057A.FacilCode1||
                                                               'FacilCode2: '||crCD057A.FacilCode2||
                                                               'FacilCode3: '||crCD057A.FacilCode3||
                                                               'InstCode: '||crCD057A.InstCode
                                                               ,'CD057A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD057A','�����CD057A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%




  --��sCD068���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD068';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crCD068 in ccCD068 loop
      BEGIN
        V_CitemCodeStr := ChangeCitemCodeStr(crCD068.CitemCodeStr,CD019CodeMap);

        UPDATE CD068 set CitemCodeStr=V_CitemCodeStr
                     WHERE Rowid=crCD068.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����CD068���~  CitemCodeStr: '||crCD068.CitemCodeStr,'CD068');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD068','�����CD068�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO001���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO001';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO001 in ccSO001 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO001.ClassCode1,CD004CodeMap);
        v_NewName1 := GetNewName(crSO001.ClassCode1,crSO001.ClassName1,CD004NameMap);

        v_NewCode2 := GetNewCode(crSO001.ClassCode2,CD004CodeMap);
        v_NewName2 := GetNewName(crSO001.ClassCode2,crSO001.ClassName2,CD004NameMap);

        v_NewCode3 := GetNewCode(crSO001.ClassCode3,CD004CodeMap);
        v_NewName3 := GetNewName(crSO001.ClassCode3,crSO001.ClassName3,CD004NameMap);

        v_NewCode4 := GetNewCode(crSO001.OrgCode,CD025CodeMap);
        v_NewName4 := GetNewName(crSO001.OrgCode,crSO001.OrgName,CD025NameMap);

        v_NewCode5 := GetNewCode(crSO001.PwCode,CD020CodeMap);
        v_NewName5 := GetNewName(crSO001.PwCode,crSO001.PwName,CD020NameMap);


        UPDATE SO001 set ClassCode1=v_NewCode1,ClassName1=v_NewName1,
                         ClassCode2=v_NewCode2,ClassName2=v_NewName2,
                         ClassCode3=v_NewCode3,ClassName3=v_NewName3,
                         OrgCode=v_NewCode4,OrgName=v_NewName4,
                         PwCode=v_NewCode5,PwName=v_NewName5
                     WHERE Rowid=crSO001.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO001���~  CustId: '||crSO001.CustId||' CompCode:'||crSO001.CompCode,'SO001');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO001','�����SO001�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO002���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO002';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO002 in ccSO002 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO002.PRCode,CD014CodeMap);
        v_NewName1 := GetNewName(crSO002.PRCode,crSO002.PRName,CD014NameMap);

        v_NewCode2 := GetNewCode(crSO002.StopCode,CD014CodeMap);
        v_NewName2 := GetNewName(crSO002.StopCode,crSO002.StopName,CD014NameMap);

        v_NewCode3 := GetNewCode(crSO002.BankCode,CD018CodeMap);
        v_NewName3 := GetNewName(crSO002.BankCode,crSO002.BankName,CD018NameMap);

        v_NewCode4 := GetNewCode(crSO002.MediaCode,CD009CodeMap);
        v_NewName4 := GetNewName(crSO002.MediaCode,crSO002.MediaName,CD009NameMap);

        v_NewCode5 := GetNewCode(crSO002.DelCode,CD012CodeMap);
        v_NewName5 := GetNewName(crSO002.DelCode,crSO002.DelName,CD012NameMap);


        UPDATE SO002 set PRCode=v_NewCode1,PRName=v_NewName1,
                         StopCode=v_NewCode2,StopName=v_NewName2,
                         BankCode=v_NewCode3,BankName=v_NewName3,
                         MediaCode=v_NewCode4,MediaName=v_NewName4,
                         DelCode=v_NewCode5,DelName=v_NewName5
                     WHERE Rowid=crSO002.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO002���~  CustId: '||crSO002.CustId||' CompCode:'||crSO002.CompCode||'  ServiceType:'||crSO002.ServiceType,'SO002');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO002','�����SO002�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO002A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO002A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO002A in ccSO002A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO002A.BankCode,CD018CodeMap);
        v_NewName1 := GetNewName(crSO002A.BankCode,crSO002A.BankName,CD018NameMap);


        UPDATE SO002A set BankCode=v_NewCode1,BankName=v_NewName1
                     WHERE Rowid=crSO002A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO002A���~  CustId: '||crSO002A.CustId||' CompCode: '||crSO002A.CompCode||' AccountNo: '||crSO002A.AccountNo,'SO002A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO002A','�����SO002A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO003���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO003';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO003 in ccSO003 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO003.BankCode,CD018CodeMap);
        v_NewName1 := GetNewName(crSO003.BankCode,crSO003.BankName,CD018NameMap);

        v_NewCode2 := GetNewCode(crSO003.CitemCode,CD019CodeMap);
        v_NewName2 := GetNewName(crSO003.CitemCode,crSO003.CitemName,CD019NameMap);


        UPDATE SO003 set BankCode=v_NewCode1,BankName=v_NewName1,
                         CitemCode=v_NewCode2,CitemName=v_NewName2
                     WHERE Rowid=crSO003.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO003���~  CustId: '||crSO003.CustId||' CompCode: '||crSO003.CompCode||' CitemCode: '||crSO003.CitemCode||' SeqNo: '||crSO003.SeqNo,'SO003');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO003','�����SO003�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO004���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO004';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO004 in ccSO004 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO004.FaciCode,CD022CodeMap);
        v_NewName1 := GetNewName(crSO004.FaciCode,crSO004.FaciName,CD022NameMap);

        v_NewCode2 := GetNewCode(crSO004.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO004.MediaCode,crSO004.MediaName,CD009NameMap);



        UPDATE SO004 set FaciCode=v_NewCode1,FaciName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO004.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO004���~  SEQNo: '||crSO004.SEQNo,'SO004');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO004','�����SO004�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO005���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO005';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO005 in ccSO005 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO005.CitemCode,CD019CodeMap);

        UPDATE SO005 set CitemCode=v_NewCode1
                     WHERE Rowid=crSO005.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO005���~  CitemCode: '||crSO005.CitemCode,'SO005');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO005','�����SO005�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO006���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO006';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO006 in ccSO006 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO006.ServiceCode,CD008CodeMap);
        v_NewName1 := GetNewName(crSO006.ServiceCode,crSO006.ServiceName,CD008NameMap);

        v_NewCode2 := GetNewCode(crSO006.ServDescCode,CD008ACodeMap);
        v_NewName2 := GetNewName(crSO006.ServDescCode,crSO006.ServDescName,CD008ANameMap);

        v_NewCode3 := GetNewCode(crSO006.BulletinCode,CD049CodeMap);
        v_NewName3 := GetNewName(crSO006.BulletinCode,crSO006.BulletinName,CD049NameMap);

        v_NewCode4 := GetNewCode(crSO006.MediaCode,CD009CodeMap);
        v_NewName4 := GetNewName(crSO006.MediaCode,crSO006.MediaName,CD009NameMap);


        UPDATE SO006 set ServiceCode=v_NewCode1,ServiceName=v_NewName1,
                         ServDescCode=v_NewCode2,ServDescName=v_NewName2,
                         BulletinCode=v_NewCode3,BulletinName=v_NewName3,
                         MediaCode=v_NewCode4,MediaName=v_NewName4
                     WHERE Rowid=crSO006.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO006���~  SEQNo: '||crSO006.SEQNo,'SO006');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO006','�����SO006�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO007���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO007';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO007 in ccSO007 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO007.InstCode,CD005CodeMap);
        v_NewName1 := GetNewName(crSO007.InstCode,crSO007.InstName,CD005NameMap);

        v_NewCode2 := GetNewCode(crSO007.BulletinCode,CD049CodeMap);
        v_NewName2 := GetNewName(crSO007.BulletinCode,crSO007.BulletinName,CD049NameMap);

        v_NewCode3 := GetNewCode(crSO007.MediaCode,CD009CodeMap);
        v_NewName3 := GetNewName(crSO007.MediaCode,crSO007.MediaName,CD009NameMap);

        v_NewCode4 := GetNewCode(crSO007.ReturnCode,CD015CodeMap);
        v_NewName4 := GetNewName(crSO007.ReturnCode,crSO007.ReturnName,CD015NameMap);

        v_NewCode5 := GetNewCode(crSO007.SatiCode,CD026CodeMap);
        v_NewName5 := GetNewName(crSO007.SatiCode,crSO007.SatiName,CD026NameMap);


        UPDATE SO007 set InstCode=v_NewCode1,InstName=v_NewName1,
                         BulletinCode=v_NewCode2,BulletinName=v_NewName2,
                         MediaCode=v_NewCode3,MediaName=v_NewName3,
                         ReturnCode=v_NewCode4,ReturnName=v_NewName4,
                         SatiCode=v_NewCode5,SatiName=v_NewName5
                     WHERE Rowid=crSO007.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO007���~  SNo: '||crSO007.SNo,'SO007');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO007','�����SO007�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO008���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO008';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO008 in ccSO008 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO008.ServiceCode,CD006CodeMap);
        v_NewName1 := GetNewName(crSO008.ServiceCode,crSO008.ServiceName,CD006NameMap);

        v_NewCode2 := GetNewCode(crSO008.MFCode1,CD011CodeMap);
        v_NewName2 := GetNewName(crSO008.MFCode1,crSO008.MFName1,CD011NameMap);

        v_NewCode3 := GetNewCode(crSO008.MFCode2,CD011CodeMap);
        v_NewName3 := GetNewName(crSO008.MFCode2,crSO008.MFName2,CD011NameMap);

        v_NewCode4 := GetNewCode(crSO008.ReturnCode,CD015CodeMap);
        v_NewName4 := GetNewName(crSO008.ReturnCode,crSO008.ReturnName,CD015NameMap);

        v_NewCode5 := GetNewCode(crSO008.SatiCode,CD026CodeMap);
        v_NewName5 := GetNewName(crSO008.SatiCode,crSO008.SatiName,CD026NameMap);


        UPDATE SO008 set ServiceCode=v_NewCode1,ServiceName=v_NewName1,
                         MFCode1=v_NewCode2,MFName1=v_NewName2,
                         MFCode2=v_NewCode3,MFName2=v_NewName3,
                         ReturnCode=v_NewCode4,ReturnName=v_NewName4,
                         SatiCode=v_NewCode5,SatiName=v_NewName5
                     WHERE Rowid=crSO008.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO008���~  SNo: '||crSO008.SNo,'SO008');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO008','�����SO008�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO009���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO009';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO009 in ccSO009 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO009.PRCode,CD007CodeMap);
        v_NewName1 := GetNewName(crSO009.PRCode,crSO009.PRName,CD007NameMap);

        v_NewCode2 := GetNewCode(crSO009.ReasonCode,CD014CodeMap);
        v_NewName2 := GetNewName(crSO009.ReasonCode,crSO009.ReasonName,CD014NameMap);

        v_NewCode3 := GetNewCode(crSO009.ReturnCode,CD015CodeMap);
        v_NewName3 := GetNewName(crSO009.ReturnCode,crSO009.ReturnName,CD015NameMap);

        v_NewCode4 := GetNewCode(crSO009.SatiCode,CD026CodeMap);
        v_NewName4 := GetNewName(crSO009.SatiCode,crSO009.SatiName,CD026NameMap);


        UPDATE SO009 set PRCode=v_NewCode1,PRName=v_NewName1,
                         ReasonCode=v_NewCode2,ReasonName=v_NewName2,
                         ReturnCode=v_NewCode3,ReturnName=v_NewName3,
                         SatiCode=v_NewCode4,SatiName=v_NewName4
                     WHERE Rowid=crSO009.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO009���~  SNo: '||crSO009.SNo,'SO009');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO009','�����SO009�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO013���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO013';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO013���~  IntroId: '||crSO013.IntroId||'  CompCode:'||crSO013.CompCode,'SO013');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO013','�����SO013�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO014���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO014';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO014 in ccSO014 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO014.PipelineCode,CD055CodeMap);
        v_NewName1 := GetNewName(crSO014.PipelineCode,crSO014.PipelineName,CD055NameMap);

        v_NewCode2 := GetNewCode(crSO014.BTCode,CD021CodeMap);
        v_NewName2 := GetNewName(crSO014.BTCode,crSO014.BTName,CD021NameMap);


        UPDATE SO014 set PipelineCode=v_NewCode1,PipelineName=v_NewName1,
                         BTCode=v_NewCode2,BTName=v_NewName2
                     WHERE Rowid=crSO014.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO014���~  AddrNo: '||crSO014.AddrNo||'  CompCode:'||crSO014.CompCode,'SO014');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO014','�����SO014�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO016���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO016';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO016���~  StrtCode: '||crSO016.StrtCode||'  Ord:'||crSO016.Ord,'SO016');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO016','�����SO016�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO017���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO017';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO017 in ccSO017 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO017.PipelineCode,CD055CodeMap);
        v_NewName1 := GetNewName(crSO017.PipelineCode,crSO017.PipelineName,CD055NameMap);


        UPDATE SO017 set PipelineCode=v_NewCode1,PipelineName=v_NewName1
                     WHERE Rowid=crSO017.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO017���~  MduId: '||crSO017.MduId||'  CompCode:'||crSO017.CompCode,'SO017');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO017','�����SO017�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO019���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO019';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO019 in ccSO019 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO019.ClassCode,CD004CodeMap);
        v_NewName1 := GetNewName(crSO019.ClassCode,crSO019.ClassName,CD004NameMap);

        v_NewCode2 := GetNewCode(crSO019.FaciCode1,CD022CodeMap);
        v_NewName2 := GetNewName(crSO019.FaciCode1,crSO019.FaciName1,CD022NameMap);

        v_NewCode3 := GetNewCode(crSO019.FaciCode2,CD022CodeMap);
        v_NewName3 := GetNewName(crSO019.FaciCode2,crSO019.FaciName2,CD022NameMap);


        UPDATE SO019 set ClassCode=v_NewCode1,ClassName=v_NewName1,
                         FaciCode1=v_NewCode2,FaciName1=v_NewName2,
                         FaciCode2=v_NewCode3,FaciName2=v_NewName3
                     WHERE Rowid=crSO019.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO019���~  PowerNo: '||crSO019.PowerNo||'  CompCode:'||crSO019.CompCode,'SO019');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO019','�����SO019�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%




  --��sSO022���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO022';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO022 in ccSO022 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO022.MFCode,CD011CodeMap);
        v_NewName1 := GetNewName(crSO022.MFCode,crSO022.MFName,CD011NameMap);

        v_NewCode2 := GetNewCode(crSO022.MFCode2,CD011CodeMap);
        v_NewName2 := GetNewName(crSO022.MFCode2,crSO022.MFName2,CD011NameMap);


        UPDATE SO022 set MFCode=v_NewCode1,MFName=v_NewName1,
                         MFCode2=v_NewCode2,MFName2=v_NewName2
                     WHERE Rowid=crSO022.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO022���~  Sno: '||crSO022.Sno,'SO022');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO022','�����SO022�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO032���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO032';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO032 in ccSO032 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO032.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO032.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO032.BankCode,crSO032.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO032.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO032.CitemCode,crSO032.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO032.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO032.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO032.UCCode,crSO032.UCName,CD013NameMap);

        v_NewCode6 := GetNewCode(crSO032.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO032.STCode,crSO032.STName,CD016NameMap);

        v_NewCode7 := GetNewCode(crSO032.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO032.CancelCode,crSO032.CancelName,CD051NameMap);


        UPDATE SO032 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7
                     WHERE Rowid=crSO032.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO032���~  BillNo: '||crSO032.BillNo||'  Item:'||crSO032.Item,'SO032');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO032','�����SO032�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO033���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO033';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;
    v_Note1 := '';
    v_Note2 := '';
    v_AllNote := '';

    FOR crSO033 in ccSO033 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO033.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO033.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO033.BankCode,crSO033.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO033.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO033.CitemCode,crSO033.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO033.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO033.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO033.UCCode,crSO033.UCName,CD013NameMap);
        v_Note1 := GetNote(crSO033.UCCode,crSO033.UCName,'�ª���ú�O��]');

        v_NewCode6 := GetNewCode(crSO033.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO033.STCode,crSO033.STName,CD016NameMap);
        v_Note2 := GetNote(crSO033.STCode,crSO033.STName,'�ª��u����]');

        v_NewCode7 := GetNewCode(crSO033.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO033.CancelCode,crSO033.CancelName,CD051NameMap);

        v_AllNote := v_Note1||v_Note2;

        UPDATE SO033 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7,
                         Note=Note||v_AllNote
                     WHERE Rowid=crSO033.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO033���~  BillNo: '||crSO033.BillNo||'  Item:'||crSO033.Item,'SO033');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO033','�����SO033�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO033DEBT���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO033DEBT';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO033DEBT in ccSO033DEBT loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO033DEBT.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO033DEBT.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO033DEBT.BankCode,crSO033DEBT.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO033DEBT.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO033DEBT.CitemCode,crSO033DEBT.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO033DEBT.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO033DEBT.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO033DEBT.UCCode,crSO033DEBT.UCName,CD013NameMap);

        v_NewCode6 := GetNewCode(crSO033DEBT.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO033DEBT.STCode,crSO033DEBT.STName,CD016NameMap);

        v_NewCode7 := GetNewCode(crSO033DEBT.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO033DEBT.CancelCode,crSO033DEBT.CancelName,CD051NameMap);


        UPDATE SO033DEBT set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7
                     WHERE Rowid=crSO033DEBT.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO033DEBT���~  BillNo: '||crSO033DEBT.BillNo||'  Item:'||crSO033DEBT.Item,'SO033DEBT');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO033DEBT','�����SO033DEBT�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO034���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO034';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;
    v_Note1 := '';
    v_Note2 := '';
    v_AllNote := '';

    FOR crSO034 in ccSO034 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO034.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO034.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO034.BankCode,crSO034.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO034.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO034.CitemCode,crSO034.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO034.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO034.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO034.UCCode,crSO034.UCName,CD013NameMap);
        v_Note1 := GetNote(crSO034.UCCode,crSO034.UCName,'�ª���ú�O��]');

        v_NewCode6 := GetNewCode(crSO034.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO034.STCode,crSO034.STName,CD016NameMap);
        v_Note2 := GetNote(crSO034.STCode,crSO034.STName,'�ª��u����]');

        v_NewCode7 := GetNewCode(crSO034.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO034.CancelCode,crSO034.CancelName,CD051NameMap);

        v_AllNote := v_Note1||v_Note2;

        UPDATE SO034 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7,
                         Note=Note||v_AllNote
                     WHERE Rowid=crSO034.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO034���~  BillNo: '||crSO034.BillNo||'  Item:'||crSO034.Item,'SO034');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO034','�����SO034�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


--1111

--*****************************************************************
--*******************   ��  �s  �N  �X  ��   **********************
--*****************************************************************
/*
  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='CD006';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    --�m���¥N�X(�ΧR����ۦP�s�N�X���¥N�X)
    FOR crCD006 in ccCD006 loop
      BEGIN
        v_NewCode1 := GetNewCode(crCD006.CodeNo,CD006CodeMap);
        v_NewName1 := GetNewName(crCD006.CodeNo,CD006NameMap);
        v_Mark := GetMark(crCD006.CodeNo,CD006Mark);

        IF v_Mark=-1 THEN--�h���¥N�X�n��@�ӷs�N�X,�u��d�@��
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
            VALUES(C_UpdDataErrorType,'UPDATE �N�X��CD006���~  CodeNo: '||crCD006.CodeNo,'CD006');

      END;
    END LOOP;


    --�s�W�s�N�X
    FOR crCD006_MAP in ccCD006_MAP loop
      v_TempCounts := 0;
      SELECT COUNT(*) INTO v_TempCounts FROM CD006 WHERE CodeNo=crCD006_MAP.NewCode;
      IF v_TempCounts=0 THEN --�קK�s�W�ۦP���N�X
        BEGIN
          INSERT INTO CD006(CodeNo,Description,StopFlag,ServiceType,PrintWONow)
            VALUES(crCD006_MAP.NewCode,crCD006_MAP.NewName,C_StopFlag,C_ServiceType,C_PrintWONow);

          v_OKCounts := v_OKCounts + 1;
        EXCEPTION
          WHEN others THEN
            v_ErrorCounts := v_ErrorCounts + 1;

            INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
              VALUES(C_UpdDataErrorType,'INSERT �N�X��CD006���~  NewCodeNo: '||crCD006_MAP.NewCode,'CD006');

        END;
      ELSE
        v_ErrorCounts := v_ErrorCounts + 1;

        INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
          VALUES(C_UpdDataErrorType,'INSERT �N�X��CD006������  NewCodeNo: '||crCD006_MAP.NewCode,'CD006');

      END IF;
    END LOOP;


    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('CD006','�N�X��CD006�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
*/
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

  COMMIT;
  v_TotalAllTiem := GetDiffTime(v_TotalStartTime,SysDate);
  DBMS_OUTPUT.PUT_LINE('���槹��:' || v_TotalAllTiem);
exception
  WHEN OTHERS THEN
    v_ErrDesc := 'Error Code ='||SQLCODE||' '||'Error message='||SQLERRM;
    DBMS_OUTPUT.PUT_LINE(v_ErrDesc);

    INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
      VALUES(C_OtherErrorType,v_ErrDesc,'');
end;
/