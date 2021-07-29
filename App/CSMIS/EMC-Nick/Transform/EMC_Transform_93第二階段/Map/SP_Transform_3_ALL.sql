CREATE OR REPLACE PROCEDURE SP_TRANSFORM_3
  AS
/*
    �F�˥N�X����
    �ɦW: SP_TRANSFORM_3.sql
    By: Jackal   Date: 2004.01.30
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

  cursor ccCD024 is
    select Rowid,CodeNo,CitemCode,CitemName from CD024;

  cursor ccCD048A is
    select Rowid,CitemCodeStr from CD048A where CitemCodeStr is not null;
    
  cursor ccCD057 is
    select Rowid,CodeNo,ReturnCode,ReturnName from CD057;

  cursor ccCD057A is
    select Rowid,CitemCode1,CitemName1,CitemCode2,CitemName2,
           CitemCode3,CitemName3,CitemCode4,CitemName4,CitemCode5,CitemName5,
           FacilCode1,FacilName1,FacilCode2,FacilName2,FacilCode3,FacilName3,
           InstCode,InstName from CD057A;

  cursor ccCD059 is
    select Rowid,CodeNo,CompCode,DefCitemCode,DefCitemName from CD059;
    
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

  cursor ccSO020 is
    select Rowid,PowerNo,CompCode,BillNo,CountCode,CountName from SO020;

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

  cursor ccSO035 is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO035;

  cursor ccSO036 is
    select Rowid,BillNo,Item,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO036;

  cursor ccSO037 is
    select Rowid,BankCode,CustId,TranTime from SO037;

  cursor ccSO043 is
    select Rowid,CompCode,ServiceType,BankCode,CitemCode from SO043;


  cursor ccSO044 is
    select Rowid,CompCode,ServiceType,ClassCode,BankCode,UCCode,PWCode from SO044;

  cursor ccSO050 is
    select Rowid,BillNo,Item,ClassCode,CitemCode,CitemName,SCitemCode,
           UCCode,UCName,NSTCode,NSTName,STCode,STName,CancelCode,CancelName from SO050;

  cursor ccSO051 is
    select Rowid,CitemCode,CitemName,CitemCodeB,CitemNameB,UCCode,UCName,
           UCCodeB,UCNameB,STCode,STName,STCodeB,STNameB from SO051;

  cursor ccSO055 is
    select Rowid,FaciCode,FaciName,FaciCodeB,FaciNameB,MediaCode,MediaName,
           MediaCodeB,MediaNameB from SO055;

  cursor ccSO061 is
    select Rowid,BankCode,BankName,BankCodeB,BankNameB,CitemCode,CitemName,
           CitemCodeB,CitemNameB from SO061;

  cursor ccSO067 is
    select Rowid,ClassCode,BankCode,BankName,CitemCode,CitemName,
           SCitemCode,UCCode,UCName,STCode,STName,CancelCode,CancelName from SO067;

  cursor ccSO068 is
    select Rowid,CitemCode,CitemName from SO068;

  cursor ccSO071 is
    select Rowid,CustId,SNo,CompCode,ServiceType,InstCode from SO071;

  cursor ccSO072 is
    select Rowid,CustId,SNo,CompCode,ServiceType,ServiceCode from SO072;

  cursor ccSO073 is
    select Rowid,CustId,SNo,CompCode,ServiceType,PRCode from SO073;

  cursor ccSO074 is
    select Rowid,BillNo,Item,CitemCode,CitemName,STCode,STName,CancelCode,CancelName from SO074;

  cursor ccSO074A is
    select Rowid,BillNo,Item,CitemCode,CitemName,STCode,STName,CancelCode,CancelName from SO074A;

  cursor ccSO077 is
    select Rowid,BillNo,Item,CitemCode,CitemName,STCode,STName,CancelCode,CancelName from SO077;

  cursor ccSO079 is
    select Rowid,CitemCode from SO079;

  cursor ccSO089 is
    select Rowid,CitemCode,CitemName from SO089;

  cursor ccSO089B is
    select Rowid,CitemCode,CitemName from SO089B;

  cursor ccSO090 is
    select Rowid,CitemCode,CitemName from SO090;

  cursor ccSO091 is
    select Rowid,CitemCode,CitemName from SO091;

  cursor ccSO095 is
    select Rowid,BulletinCode,BulletinName from SO095;

  cursor ccSO096 is
    select Rowid,BulletinCode,BulletinName from SO096;

  cursor ccSO098 is
    select Rowid,BulletinCode,BulletinName,MediaCode,MediaName from SO098;

  cursor ccSO100 is
    select Rowid,InstCode,InstName from SO100;

  cursor ccSO101 is
    select Rowid,ClassCode1,ClassName1,ClassCode1B,ClassName1B,DelCode,DelName from SO101;

  cursor ccSO104 is
    select Rowid,InstCode from SO104;

  cursor ccSO105 is
    select Rowid,OrderNo,ServiceCode,ServDescCode,BulletinCode,MediaCode,MediaName,
           ReturnCode,ReturnName,CATVReturnCode,CATVReturnName,
           DVSReturnCode,DVSReturnName,CitemCodeStr from SO105;

  cursor ccSO105B is
    select Rowid,CitemCode,CitemName from SO105B;

  cursor ccSO105C is
    select Rowid,ArticleNo,ArticleName from SO105C;

  cursor ccSO105D is
    select Rowid,FaciCode,FaciName from SO105D;

  cursor ccSO105T is
    select Rowid,OrderNo,CompCode,BulletinCode,BulletinName,MediaCode,MediaName,
           ReturnCode,ReturnName,CATVReturnCode,CATVReturnName,
           CMReturnCode,CMReturnName,DVSReturnCode,DVSReturnName from SO105T;

  cursor ccSO106 is
    select Rowid,BankCode,BankName,MediaCode,MediaName from SO106;

  cursor ccSO107 is
    select Rowid,CitemCode,CitemName from SO107;

  cursor ccSO108 is
    select Rowid,BillNo,Item,CitemCode,CitemName from SO108;

  cursor ccSO109 is
    select Rowid,CitemCode,CitemName from SO109;

  cursor ccSO120 is
    select Rowid,CodeNo,Description from SO120;

  cursor ccSO121 is
    select Rowid,MediaCode,MediaName from SO121;

  cursor ccSO122 is
    select Rowid,CitemCode,CitemName,MediaCode,MediaName from SO122;

  cursor ccSO131 is
    select Rowid,CitemCode,CitemName,MediaCode,MediaName from SO131;

  cursor ccSO132 is
    select Rowid,CitemCode,CitemName from SO132;

  cursor ccSO134 is
    select Rowid,CitemCode,CitemName,MediaCode,MediaName from SO134;

  cursor ccSO301A is
    select Rowid,CitemCode from SO301A;

  cursor ccSO301B is
    select Rowid,CitemCode from SO301B;

  cursor ccSO302A is
    select Rowid,CitemCode from SO302A;

  cursor ccSO302B is
    select Rowid,CitemCode from SO302B;

  cursor ccSO303 is
    select Rowid,CitemCode from SO303;

  cursor ccSO503 is
    select Rowid,MediaCode,MediaName from SO503;

  cursor ccSO508 is
    select Rowid,CitemCode,CitemName from SO508;

  cursor ccSO508A is
    select Rowid,CitemCode,CitemName from SO508A;

  cursor ccSO511A is
    select Rowid,ClassCode from SO511A;

  cursor ccSO511B is
    select Rowid,ClassCode from SO511B;

  cursor ccSO511C is
    select Rowid,ClassCode from SO511C;

  cursor ccSO511Z is
    select Rowid,ClassCode from SO511Z;

  cursor ccSOAC0102 is
    select Rowid,CitemCode,CitemName from SOAC0102;

  cursor ccSOAC0402 is
    select Rowid,CitemCode,CitemName from SOAC0402;


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

  CD041CodeMap NumberAry;
  CD041Mark NumberAry;
  CD041NameMap Varchar2Ary;

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


  -- ��M CD041 �̤j���¥N�X��,�n�@���}�C���̤jIndex
  Select Max(CodeNo) into v_MaxOldCode from CD041;

  IF v_MaxOldCode IS NOT NULL THEN
    FOR I IN 1 .. v_MaxOldCode LOOP
      BEGIN
        Select OldCode,NewCode,NewName,Mark into v_OldCode,v_NewCode,v_NewName,v_Mark
          from CD041_MAP WHERE OldCode=I;

        --�N��Ӫ�CD041_MAP�ন�}�C
        CD041CodeMap(I) := v_NewCode;
        CD041NameMap(I) := v_NewName;
        CD041Mark(I) := v_Mark;
      EXCEPTION
        WHEN too_many_rows then
          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_CodeMapErrorType,'�¥N�X[ '||I||' ]�b����ɤ����ⵧ�H�W','CD041_MAP');

          COMMIT;
          DBMS_OUTPUT.PUT_LINE('�¥N�X[ '||I||' ]�bCD041_MAP����ɤ����ⵧ�H�W');
          RETURN;
        WHEN OTHERS then
          CD041CodeMap(I) := I;
          CD041NameMap(I) := null;
          CD041Mark(I) := 0;
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


  --��sSO020���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO020';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO020 in ccSO020 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO020.CountCode,CD041CodeMap);
        v_NewName1 := GetNewName(crSO020.CountCode,crSO020.CountName,CD041NameMap);


        UPDATE SO020 set CountCode=v_NewCode1,CountName=v_NewName1
                     WHERE Rowid=crSO020.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO020���~  PowerNo: '||crSO020.PowerNo||'  CompCode:'||crSO020.CompCode||'  BillNo:'||crSO020.BillNo,'SO020');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO020','�����SO020�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

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


  --��sSO035���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO035';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;
    v_Note1 := '';
    v_Note2 := '';
    v_AllNote := '';

    FOR crSO035 in ccSO035 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO035.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO035.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO035.BankCode,crSO035.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO035.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO035.CitemCode,crSO035.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO035.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO035.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO035.UCCode,crSO035.UCName,CD013NameMap);
        v_Note1 := GetNote(crSO035.UCCode,crSO035.UCName,'�ª���ú�O��]');

        v_NewCode6 := GetNewCode(crSO035.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO035.STCode,crSO035.STName,CD016NameMap);
        v_Note2 := GetNote(crSO035.STCode,crSO035.STName,'�ª��u����]');

        v_NewCode7 := GetNewCode(crSO035.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO035.CancelCode,crSO035.CancelName,CD051NameMap);

        v_AllNote := v_Note1||v_Note2;

        UPDATE SO035 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7,
                         Note=Note||v_AllNote
                     WHERE Rowid=crSO035.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO035���~  BillNo: '||crSO035.BillNo||'  Item:'||crSO035.Item,'SO035');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO035','�����SO035�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO036���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO036';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;
    v_Note1 := '';
    v_Note2 := '';
    v_AllNote := '';

    FOR crSO036 in ccSO036 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO036.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO036.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO036.BankCode,crSO036.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO036.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO036.CitemCode,crSO036.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO036.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO036.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO036.UCCode,crSO036.UCName,CD013NameMap);
        v_Note1 := GetNote(crSO036.UCCode,crSO036.UCName,'�ª���ú�O��]');

        v_NewCode6 := GetNewCode(crSO036.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO036.STCode,crSO036.STName,CD016NameMap);
        v_Note2 := GetNote(crSO036.STCode,crSO036.STName,'�ª��u����]');

        v_NewCode7 := GetNewCode(crSO036.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO036.CancelCode,crSO036.CancelName,CD051NameMap);

        v_AllNote := v_Note1||v_Note2;

        UPDATE SO036 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7,
                         Note=Note||v_AllNote
                     WHERE Rowid=crSO036.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO036���~  BillNo: '||crSO036.BillNo||'  Item:'||crSO036.Item,'SO036');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO036','�����SO036�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO037���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO037';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO037 in ccSO037 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO037.BankCode,CD018CodeMap);

        UPDATE SO037 set BankCode=v_NewCode1
                     WHERE Rowid=crSO037.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO037���~  BankCode: '||crSO037.BankCode||' CustID: '||crSO037.CustID||' TranTime: '||crSO037.TranTime,'SO037');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO037','�����SO037�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO043���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO043';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO043 in ccSO043 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO043.BankCode,CD018CodeMap);

      	v_NewCode2 := GetNewCode(crSO043.CitemCode,CD019CodeMap);

        UPDATE SO043 set BankCode=v_NewCode1,
                         CitemCode=v_NewCode
                     WHERE Rowid=crSO043.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO043���~  ServiceType: '||crSO043.ServiceType||' CompCode: '||crSO043.CompCode,'SO043');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO043','�����SO043�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO044���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO044';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO044 in ccSO044 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO044.ClassCode,CD004CodeMap);

      	v_NewCode2 := GetNewCode(crSO044.BankCode,CD018CodeMap);

        v_NewCode3 := GetNewCode(crSO044.UCCode,CD013CodeMap);

      	v_NewCode4 := GetNewCode(crSO044.PWCode,CD020CodeMap);

        UPDATE SO044 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,
                         UCCode=v_NewCode3,
                         PWCode=v_NewCode4
                     WHERE Rowid=crSO044.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO044���~  ServiceType: '||crSO044.ServiceType||' CompCode: '||crSO044.CompCode,'SO044');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO044','�����SO044�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO050���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO050';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO050 in ccSO050 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO050.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO050.CitemCode,CD019CodeMap);
        v_NewName2 := GetNewName(crSO050.CitemCode,crSO050.CitemName,CD019NameMap);

        v_NewCode3 := GetNewCode(crSO050.SCitemCode,CD019CodeMap);

        v_NewCode4 := GetNewCode(crSO050.UCCode,CD013CodeMap);
        v_NewName4 := GetNewName(crSO050.UCCode,crSO050.UCName,CD013NameMap);

        v_NewCode5 := GetNewCode(crSO050.NSTCode,CD016CodeMap);
        v_NewName5 := GetNewName(crSO050.NSTCode,crSO050.NSTName,CD016NameMap);

        v_NewCode6 := GetNewCode(crSO050.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO050.STCode,crSO050.STName,CD016NameMap);

        v_NewCode7 := GetNewCode(crSO050.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO050.CancelCode,crSO050.CancelName,CD051NameMap);


        UPDATE SO050 set ClassCode=v_NewCode1,
                         CitemCode=v_NewCode2,CitemName=v_NewName2,
                         SCitemCode=v_NewCode3,
                         UCCode=v_NewCode4,UCName=v_NewName4,
                         NSTCode=v_NewCode5,NSTName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7
                     WHERE Rowid=crSO050.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO050���~  ClassCode: '||crSO050.ClassCode||
                                                             ' CitemCode: '||crSO050.CitemCode||
                                                             ' SCitemCode: '||crSO050.SCitemCode||
                                                             ' UCCode: '||crSO050.UCCode||
                                                             ' NSTCode:'||crSO050.NSTCode||
                                                             '  STCode:'||crSO050.STCode||
                                                             '  CancelCode:'||crSO050.CancelCode,'SO050');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO050','�����SO050�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO051���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO051';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO051 in ccSO051 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO051.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO051.CitemCode,crSO051.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO051.CitemCodeB,CD019CodeMap);
        v_NewName2 := GetNewName(crSO051.CitemCodeB,crSO051.CitemNameB,CD019NameMap);

        v_NewCode3 := GetNewCode(crSO051.UCCode,CD013CodeMap);
        v_NewName3 := GetNewName(crSO051.UCCode,crSO051.UCName,CD013NameMap);

        v_NewCode4 := GetNewCode(crSO051.UCCodeB,CD013CodeMap);
        v_NewName4 := GetNewName(crSO051.UCCodeB,crSO051.UCNameB,CD013NameMap);

        v_NewCode5 := GetNewCode(crSO051.STCode,CD016CodeMap);
        v_NewName5 := GetNewName(crSO051.STCode,crSO051.STName,CD016NameMap);

        v_NewCode6 := GetNewCode(crSO051.STCodeB,CD016CodeMap);
        v_NewName6 := GetNewName(crSO051.STCodeB,crSO051.STNameB,CD016NameMap);


        UPDATE SO051 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         CitemCodeB=v_NewCode2,CitemNameB=v_NewName2,
                         UCCode=v_NewCode3,UCName=v_NewName3,
                         UCCodeB=v_NewCode4,UCNameB=v_NewName4,
                         STCode=v_NewCode5,STName=v_NewName5,
                         STCodeB=v_NewCode6,STNameB=v_NewName6
                     WHERE Rowid=crSO051.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO051���~  UCCode: '||crSO051.UCCode||' UCCodeB:'||crSO051.UCCodeB||'  STCode:'||crSO051.STCode||'  STCodeB:'||crSO051.STCodeB||'  CitemCode:'||crSO051.CitemCode||'  CitemCodeB:'||crSO051.CitemCodeB,'SO051');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO051','�����SO051�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO055���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO055';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO055 in ccSO055 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO055.FaciCode,CD022CodeMap);
        v_NewName1 := GetNewName(crSO055.FaciCode,crSO055.FaciName,CD022NameMap);

        v_NewCode2 := GetNewCode(crSO055.FaciCodeB,CD022CodeMap);
        v_NewName2 := GetNewName(crSO055.FaciCodeB,crSO055.FaciNameB,CD022NameMap);

        v_NewCode3 := GetNewCode(crSO055.MediaCode,CD009CodeMap);
        v_NewName3 := GetNewName(crSO055.MediaCode,crSO055.MediaName,CD009NameMap);

        v_NewCode4 := GetNewCode(crSO055.MediaCodeB,CD009CodeMap);
        v_NewName4 := GetNewName(crSO055.MediaCodeB,crSO055.MediaNameB,CD009NameMap);


        UPDATE SO055 set FaciCode=v_NewCode1,FaciName=v_NewName1,
                         FaciCodeB=v_NewCode2,FaciNameB=v_NewName2,
                         MediaCode=v_NewCode3,MediaName=v_NewName3,
                         MediaCodeB=v_NewCode4,MediaNameB=v_NewName4
                     WHERE Rowid=crSO055.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO055���~  MediaCode: '||crSO055.MediaCode||'  MediaCodeB:'||crSO055.MediaCodeB||
                   'FaciCode: '||crSO055.FaciCode||'  FaciaCodeB:'||crSO055.FaciCodeB,'SO055');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO055','�����SO055�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO061���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO061';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO061 in ccSO061 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO061.BankCode,CD018CodeMap);
        v_NewName1 := GetNewName(crSO061.BankCode,crSO061.BankName,CD018NameMap);

        v_NewCode2 := GetNewCode(crSO061.BankCodeB,CD018CodeMap);
        v_NewName2 := GetNewName(crSO061.BankCodeB,crSO061.BankNameB,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO061.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO061.CitemCode,crSO061.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO061.CitemCodeB,CD019CodeMap);
        v_NewName4 := GetNewName(crSO061.CitemCodeB,crSO061.CitemNameB,CD019NameMap);


        UPDATE SO061 set BankCode=v_NewCode1,BankName=v_NewName1,
                         BankCodeB=v_NewCode2,BankNameB=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         CitemCodeB=v_NewCode4,CitemNameB=v_NewName4
                     WHERE Rowid=crSO061.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO061���~  CitemCode: '||crSO061.CitemCode||'  CitemCodeB:'||crSO061.CitemCodeB||
                   'BankCode: '||crSO061.BankCode||'  BankaCodeB:'||crSO061.BankCodeB,'SO061');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO061','�����SO061�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO067���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO067';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO067 in ccSO067 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO067.ClassCode,CD004CodeMap);

        v_NewCode2 := GetNewCode(crSO067.BankCode,CD018CodeMap);
        v_NewName2 := GetNewName(crSO067.BankCode,crSO067.BankName,CD018NameMap);

        v_NewCode3 := GetNewCode(crSO067.CitemCode,CD019CodeMap);
        v_NewName3 := GetNewName(crSO067.CitemCode,crSO067.CitemName,CD019NameMap);

        v_NewCode4 := GetNewCode(crSO067.SCitemCode,CD019CodeMap);

        v_NewCode5 := GetNewCode(crSO067.UCCode,CD013CodeMap);
        v_NewName5 := GetNewName(crSO067.UCCode,crSO067.UCName,CD013NameMap);

        v_NewCode6 := GetNewCode(crSO067.STCode,CD016CodeMap);
        v_NewName6 := GetNewName(crSO067.STCode,crSO067.STName,CD016NameMap);

        v_NewCode7 := GetNewCode(crSO067.CancelCode,CD051CodeMap);
        v_NewName7 := GetNewName(crSO067.CancelCode,crSO067.CancelName,CD051NameMap);


        UPDATE SO067 set ClassCode=v_NewCode1,
                         BankCode=v_NewCode2,BankName=v_NewName2,
                         CitemCode=v_NewCode3,CitemName=v_NewName3,
                         SCitemCode=v_NewCode4,
                         UCCode=v_NewCode5,UCName=v_NewName5,
                         STCode=v_NewCode6,STName=v_NewName6,
                         CancelCode=v_NewCode7,CancelName=v_NewName7
                     WHERE Rowid=crSO067.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO067���~  ClassCode: '||crSO067.ClassCode||
                                                            '  BankCode:'||crSO067.BankCode||
                                                            '  CitemCode:'||crSO067.CitemCode||
                                                            '  SCitemCode:'||crSO067.SCitemCode||
                                                            '  UCCode:'||crSO067.UCCode||
                                                            '  STCode:'||crSO067.STCode||
                                                            '  CancelCode:'||crSO067.CancelCode
                                                            ,'SO067');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO067','�����SO067�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO068���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO068';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO068 in ccSO068 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO068.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO068.CitemCode,crSO068.CitemName,CD019NameMap);


        UPDATE SO068 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO068.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO068���~  CitemCode: '||crSO068.CitemCode,'SO068');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO068','�����SO068�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO071���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO071';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO071 in ccSO071 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO071.InstCode,CD005CodeMap);


        UPDATE SO071 set InstCode=v_NewCode1
                     WHERE Rowid=crSO071.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO071���~  CustId: '||crSO071.CustId||' SNo: '||crSO071.SNo||' CompCode: '||crSO071.CompCode||' ServiceType: '||crSO071.ServiceType,'SO071');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO071','�����SO071�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO072���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO072';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO072 in ccSO072 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO072.ServiceCode,CD006CodeMap);


        UPDATE SO072 set ServiceCode=v_NewCode1
                     WHERE Rowid=crSO072.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO072���~  CustId: '||crSO072.CustId||' SNo: '||crSO072.SNo||' CompCode: '||crSO072.CompCode||' ServiceType: '||crSO072.ServiceType,'SO072');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO072','�����SO072�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO073���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO073';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO073 in ccSO073 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO073.PRCode,CD007CodeMap);


        UPDATE SO073 set PRCode=v_NewCode1
                     WHERE Rowid=crSO073.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO073���~  CustId: '||crSO073.CustId||' SNo: '||crSO073.SNo||' CompCode: '||crSO073.CompCode||' ServiceType: '||crSO073.ServiceType,'SO073');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO073','�����SO073�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO074���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO074';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO074 in ccSO074 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO074.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO074.CitemCode,crSO074.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO074.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO074.STCode,crSO074.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO074.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO074.CancelCode,crSO074.CancelName,CD051NameMap);


        UPDATE SO074 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO074.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO074���~  BillNo: '||crSO074.BillNo||'  Item:'||crSO074.Item,'SO074');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO074','�����SO074�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO074A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO074A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO074A in ccSO074A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO074A.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO074A.CitemCode,crSO074A.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO074A.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO074A.STCode,crSO074A.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO074A.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO074A.CancelCode,crSO074A.CancelName,CD051NameMap);


        UPDATE SO074A set CitemCode=v_NewCode1,CitemName=v_NewName1,
                          STCode=v_NewCode2,STName=v_NewName2,
                          CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO074A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO074A���~  BillNo: '||crSO074A.BillNo||'  Item:'||crSO074A.Item,'SO074A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO074A','�����SO074A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO077���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO077';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO077 in ccSO077 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO077.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO077.CitemCode,crSO077.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO077.STCode,CD016CodeMap);
        v_NewName2 := GetNewName(crSO077.STCode,crSO077.STName,CD016NameMap);

        v_NewCode3 := GetNewCode(crSO077.CancelCode,CD051CodeMap);
        v_NewName3 := GetNewName(crSO077.CancelCode,crSO077.CancelName,CD051NameMap);


        UPDATE SO077 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         STCode=v_NewCode2,STName=v_NewName2,
                         CancelCode=v_NewCode3,CancelName=v_NewName3
                     WHERE Rowid=crSO077.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO077���~  BillNo: '||crSO077.BillNo||'  Item:'||crSO077.Item,'SO077');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO077','�����SO077�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO079���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO079';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO079 in ccSO079 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO079.CitemCode,CD019CodeMap);


        UPDATE SO079 set CitemCode=v_NewCode1
                     WHERE Rowid=crSO079.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO079���~  CitemCode: '||crSO079.CitemCode,'SO079');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO079','�����SO079�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO089���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO089';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO089 in ccSO089 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO089.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO089.CitemCode,crSO089.CitemName,CD019NameMap);


        UPDATE SO089 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO089.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO089���~  CitemCode: '||crSO089.CitemCode,'SO089');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO089','�����SO089�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO089B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO089B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO089B in ccSO089B loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO089B.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO089B.CitemCode,crSO089B.CitemName,CD019NameMap);


        UPDATE SO089B set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO089B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO089B���~  CitemCode: '||crSO089B.CitemCode,'SO089B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO089B','�����SO089B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO090���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO090';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO090 in ccSO090 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO090.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO090.CitemCode,crSO090.CitemName,CD019NameMap);


        UPDATE SO090 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO090.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO090���~  CitemCode: '||crSO090.CitemCode,'SO090');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO090','�����SO090�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO091���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO091';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO091 in ccSO091 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO091.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO091.CitemCode,crSO091.CitemName,CD019NameMap);


        UPDATE SO091 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO091.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO091���~  CitemCode: '||crSO091.CitemCode,'SO091');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO091','�����SO091�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO095���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO095';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO095���~  BulletinCode: '||crSO095.BulletinCode,'SO095');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO095','�����SO095�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO096���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO096';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO096���~  BulletinCode: '||crSO096.BulletinCode,'SO096');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO096','�����SO096�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO098���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO098';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO098 in ccSO098 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO098.BulletinCode,CD049CodeMap);
        v_NewName1 := GetNewName(crSO098.BulletinCode,crSO098.BulletinName,CD049NameMap);

        v_NewCode2 := GetNewCode(crSO098.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO098.MediaCode,crSO098.MediaName,CD009NameMap);

        UPDATE SO098 set BulletinCode=v_NewCode1,BulletinName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO098.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO098���~  BulletinCode: '||crSO098.BulletinCode||' MediaCode: '||crSO098.MediaCode,'SO098');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO098','�����SO098�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO100���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO100';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO100 in ccSO100 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO100.InstCode,CD005CodeMap);
        v_NewName1 := GetNewName(crSO100.InstCode,crSO100.InstName,CD005NameMap);

        UPDATE SO100 set InstCode=v_NewCode1,InstName=v_NewName1
                     WHERE Rowid=crSO100.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO100���~  InstCode: '||crSO100.InstCode,'SO100');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO100','�����SO100�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO101���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO101';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO101 in ccSO101 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO101.ClassCode1,CD004CodeMap);
        v_NewName1 := GetNewName(crSO101.ClassCode1,crSO101.ClassName1,CD004NameMap);

        v_NewCode2 := GetNewCode(crSO101.ClassCode1B,CD004CodeMap);
        v_NewName2 := GetNewName(crSO101.ClassCode1B,crSO101.ClassName1B,CD004NameMap);

        v_NewCode3 := GetNewCode(crSO101.DelCode,CD012CodeMap);
        v_NewName3 := GetNewName(crSO101.DelCode,crSO101.DelName,CD012NameMap);


        UPDATE SO101 set ClassCode1=v_NewCode1,ClassName1=v_NewName1,
                         ClassCode1B=v_NewCode2,ClassName1B=v_NewName2,
                         DelCode=v_NewCode3,DelName=v_NewName3
                     WHERE Rowid=crSO101.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO101���~  ClassCode1: '||crSO101.ClassCode1||'  ClassCode1B:'||crSO101.ClassCode1B||'  DelCode:'||crSO101.DelCode,'SO101');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO101','�����SO101�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO104���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO104';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO104 in ccSO104 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO104.InstCode,CD005CodeMap);

        UPDATE SO104 set InstCode=v_NewCode1
                     WHERE Rowid=crSO104.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO104���~  InstCode: '||crSO104.InstCode,'SO104');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO104','�����SO104�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO105���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105 in ccSO105 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105.ServiceCode,CD008CodeMap);

        v_NewCode2 := GetNewCode(crSO105.ServDescCode,CD008ACodeMap);

        v_NewCode3 := GetNewCode(crSO105.BulletinCode,CD049CodeMap);

        v_NewCode4 := GetNewCode(crSO105.MediaCode,CD009CodeMap);
        v_NewName4 := GetNewName(crSO105.MediaCode,crSO105.MediaName,CD009NameMap);

        v_NewCode5 := GetNewCode(crSO105.ReturnCode,CD015CodeMap);
        v_NewName5 := GetNewName(crSO105.ReturnCode,crSO105.ReturnName,CD015NameMap);

        v_NewCode6 := GetNewCode(crSO105.CATVReturnCode,CD015CodeMap);
        v_NewName6 := GetNewName(crSO105.CATVReturnCode,crSO105.CATVReturnName,CD015NameMap);

        v_NewCode7 := GetNewCode(crSO105.DVSReturnCode,CD015CodeMap);
        v_NewName7 := GetNewName(crSO105.DVSReturnCode,crSO105.DVSReturnName,CD015NameMap);
        
        V_CitemCodeStr := ChangeCitemCodeStr(crSO105.CitemCodeStr,CD019CodeMap);


        UPDATE SO105 set ServiceCode=v_NewCode1,
                         ServDescCode=v_NewCode2,
                         BulletinCode=v_NewCode3,
                         MediaCode=v_NewCode4,MediaName=v_NewName4,
                         ReturnCode=v_NewCode5,ReturnName=v_NewName5,
                         CATVReturnCode=v_NewCode6,CATVReturnName=v_NewName6,
                         DVSReturnCode=v_NewCode7,DVSReturnName=v_NewName7,
                         CitemCodeStr=V_CitemCodeStr
                     WHERE Rowid=crSO105.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO105���~  OrderNo: '||crSO105.OrderNo,'SO105');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105','�����SO105�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

/*JACKAL
  --��sSO105B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105B in ccSO105B loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105B.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO105B.CitemCode,crSO105B.CitemName,CD019NameMap);

        UPDATE SO105B set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO105B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO105B���~  CitemCode: '||crSO105B.CitemCode,'SO105B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105B','�����SO105B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO105C���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105C';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105C in ccSO105C loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105C.ArticleNo,CD059CodeMap);
        v_NewName1 := GetNewName(crSO105C.ArticleNo,crSO105C.ArticleName,CD059NameMap);

        UPDATE SO105C set ArticleNo=v_NewCode1,ArticleName=v_NewName1
                     WHERE Rowid=crSO105C.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO105C���~  ArticleNo: '||crSO105C.ArticleNo,'SO105C');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105C','�����SO105C�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO105D���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105D';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105D in ccSO105D loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105D.FaciCode,CD022CodeMap);
        v_NewName1 := GetNewName(crSO105D.FaciCode,crSO105D.FaciName,CD022NameMap);

        UPDATE SO105D set FaciCode=v_NewCode1,FaciName=v_NewName1
                     WHERE Rowid=crSO105D.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO105D���~  FaciCode: '||crSO105D.FaciCode,'SO105D');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105D','�����SO105D�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO105T���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO105T';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO105T in ccSO105T loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO105T.BulletinCode,CD049CodeMap);
        v_NewName1 := GetNewName(crSO105T.BulletinCode,crSO105T.BulletinName,CD049NameMap);

        v_NewCode2 := GetNewCode(crSO105T.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO105T.MediaCode,crSO105T.MediaName,CD009NameMap);

        v_NewCode3 := GetNewCode(crSO105T.ReturnCode,CD015CodeMap);
        v_NewName3 := GetNewName(crSO105T.ReturnCode,crSO105T.ReturnName,CD015NameMap);

        v_NewCode4 := GetNewCode(crSO105T.CATVReturnCode,CD015CodeMap);
        v_NewName4 := GetNewName(crSO105T.CATVReturnCode,crSO105T.CATVReturnName,CD015NameMap);

        v_NewCode5 := GetNewCode(crSO105T.CMReturnCode,CD015CodeMap);
        v_NewName5 := GetNewName(crSO105T.CMReturnCode,crSO105T.CMReturnName,CD015NameMap);

        v_NewCode6 := GetNewCode(crSO105T.DVSReturnCode,CD015CodeMap);
        v_NewName6 := GetNewName(crSO105T.DVSReturnCode,crSO105T.DVSReturnName,CD015NameMap);



        UPDATE SO105T set BulletinCode=v_NewCode1,BulletinName=v_NewName1,
                          MediaCode=v_NewCode2,MediaName=v_NewName2,
                          ReturnCode=v_NewCode3,ReturnName=v_NewName3,
                          CATVReturnCode=v_NewCode4,CATVReturnName=v_NewName4,
                          CMReturnCode=v_NewCode5,CMReturnName=v_NewName5,
                          DVSReturnCode=v_NewCode6,DVSReturnName=v_NewName6
                     WHERE Rowid=crSO105T.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO105T���~  OrderNo: '||crSO105T.OrderNo||'  CompCode'||crSO105T.CompCode,'SO105T');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO105T','�����SO105T�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO106���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO106';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO106 in ccSO106 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO106.BankCode,CD018CodeMap);
        v_NewName1 := GetNewName(crSO106.BankCode,crSO106.BankName,CD018NameMap);

        v_NewCode2 := GetNewCode(crSO106.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO106.MediaCode,crSO106.MediaName,CD009NameMap);


        UPDATE SO106 set BankCode=v_NewCode1,BankName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO106.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO106���~  BankCode: '||crSO106.BankCode||' MediaCode: '||crSO106.MediaCode,'SO106');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO106','�����SO106�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO107���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO107';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO107 in ccSO107 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO107.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO107.CitemCode,crSO107.CitemName,CD019NameMap);

        UPDATE SO107 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO107.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO107���~  CitemCode: '||crSO107.CitemCode,'SO107');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO107','�����SO107�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO108���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO108';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO108 in ccSO108 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO108.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO108.CitemCode,crSO108.CitemName,CD019NameMap);

        UPDATE SO108 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO108.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO108���~  BillNo: '||crSO108.BillNo||' Item: '||crSO108.Item||' CitemCode: '||crSO108.CitemCode,'SO108');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO108','�����SO108�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



  --��sSO109���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO109';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO109 in ccSO109 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO109.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO109.CitemCode,crSO109.CitemName,CD019NameMap);

        UPDATE SO109 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO109.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO109���~  CitemCode: '||crSO109.CitemCode,'SO109');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO109','�����SO109�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO120���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO120';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO120 in ccSO120 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO120.CodeNo,CD019CodeMap);
        v_NewName1 := GetNewName(crSO120.CodeNo,crSO120.Description,CD019NameMap);

        UPDATE SO120 set CodeNo=v_NewCode1,Description=v_NewName1
                     WHERE Rowid=crSO120.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO120���~  CodeNo: '||crSO120.CodeNo,'SO120');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO120','�����SO120�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO121���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO121';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO121���~  MediaCode: '||crSO121.MediaCode,'SO121');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO121','�����SO121�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO122���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO122';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO122 in ccSO122 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO122.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO122.CitemCode,crSO122.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO122.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO122.MediaCode,crSO122.MediaName,CD009NameMap);


        UPDATE SO122 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO122.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO122���~  CitemCode: '||crSO122.CitemCode||' MediaCode: '||crSO122.MediaCode,'SO122');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO122','�����SO122�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO131���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO131';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO131 in ccSO131 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO131.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO131.CitemCode,crSO131.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO131.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO131.MediaCode,crSO131.MediaName,CD009NameMap);


        UPDATE SO131 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO131.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO131���~  CitemCode: '||crSO131.CitemCode||' MediaCode: '||crSO131.MediaCode,'SO131');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO131','�����SO131�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO132���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO132';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO132 in ccSO132 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO132.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO132.CitemCode,crSO132.CitemName,CD019NameMap);

        UPDATE SO132 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO132.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO132���~  CitemCode: '||crSO132.CitemCode,'SO132');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO132','�����SO132�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



  --��sSO134���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO134';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO134 in ccSO134 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO134.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO134.CitemCode,crSO134.CitemName,CD019NameMap);

        v_NewCode2 := GetNewCode(crSO134.MediaCode,CD009CodeMap);
        v_NewName2 := GetNewName(crSO134.MediaCode,crSO134.MediaName,CD009NameMap);


        UPDATE SO134 set CitemCode=v_NewCode1,CitemName=v_NewName1,
                         MediaCode=v_NewCode2,MediaName=v_NewName2
                     WHERE Rowid=crSO134.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO134���~  CitemCode: '||crSO134.CitemCode||' MediaCode: '||crSO134.MediaCode,'SO134');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO134','�����SO134�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO301A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO301A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO301A in ccSO301A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO301A.CitemCode,CD019CodeMap);

        UPDATE SO301A set CitemCode=v_NewCode1
                     WHERE Rowid=crSO301A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO301A���~  CitemCode: '||crSO301A.CitemCode,'SO301A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO301A','�����SO301A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO301B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO301B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO301B in ccSO301B loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO301B.CitemCode,CD019CodeMap);

        UPDATE SO301B set CitemCode=v_NewCode1
                     WHERE Rowid=crSO301B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO301B���~  CitemCode: '||crSO301B.CitemCode,'SO301B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO301B','�����SO301B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO302A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO302A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO302A in ccSO302A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO302A.CitemCode,CD019CodeMap);

        UPDATE SO302A set CitemCode=v_NewCode1
                     WHERE Rowid=crSO302A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO302A���~  CitemCode: '||crSO302A.CitemCode,'SO302A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO302A','�����SO302A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO302B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO302B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO302B in ccSO302B loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO302B.CitemCode,CD019CodeMap);

        UPDATE SO302B set CitemCode=v_NewCode1
                     WHERE Rowid=crSO302B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO302B���~  CitemCode: '||crSO302B.CitemCode,'SO302B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO302B','�����SO302B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO303���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO303';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO303 in ccSO303 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO303.CitemCode,CD019CodeMap);

        UPDATE SO303 set CitemCode=v_NewCode1
                     WHERE Rowid=crSO303.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO303���~  CitemCode: '||crSO303.CitemCode,'SO303');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO303','�����SO303�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO503���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO503';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
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
            VALUES(C_UpdDataErrorType,'UPDATE �����SO503���~  MediaCode: '||crSO503.MediaCode,'SO503');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO503','�����SO503�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO508���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO508';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO508 in ccSO508 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO508.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO508.CitemCode,crSO508.CitemName,CD019NameMap);

        UPDATE SO508 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO508.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO508���~  CitemCode: '||crSO508.CitemCode,'SO508');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO508','�����SO508�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO508A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO508A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO508A in ccSO508A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO508A.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSO508A.CitemCode,crSO508A.CitemName,CD019NameMap);

        UPDATE SO508A set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSO508A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO508A���~  CitemCode: '||crSO508A.CitemCode,'SO508A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO508A','�����SO508A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO511A���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO511A';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO511A in ccSO511A loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO511A.ClassCode,CD004CodeMap);

        UPDATE SO511A set ClassCode=v_NewCode1
                     WHERE Rowid=crSO511A.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO511A���~  ClassCode: '||crSO511A.ClassCode,'SO511A');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO511A','�����SO511A�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO511B���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO511B';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO511B in ccSO511B loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO511B.ClassCode,CD004CodeMap);

        UPDATE SO511B set ClassCode=v_NewCode1
                     WHERE Rowid=crSO511B.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO511B���~  ClassCode: '||crSO511B.ClassCode,'SO511B');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO511B','�����SO511B�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO511C���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO511C';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO511C in ccSO511C loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO511C.ClassCode,CD004CodeMap);

        UPDATE SO511C set ClassCode=v_NewCode1
                     WHERE Rowid=crSO511C.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO511C���~  ClassCode: '||crSO511C.ClassCode,'SO511C');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO511C','�����SO511C�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSO511Z���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SO511Z';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSO511Z in ccSO511Z loop
      BEGIN
        v_NewCode1 := GetNewCode(crSO511Z.ClassCode,CD004CodeMap);

        UPDATE SO511Z set ClassCode=v_NewCode1
                     WHERE Rowid=crSO511Z.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SO511Z���~  ClassCode: '||crSO511Z.ClassCode,'SO511Z');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SO511Z','�����SO511Z�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSOAC0102���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SOAC0102';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSOAC0102 in ccSOAC0102 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSOAC0102.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSOAC0102.CitemCode,crSOAC0102.CitemName,CD019NameMap);

        UPDATE SOAC0102 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSOAC0102.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SOAC0102���~  CitemCode: '||crSOAC0102.CitemCode,'SOAC0102');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SOAC0102','�����SOAC0102�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


  --��sSOAC0402���
  v_TempCounts := SetSegment(v_HaveRBS27);

  v_TempCounts :=0;
  SELECT Count(*) INTO v_TempCounts FROM TransFormLog_1 WHERE TABLENAME='SOAC0402';

  IF v_TempCounts=0 THEN --��N�X�e���T�w�S����L
    v_OKCounts := 0;
    v_ErrorCounts := 0;
    v_StartTime := SysDate;

    FOR crSOAC0402 in ccSOAC0402 loop
      BEGIN
        v_NewCode1 := GetNewCode(crSOAC0402.CitemCode,CD019CodeMap);
        v_NewName1 := GetNewName(crSOAC0402.CitemCode,crSOAC0402.CitemName,CD019NameMap);

        UPDATE SOAC0402 set CitemCode=v_NewCode1,CitemName=v_NewName1
                     WHERE Rowid=crSOAC0402.Rowid;


        v_OKCounts := v_OKCounts + 1;
      EXCEPTION
        WHEN others THEN
          v_ErrorCounts := v_ErrorCounts + 1;

          INSERT INTO TransFormLog_2(TYPE,DESCRIPTION,TABLENAME)
            VALUES(C_UpdDataErrorType,'UPDATE �����SOAC0402���~  CitemCode: '||crSOAC0402.CitemCode,'SOAC0402');

      END;
    END LOOP;

    v_AllTiem := GetDiffTime(v_StartTime,SysDate);

    INSERT INTO TransFormLog_1(TABLENAME,DESCRIPTION,OKCOUNTS,ERRORCOUNTS,STARTTIME,STOPTIME,ALLTIME)
      VALUES('SOAC0402','�����SOAC0402�N�X�ഫ����',v_OKCounts,v_ErrorCounts,v_StartTime,SysDate,v_AllTiem);

    COMMIT;
  END IF;
JACKAL*/
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