unit E2CImp;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, prjE2C_TLB, StdVcl, Classes, ADODB, SysUtils, inifiles, Forms,
  Controls,Dialogs, MSXML_TLB, Variants;


const
    VENDORNAME_LENGTH = 20;
    VENDORCODE_LENGTH = 5;
    CONN_TO_WAREHOUSE = True;
    DATA_AREA_HEADER = 'DA';
    STR_SEP = '''';

    SYS_INI_FILE_NAME = 'C:\WINNT\E2C_DLL.ini';
    TMP_SYS_INI_FILE_NAME = 'C:\WINNT\E2C_DLL_TEMP.ini';

    //SYS_INI_FILE_NAME = 'D:\App\EMC\MSS\Project\E2C\E2C_DLL.ini';
    //TMP_SYS_INI_FILE_NAME = 'D:\App\EMC\MSS\Project\E2C\E2C_DLL_TEMP.ini';


    BATCH_NO_LENGTH = 12;             //�帹���̤j����
    DATE_LENGTH = 7;                  //���������
    TIME_LENGTH = 4;                  //�ɶ�������
    MATERIAL_NO_LENGTH = 15;          //�Ƹ����̤j����
    SEQ_NO_LENGTH = 15;               //�Ǹ����̤j����
    NO_AVAILABLE_QUERY_COMP = '-9: �L�k���o�A������reference.�H�P�L�kŪ����Ʈw�����.';
    SELECT_MODE = 1;
    IUD_MODE = 2;
    ERROR_TOTAL_WAREHOUSE_ID = 54088; //�P�w�ǤJ�r��O�_���¼Ʀr,�Y���O�h���󦹱`��
    XML_NO_DATA = 'EMPTY';
    ERROR_STRING = 'ERROR';
    //TOTAL_WAREHOUSE_ID = 0;

type

  TMssIniData = class(TObject)
  public
    nDataAreaNo : Integer;
    sDataArea : String;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
  end;
  

  TE2C = class(TAutoObject, IE2C)
  private
    function getRendomFileName:String;
    function VdCD043(const aCompCode, aModelNo: String): String;
    function VdCD022(const aCompCode, aDescription: String): String;
    function GetModelNo(const aCompCode, aFacisNo: String; var aModelNo: String): String;
    function GetDescription(const aType, aCompCode, aFacisNo: String; var aDescription: String): String;
  protected
    G_InCompCodeStrList : TStringList;
    G_DbConnAry : array of TADOConnection;
    G_AdoQueryAry : array of TADOQuery;
    sG_E2C_DLL1_RESULT : String;
    sG_E2C_DLL2_RESULT : String;
    sG_E2C_DLL3_RESULT : String;    
    G_Conn : TADOConnection;
    G_DbAreaNameStrList, G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList : TStringList;
    nG_TOTAL_WAREHOUSE_ID : Integer;
    G_StrList : TStringList;
    procedure TransTmpIniFile;
    procedure DeleteTmpIniFile;
    function RunSQL(aMode: Integer; aCompID, aSQL:String):TADOQuery;
    function getActiveQuery(sI_CompID : String):TADOQuery;
    procedure E2C_DLL1(vI_Data1, vI_Data2: OleVariant); safecall;
    procedure E2C_DLL2(vI_Data1, vI_Data2, vI_Data3: OleVariant); safecall;
//    procedure E2C_DLL3(vI_Data1: OleVariant); safecall;    
    { Protected declarations }
    procedure setDateTimeFormat;
    procedure createDbInfoStrList;
    procedure saveToFile(sI_Content:String; sI_FileName:String);
    function Get_E2C_DLL1_RESULT: WideString; safecall;
    procedure Set_E2C_DLL1_RESULT(const Value: WideString); safecall;
    function Get_E2C_DLL2_RESULT: WideString; safecall;
    procedure Set_E2C_DLL2_RESULT(const Value: WideString); safecall;

    function processXML1(sI_Xml:WideString):WideString;  //�B�z�J�w���
    function processXML2(sI_Xml:WideString):WideString;  //�B�z��h���
    function processXML3(sI_Xml:WideString):WideString;  //�B�z�ռ����
    function processXML4(sI_Xml:WideString):WideString;  //�B�z�t����
    function processXML5(sI_Xml:WideString):WideString;  //�B�z��h�Ƹ��    

    function getDbConnInfo(var I_DbAliasStrList, I_DbUserNameStrList, I_DbPasswordStrList : TStringList):WideString;
    function connectToDB(bI_ConnToWarehouse : boolean; sI_InCompCodeList : String) : String;

    function checkSeqNo1(sI_CompID, sI_EquipNo, sI_SeqNo:String;
             var sI_Msg:String; var bI_EquipExist:boolean ):boolean;  //�J�w
    function checkSeqNo2(sI_Function, sI_CompID, sI_EquipNo, sI_SeqNo:String;
             var sI_Msg:String; var bI_EquipExist:boolean ):boolean;  //��h

    function isValidDateStr(sI_Date:String):boolean;
    function isValidTimeStr(sI_Time:String):boolean;
    function getExexDateTime(sI_Date, sI_Time:String):TDate;
    function validCompCode(sI_CompCode : String;var sI_Msg : String) : Boolean;

    function allocateEquip(sI_Status,sI_OldCompCode,sI_InCompCode,sI_MaterialNo,
             sI_EquipType,sI_SeqNo,sI_EquipTableName,sI_CurDateTime : String;var sI_Msg : String) : Boolean;
    function queryOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo : String;
             var sI_BatchNo,sI_ExecTime,sI_Description,sI_Spec,sI_VendorCode,
             sI_VendorName,sI_ModeFlag, sI_ModelNo,
             sI_MaterialType, sI_HwVer, sI_CallbackMode, sI_Direction, sI_HdSize,
             sI_Msg : String) : Boolean;
    procedure insertEquipToCompany(sI_EquipTableName,sI_InCompCode,sI_BatchNo,
              sI_ExecTime,sI_MaterialNo,sI_SeqNo,sI_Description,sI_Spec,
              sI_VendorCode,sI_VendorName,sI_CurDateTime, sI_ModeFlag, sI_ModelNo,
              sI_MaterialType, sI_HwVer, sI_CallbackMode, sI_Direction, sI_HdSize: String);
    procedure deleteOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo : String);
    procedure logDataToSOAC0204(sI_CurDateTime,sI_CompCode,sI_FunctionName,sI_Para : String);
    function checkEquipIsExist(sI_CompCode,sI_EquipType,sI_SeqNo,sI_ErrMsg : String;var sI_Msg : String) : Boolean;
    procedure insertPairData(sI_STBFaciSNo,sI_IccFaciSNo : String);
    procedure deletePairData1(sI_STBFaciSNo,sI_IccFaciSNo : String);
    procedure deletePairData2(sI_FaciSNo : String);
    function getAreaDataNo(sI_CompCode : String) : String;
    procedure getIniData(sI_CompCode: String;var sI_DbPassword,sI_DbAlias,sI_DbUserName : String);
    function checkIsConnectCompCode(sI_CompCode : String) : Boolean;
    procedure allocateSpace(sI_CompID : String);
    procedure deallocateSpace(sI_CompID : String);
    procedure DbConnAryRollBack;
    function isNotWareHouseCode(sI_TableName,sI_SeqNo : String;var sI_CompCode : String) : String;
    procedure E2C_DLL3(vI_Data1: OleVariant); safecall;
    function Get_E2C_DLL3_RESULT: WideString; safecall;
    procedure Set_E2C_DLL3_RESULT(const Value: WideString); safecall;

  end;

implementation

uses ComServ, xmlU, UdateTimeu, Ustru, DevPower_TLB;

function TE2C.connectToDB(bI_ConnToWarehouse : boolean; sI_InCompCodeList : String) : String;
var
    aIndex: Integer;
    sL_DbConnStr,sL_CompCode : String;
    sL_DbPassword, sL_DbAlias, sL_DbUserName : String;

begin
    try
      G_InCompCodeStrList.Clear;

      if bI_ConnToWarehouse then //�Y�n connect ���`��
      begin
        if trim(sI_InCompCodeList) <> '' then
          sI_InCompCodeList := '0,' + sI_InCompCodeList
        else
          sI_InCompCodeList := '0';
      end
      else
      begin
        if trim(sI_InCompCodeList) <> '' then
          sI_InCompCodeList := sI_InCompCodeList
        else
        begin
          Result := 'Could not connect to db.';
          Exit;
        end;
      end;

      G_InCompCodeStrList := TUstr.ParseStrings(sI_InCompCodeList,',');

      setlength(G_DbConnAry, G_InCompCodeStrList.Count);
      setlength(G_AdoQueryAry, G_InCompCodeStrList.Count);

      for aIndex := 0 to G_InCompCodeStrList.Count -1 do
      begin
        sL_CompCode := G_InCompCodeStrList.Strings[aIndex];

        //���o�ǤJ���q����Ʈw�s�u���
        getIniData(sL_CompCode,sL_DbPassword,sL_DbAlias,sL_DbUserName);

        sL_DbConnStr :=
          'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbPassword +
          ';Persist Security Info=True;User ID='+sL_DbUserName+
          ';Data Source='+sL_DbAlias;

        G_DbConnAry[aIndex] := TADOConnection.Create(nil);

        if not G_DbConnAry[aIndex].Connected then
        begin
          G_DbConnAry[aIndex].ConnectionString := sL_DbConnStr;
          G_DbConnAry[aIndex].LoginPrompt := False;
          G_DbConnAry[aIndex].Connected := True;
          G_AdoQueryAry[aIndex] := TADOQuery.Create(nil);
          G_AdoQueryAry[aIndex].Connection := G_DbConnAry[aIndex];
          G_AdoQueryAry[aIndex].Tag := StrToInt(sL_CompCode); //howard 0214
        end;
      end;
    except
      Result := '-1: �PCC&B��Ʈw�s�u����, ��Ʈw�O�W=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';
      exit;
    end;
    Result :=EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TE2C.E2C_DLL1(vI_Data1, vI_Data2: OleVariant);
var
  aStrList : TStringList;
  sL_Data1, sL_Data2, aResult: String;
begin
  sL_Data1 := VarToStr( vI_Data1 );
  sL_Data2 := VarToStr( vI_Data2 );
  setDateTimeFormat;
  //�N�[�K��ini�ɮ׸ѱK
  TransTmpIniFile;
  aResult := EmptyStr;
  if ( UpperCase( trim( sL_Data1 ) ) <> XML_NO_DATA ) or
     ( UpperCase( trim( sL_Data2 ) ) <> XML_NO_DATA ) then
  begin
    createDbInfoStrList;
    // �q ini ��Ū���s�u�� DB �� information
    aResult := getDbConnInfo( G_DbAliasStrList, G_DbUserNameStrList,
      G_DbPasswordStrList );
    if ( aResult<> EmptyStr ) then
    begin
      Set_E2C_DLL1_RESULT( aResult );
      Exit;
    end;
    // �إ�DB�� connection
    aResult := connectToDB( CONN_TO_WAREHOUSE, EmptyStr );
    if ( aResult <> '' ) then
    begin
      Set_E2C_DLL1_RESULT( aResult );
      Exit;
    end;
  end;

  aStrList := TStringList.Create;
  try
    //�B�z�J�w..
    if UpperCase( trim( sL_Data1 ) ) <> XML_NO_DATA then
    begin
      aResult := processXML1( sL_Data1 );
      if ( aResult <> EmptyStr ) then
      begin
        DeleteTmpIniFile;
        Set_E2C_DLL1_RESULT( aResult );
        exit;
      end;
    end;

    //down, �B�z��h..
    if UpperCase(trim(sL_Data2))<>XML_NO_DATA then
    begin
      aResult := processXML2(sL_Data2);
      if ( aResult <> EmptyStr ) then
      begin
        DeleteTmpIniFile;
        Set_E2C_DLL1_RESULT( aResult );
        Exit;
      end;
    end;
  finally
    aStrList.Free;
  end;
  //�R���ѱK�᪺ini�ɮ�
  DeleteTmpIniFile;
  G_StrList.Free;
  Set_E2C_DLL1_RESULT( aResult );
end;

{ ---------------------------------------------------------------------------- }

procedure TE2C.E2C_DLL2(vI_Data1, vI_Data2, vI_Data3: OleVariant);
var
    sL_Data1, sL_Data2, sL_Data3, aResult: String;
begin
    sL_Data1 := VarToStr( vI_Data1 );
    sL_Data2 := VarToStr( vI_Data2 );
    sL_Data3 := VarToStr( vI_Data3 );

    setDateTimeFormat;
    //�N�[�K��ini�ɮ׸ѱK
    TransTmpIniFile;
    aResult := EmptyStr;
    if ((UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (trim(sL_Data1)<>''))
    or  ((UpperCase(trim(sL_Data2))<>XML_NO_DATA) and (trim(sL_Data2)<>''))then
    begin
      createDbInfoStrList;
      //down, �q ini ��Ū���s�u�� DB �� information
      aResult := getDbConnInfo(G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList);
      if ( aResult <> EmptyStr ) then
      begin
        Set_E2C_DLL2_RESULT(aResult);
        exit;
      end;
      //up, �q ini ��Ū���s�u�� DB �� information
      aResult := connectToDB(CONN_TO_WAREHOUSE, sL_Data3);
      if ( aResult <> EmptyStr ) then
      begin
        Set_E2C_DLL2_RESULT(aResult);
        exit;
      end;
    end;
    //�B�z�ռ�..
    if (UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (trim(sL_Data1)<>'') then
    begin
      aResult := processXML3(sL_Data1);
      if aResult<>'' then
      begin
        DeleteTmpIniFile;
        Set_E2C_DLL2_RESULT(aResult);
        exit;
      end;
    end;
    //down, �B�z�t��..
    {
    if UpperCase(trim(vI_Data2))<>XML_NO_DATA then
    begin
      sL_Result := processXML4(vI_Data2);
      if sL_Result<>'' then
      begin
        DeleteTmpIniFile;
        Set_E2C_DLL2_RESULT(sL_Result);
        exit;
      end;
    end;
    }
    //�R���ѱK�᪺ini�ɮ�
    DeleteTmpIniFile;
    G_StrList.Free;
    Set_E2C_DLL2_RESULT(aResult);
end;

{ ---------------------------------------------------------------------------- }

procedure TE2C.saveToFile(sI_Content: String; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    if sI_Content='' then
    begin
      L_StrList.Add('xxxxxxxxx');
      L_StrList.SaveToFile('c:\a5.txt');
      L_StrList.Clear;
    end
    else
    begin
      L_StrList.Add('yyyyyyyyyyyy');
      L_StrList.SaveToFile('c:\a6.txt');
      L_StrList.Clear;
    end;
    L_StrList.Add(sI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;


function TE2C.Get_E2C_DLL1_RESULT: WideString;
begin
    result := sG_E2C_DLL1_RESULT;
end;

procedure TE2C.Set_E2C_DLL1_RESULT(const Value: WideString);
begin
    sG_E2C_DLL1_RESULT := value;
end;

function TE2C.Get_E2C_DLL2_RESULT: WideString;
begin
    result := sG_E2C_DLL2_RESULT;
end;

procedure TE2C.Set_E2C_DLL2_RESULT(const Value: WideString);
begin
    sG_E2C_DLL2_RESULT := Value;
end;



//�B�z�J�w���
function TE2C.processXML1(sI_Xml: WideString):WideString;
var
    L_XMLDoc : IXMLDOMDocument;
    L_NodeList, L_MaterialNodeList, L_SeqNodeList : IXMLDOMNodeList;
    L_BatchNoNode, L_MaterialNode, L_SeqNode : IXMLDOMNODE;
    ii, jj, kk : Integer;

    sL_TmpXmlFileName, sL_EquipTableName,sL_Result : String;
    sL_BatchNo, sL_Date, sL_Time, sL_CurDateTime : String;
    sL_MakerID, sL_MakerName, sL_Status, sL_CompID : String;
    sL_MaterialNo,sL_Quantity, sL_ItemName, sL_Spec, sL_Equip : String;
    sL_SeqNo, sL_Msg, sL_SQL, sL_UpdEn ,sL_CheckCompCode : String;
    nL_Seed, nL_Quantity  : Integer;
    bL_EquipExist : boolean;
    dL_ExecDateTime : TDate;
    sL_Error,sL_ReturnCompCode : String;
    aModeFlag, aModelNo: String;
    aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize: String;
begin

    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��
    allocateSpace(IntToStr(nG_TOTAL_WAREHOUSE_ID));

    try
      //�N�{���]�b Transaction �����
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].BeginTrans;


      L_XMLDoc := CoDOMDocument.Create;



      L_XMLDoc.async := False;//�����J������,�{���~�|���U����
      L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�


      //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

      sL_UpdEn := '���ƨt��';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        //down, �ˬd���q�N�X�O�_���`��
        L_NodeList := L_XMLDoc.selectNodes('//���q�N�X');
        if L_NodeList<> nil then
        begin
          for ii:=0 to L_NodeList.length -1 do
          begin
            sL_CheckCompCode := L_NodeList.item[ii].text;
            //showmessage('�`�ܥN�X:' + IntToStr(nG_TOTAL_WAREHOUSE_ID) + '   ���q�N�X:' + sL_CheckCompCode);
            if sL_CheckCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-5: �J�w���q�N�X�����`�ܥN�X<' + IntToStr(nG_TOTAL_WAREHOUSE_ID) + '>,���~���q�N�X=<'+sL_CheckCompCode + '>';
            Exit;
            end;
          end;
        end;
        //up, �ˬd���q�N�X�O�_���`��



        L_NodeList := L_XMLDoc.getElementsByTagName('�帹');     //���o��󤤤��Ҧ��帹 tag
        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// �帹
          //down, �ˬd�帹�榡
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: �帹���׿��~, �帹=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, �ˬd�帹�榡

          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '���');
          //down, �ˬd����榡
          if length(sL_Date)<>DATE_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: ������׿��~, ���=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: �������榡����: <' + sL_Date + '>';
            exit;
          end;
          //up, �ˬd����榡

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '�ɶ�');
          //down, �ˬd�ɶ��榡
          if length(sL_Time)<>TIME_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: �ɶ����׿��~, �ɶ�=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: ����ɶ��榡����: <' + sL_Time + '>';
            exit;
          end;
          //up, �ˬd�ɶ��榡

          sL_MakerID := TMyXML.getSingleNodeText(L_BatchNoNode, '�t�ӥN�X');
          if length(sL_MakerID)>VENDORCODE_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: �t�ӥN�X�����׳̦h ' + IntToStr(VENDORCODE_LENGTH) + ' �X';
            exit;
          end;

          sL_MakerName := TMyXML.getSingleNodeText(L_BatchNoNode, '�t�ӦW��');
          if length(sL_MakerName)>VENDORNAME_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: �t�ӦW�٪����׳̦h ' + IntToStr(VENDORNAME_LENGTH) + ' �X';
            exit;
          end;


          // �ڥ��S�Ψ�, ����, ������ xml ���Ǥ] ok
          //sL_Status := TMyXML.getSingleNodeText(L_BatchNoNode, '���p');
          sL_Status := '1';
          //down, �ˬd�i�X(�J�w/��h)���p�榡 => ���ѼƫO�d,�@�w�n�ǤJ 1, �d�ݤ���X�R
          if not (StrToInt(sL_Status) in [1]) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: ���帹���i�X���p����, �帹=<'+sL_BatchNo+'>, �i�X���p=<'+sL_Status+'>';
            exit;
          end;
          //up, �ˬd�i�X(�J�w/��h)���p�榡 => ���ѼƫO�d,�@�w�n�ǤJ 1, �d�ݤ���X�R

          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '���q�N�X');


          L_MaterialNodeList := L_BatchNoNode.selectNodes('./�Ƹ�');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// �Ƹ�
            //down, �ˬd�Ƹ��榡
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, �ˬd�Ƹ��榡

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'�ƶq');
            //down, �ˬd�ƶq�榡
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: �ƶq���~, �ƶq=<'+sL_Quantity+'>';
              exit;
            end;
            //up, �ˬd�ƶq�榡

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode, '�ȪA�~�W' );
            if ( Length( sL_ItemName ) > 20 ) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: �ȪA�~�W���׳̦h20�X, �Ƹ�=<'+sL_MaterialNo+'>, �ȪA�~�W=<'+sL_ItemName+'>';
              Exit;
            end;


            sL_Spec := TMyXML.getAttributeValue(L_MaterialNode,'�W��');

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'�]��');
            //down, �ˬd�]�Ʈ榡
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: ���Ƹ����]�ƧO����, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+sL_Equip+'>';
              exit;
            end;
            //up, �ˬd�]�Ʈ榡

            { �� STB ��, �~���� Flag }
            
            aModeFlag := EmptyStr;
            aModelNo := EmptyStr;
            aMaterialType := EmptyStr;
            aHwVer := EmptyStr;
            aCallbackMode := EmptyStr;
            aDirection := EmptyStr;
            aHdSize := EmptyStr;

            if ( sL_Equip = '1' ) then
            begin
              aModeFlag := TMyXML.getAttributeValue(L_MaterialNode,'����' );
              if ( Trim( aModeFlag ) <> '0' ) and ( Trim( aModeFlag ) <> '1'  ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: ���Ƹ�����������, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aModelNo := TMyXML.getAttributeValue(L_MaterialNode,'����' );
              if ( Trim( aModelNo ) = EmptyStr ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: ���Ƹ������ج��ŭ�, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              if ( Length( aModelNo ) > 20 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: ���ئW�٪��׳̦h20�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aMaterialType := TMyXML.getAttributeValue( L_MaterialNode, '���ƫ���' );
              if ( Length( aMaterialType ) > 50 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: ���ƫ������׳̦h50�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aHwVer := TMyXML.getAttributeValue( L_MaterialNode, '�w�骩��' );
              if ( Length( aHwVer ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: �w�骩�����׳̦h10�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aCallbackMode := TMyXML.getAttributeValue( L_MaterialNode, '�^�ǼҦ�' );
              if ( Length( aCallbackMode ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: �^�ǼҦ����׳̦h10�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aDirection := TMyXML.getAttributeValue( L_MaterialNode, '�����V' );
              if ( Length( aDirection ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: �����V���׳̦h10�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aHdSize := TMyXML.getAttributeValue( L_MaterialNode, '�w�Юe�q' );
              if ( Length( aHdSize ) > 20 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: �w�Юe�q���׳̦h20�X, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
            end;

            L_SeqNodeList := L_MaterialNode.selectNodes('./�Ǹ�');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
              Result := '-2: ���Ƹ����ƶq�P���ӵ��Ƥ��X, �Ƹ�=<'+sL_MaterialNo+'>, �ƶq=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
            case StrToInt(sL_Equip) of
              1: sL_EquipTableName := 'SOAC0201A'; //�� STB
              2: sL_EquipTableName := 'SOAC0201B'; //�� ICC
            end;


            for kk:=0 to L_SeqNodeList.length -1 do
            begin
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, �ˬd�Ǹ��榡
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                Result := '-2: �Ǹ����׿��~, �Ǹ�=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, �ˬd�Ǹ��榡


              //�ˬd�ӳ]�ƧǸ��O�_�s�b���`��,�B���q�O�������`�ܥN�X
              sL_Error := isNotWareHouseCode(sL_EquipTableName,sL_SeqNo,sL_ReturnCompCode);
              if sL_Error <> '' then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                if sL_Error = ERROR_STRING then
                  Result := '-2: ���Ǹ�=<' + sL_SeqNo + '>�]�Ƥw�ռ��ܤ��q�O=<' + sL_ReturnCompCode + '> ����A�J�w'
                else
                  Result := sL_Error;
                exit;
              end;


              //down, �ˬd�Ǹ��O�_�X�z
              if not checkSeqNo1(sL_CompID, sL_Equip, sL_SeqNo,sL_Msg, bL_EquipExist) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                Result := sL_Msg;
                exit;
              end;
              //up, �ˬd�Ǹ��O�_�X�z
              //down, ��sCC&B�������]�Ƹ����
              //howard..
              if bL_EquipExist then
              begin
                sL_SQL := ' UPDATE '+ sL_EquipTableName +
                          '    SET COMPCODE=' + sL_CompID +',BATCHNO=' + STR_SEP + sL_BatchNo + STR_SEP +',';
                { STB Flag }
                if ( sL_Equip = '1' ) then
                begin
                  sL_SQL := sL_SQL + ' MODEFLAG=' + STR_SEP + aModeFlag + STR_SEP +',';
                  sL_SQL := sL_SQL + ' MODELNO=' + STR_SEP + aModelNo + STR_SEP +',';
                  sL_SQL := sL_SQL + ' MATERIALTYPE=' + STR_SEP + aMaterialType + STR_SEP +',';
                  sL_SQL := sL_SQL + ' HWVER=' + STR_SEP + aHwVer + STR_SEP +',';
                  sL_SQL := sL_SQL + ' CALLBACKMODE=' + STR_SEP + aCallbackMode + STR_SEP +',';
                  sL_SQL := sL_SQL + ' DIRECTION=' + STR_SEP + aDirection + STR_SEP +',';
                  sL_SQL := sL_SQL + ' HDSIZE=' + STR_SEP + aHdSize + STR_SEP +',';
                end;
                sL_SQL := sL_SQL + ' EXECTIME=' + TUstr.getOracleSQLDateTimeStr(dL_ExecDateTime) + ',';
                sL_SQL := sL_SQL + ' MATERIALNO='+STR_SEP + sL_MaterialNo + STR_SEP +',';
                sL_SQL := sL_SQL + ' DESCRIPTION='+ STR_SEP + sL_ItemName + STR_SEP +',';
                sL_SQL := sL_SQL + ' SPEC='+ STR_SEP + sL_Spec + STR_SEP +',';
                sL_SQL := sL_SQL + ' VENDORCODE='+ STR_SEP + sL_MakerID + STR_SEP +',';
                sL_SQL := sL_SQL + ' VENDORNAME='+ STR_SEP + sL_MakerName + STR_SEP +',';
                sL_SQL := sL_SQL + ' UPDTIME='+ STR_SEP + sL_CurDateTime + STR_SEP +',';
                sL_SQL := sL_SQL + ' UPDEN='+ STR_SEP + sL_UpdEn + STR_SEP ;
                sL_SQL := sL_SQL + ' WHERE FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              end
              else
              begin
                { �� = 1 ��(STB), �h7����� MODEFLAG, MODELNO, MaterialType, HwVer, CallbackMode, Direction, HdSize }
                if ( sL_Equip = '1' ) then
                begin
                  sL_SQL :=
                    ' INSERT INTO SOAC0201A ( ' +
                    ' COMPCODE, BATCHNO, EXECTIME, MATERIALNO, FACISNO, DESCRIPTION,' +
                    ' SPEC, VENDORCODE, VENDORNAME, UPDTIME, UPDEN, USEFLAG,'+
                    ' MODEFLAG,MODELNO, MATERIALTYPE, HWVER, CALLBACKMODE, DIRECTION, HDSIZE )';
                  sL_SQL := sL_SQL +' VALUES( ' +sL_CompID+',' + STR_SEP + sL_BatchNo + STR_SEP +',' + TUstr.getOracleSQLDateTimeStr(dL_ExecDateTime) + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_MaterialNo + STR_SEP +',' +STR_SEP + sL_SeqNo + STR_SEP +',';
                  sL_SQL := sL_SQL + STR_SEP + sL_ItemName + STR_SEP + ',' + STR_SEP + sL_Spec + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_MakerID + STR_SEP + ',' + STR_SEP + sL_MakerName + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_CurDateTime + STR_SEP + ',' + STR_SEP + sL_UpdEn + STR_SEP + ',' + 'NULL' + ',';
                  sL_SQL := sL_SQL + STR_SEP + aModeFlag + STR_SEP + ',' + STR_SEP + aModelNo + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + aMaterialType + STR_SEP + ',' + STR_SEP + aHwVer + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + aCallbackMode + STR_SEP + ',' + STR_SEP + aDirection + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + aHdSize + STR_SEP + ')';
                end else
                begin
                  sL_SQL :='insert into '+sL_EquipTableName +'(COMPCODE,BATCHNO,EXECTIME,MATERIALNO,FACISNO,DESCRIPTION,SPEC,VENDORCODE,VENDORNAME,UPDTIME,UPDEN,USEFLAG)';
                  sL_SQL := sL_SQL +' values(' +sL_CompID+',' + STR_SEP + sL_BatchNo + STR_SEP +',' + TUstr.getOracleSQLDateTimeStr(dL_ExecDateTime) + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_MaterialNo + STR_SEP +',' +STR_SEP + sL_SeqNo + STR_SEP +',';
                  sL_SQL := sL_SQL + STR_SEP + sL_ItemName + STR_SEP + ',' + STR_SEP + sL_Spec + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_MakerID + STR_SEP + ',' + STR_SEP + sL_MakerName + STR_SEP + ',';
                  sL_SQL := sL_SQL + STR_SEP + sL_CurDateTime + STR_SEP + ',' + STR_SEP + sL_UpdEn + STR_SEP + ',' + 'NULL' + ')';
                end;
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              end;

              //up, ��sCC&B�������]�Ƹ����
            end;
          end;
        end;

        Result := '';
      end
      else
        Result := '-1:�ǤJ���J�w XML �榡���~!���˵�: ' + sL_TmpXmlFileName;


      //�Y���椤�S�����D,�hCommit
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].CommitTrans;

    except
      //�Y���椤�����D,�hRollback
      on E: Exception do
      begin
        Set_E2C_DLL1_RESULT('ErrorMsg:' + E.Message);
        G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
      end;
    end;


    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��
    deallocateSpace(IntToStr(nG_TOTAL_WAREHOUSE_ID));
end;


//�B�z�t����
function TE2C.processXML4(sI_Xml: WideString): WideString;
var
    L_XMLDoc : IXMLDOMDocument;
    L_NodeList : IXMLDOMNodeList;
    L_PairNode : IXMLDOMNODE;
    sL_TmpXmlFileName,sL_Para,sL_UpdEn,sL_CurDateTime,sL_Msg,sL_CompCode : String;
    sL_STBMaterailNo,sL_STBFaciSNo,sL_ICCMaterailNo,sL_ICCFaciSNo : String;
    sL_ErrStr1,sL_ErrStr2 : String;

    nL_Seed,ii : Integer;
begin
    L_XMLDoc := CoDOMDocument.Create;

    L_XMLDoc.async := False;//�����J������,�{���~�|���U����
    L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�

    //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
    sL_TmpXmlFileName := getRendomFileName;
    saveToFile(sI_Xml, sL_TmpXmlFileName);

    sL_Para := sL_TmpXmlFileName + ';';
    //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

    sL_UpdEn := '���ƨt��';
    sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

    if L_XMLDoc.load(sL_TmpXmlFileName) then
    begin

      L_NodeList := L_XMLDoc.getElementsByTagName('�t��');  //���o��󤤤��Ҧ��t�� tag

      for ii := 0 to L_NodeList.length-1 do
      begin
        L_PairNode := L_NodeList.item[ii];


        sL_STBMaterailNo := L_PairNode.attributes.item[0].text;
        sL_STBFaciSNo := L_PairNode.attributes.item[1].text;
        sL_ICCMaterailNo := L_PairNode.attributes.item[2].text;
        sL_ICCFaciSNo := L_PairNode.attributes.item[3].text;

        //down, �ˬd�Ƹ��榡
        if not (length(sL_STBMaterailNo)<=MATERIAL_NO_LENGTH) then
        begin
          Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_STBMaterailNo + '>';
          exit;
        end
        else if (not (length(sL_ICCMaterailNo)<=MATERIAL_NO_LENGTH))then
        begin
          Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_ICCMaterailNo + '>';
          exit;
        end;
        //up, �ˬd�Ƹ��榡

        //�ˬdSTB�]�ƧǸ��O�_�s�b
        sL_ErrStr1 := '-4: CC&B��Ʈw�L��STB�Ǹ�, �L�k�t��, �Ǹ�=<'+sL_STBFaciSNo+'>';
        if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),'1',sL_STBFaciSNo,sL_ErrStr1,sL_Msg) then
        begin
          Result := sL_Msg;
          exit;
        end;

        //�ˬdICC�]�ƧǸ��O�_�s�b
        sL_ErrStr2 := '-4: CC&B��Ʈw�L��ICC�Ǹ�, �L�k�t��, �Ǹ�=<'+sL_ICCFaciSNo+'>';
        if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),'2',sL_ICCFaciSNo,sL_ErrStr2,sL_Msg) then
        begin
          Result := sL_Msg;
          exit;
        end;

        //sL_CompCode := '';  //SOAC0201C �Ȥ��x�s���q�N�X
        //���R���w�s�b���ӧǸ�,�A�s�W,�H�T�O���
        deletePairData1(sL_STBFaciSNo,sL_IccFaciSNo);
        insertPairData(sL_STBFaciSNo,sL_IccFaciSNo);

      end;

      //Log�@���ܦ@�P��ưϪF�IMSS�I�s�O����(SOAC0204)
      //SOAC0201C �Ȥ��x�s���q�N�X
      sL_Para := '�t���� : ' + sL_Para;
      logDataToSOAC0204(sL_CurDateTime,'','E2C_DLL2',sL_Para);
      Result := '';

    end
    else
      Result := '-1:�ǤJ���t�� XML �榡���~!���˵�: ' + sL_TmpXmlFileName;


end;


//�B�z�ռ����
function TE2C.processXML3(sI_Xml: WideString): WideString;
var
  L_XMLDoc : IXMLDOMDocument;
  L_NodeList,L_SeqNodeList : IXMLDOMNodeList;
  L_MaterialNoNode,L_SeqNode : IXMLDOMNODE;

  sL_TmpXmlFileName,sL_UpdEn,sL_CurDateTime,sL_EquipTableName,sL_ErrStr1 : String;
  sL_MaterialNo,sL_EquipType,sL_InCompCode,sL_OutCompCode,sL_SeqNo,sL_Msg : String;
  sL_SQL,sL_Function,sL_Status,sL_Para,sL_Result : String;
  nL_Seed,ii,kk,nL_AreaDataNo: Integer;
  L_AdoQry : TADOQuery;
  bL_EquipExist,bL_IsConnectCompCode : Boolean;
  aModelNo, aDescription, aMsg: String;
begin

    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��
    for ii:=0 to G_InCompCodeStrList.Count-1 do
      allocateSpace(G_InCompCodeStrList.Strings[ii]);

    try
      //�N�{���]�b Transaction �����
      for ii:=0 to High(G_DbConnAry) do
      begin
        G_DbConnAry[ii].BeginTrans;
      end;

      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//�����J������,�{���~�|���U����
      L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�

      //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);

      sL_Para := sL_TmpXmlFileName + ';';
      //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

      sL_UpdEn := '���ƨt��';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        L_NodeList := L_XMLDoc.getElementsByTagName('�Ƹ�');  //���o��󤤤��Ҧ��Ƹ� tag
        //�ˬdXML���n�ռ������q�O,�O�_�����s�b���s�u����Ʈw
        for ii := 0 to L_NodeList.length-1 do
        begin
          L_MaterialNoNode := L_NodeList.item[ii];
          sL_MaterialNo :=  L_MaterialNoNode.childNodes.item[0].text;
          sL_InCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'�դJ���q');


          bL_IsConnectCompCode := checkIsConnectCompCode(sL_InCompCode);
          if not bL_IsConnectCompCode then
          begin
            //�Y���椤�����D,�hRollback
            DbConnAryRollBack;

            Result := '-3: , �n�դJ�����q�O,���]�t�b���s�u�����q��' + sL_InCompCode;
            exit;
          end;


          sL_OutCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'�եX���q');
          bL_IsConnectCompCode := checkIsConnectCompCode(sL_OutCompCode);
          if not bL_IsConnectCompCode then
          begin
            //�Y���椤�����D,�hRollback
            DbConnAryRollBack;

            Result := '-3: , �n�եX�����q�O,���]�t�b���s�u�����q��' + sL_OutCompCode;
            exit;
          end;
        end;


        for ii := 0 to L_NodeList.length-1 do
        begin

          L_MaterialNoNode := L_NodeList.item[ii];
          sL_MaterialNo :=  L_MaterialNoNode.childNodes.item[0].text;

          sL_EquipType := TMyXML.getAttributeValue(L_MaterialNoNode,'�]�ƧO');
          sL_InCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'�դJ���q');
          sL_OutCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'�եX���q');

          //down, �ˬd�Ƹ��榡
          if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
          begin
            //�Y���椤�����D,�hRollback
            DbConnAryRollBack;

            Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_MaterialNo + '>';
            exit;
          end;
          //up, �ˬd�Ƹ��榡

          //down, �ˬd���q�O�N�X
          if sL_InCompCode = '' then
          begin
            //�Y���椤�����D,�hRollback
            DbConnAryRollBack;

            Result := '-2: ���ռ����դJ���q�N�X���ର�ŭ�';
            exit;
          end;

          //�ˬd�դJ���q�N�X�O�_�s�b�� CD039 ��
          if sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
          begin
            if not validCompCode(sL_InCompCode,sL_Msg) then
            begin
              //�Y���椤�����D,�hRollback
              DbConnAryRollBack;

              Result := sL_Msg;
              exit;
            end;
          end;

          //UP, �ˬd���q�O�N�X


          //down, �ˬd�]�Ʈ榡
          if not (StrToInt(sL_EquipType) in [1,2]) then
          begin
            //�Y���椤�����D,�hRollback
            DbConnAryRollBack;

            Result := '-2: ���ռ����]�ƧO����, �]�ƧO=<'+sL_EquipType+'>';
            exit;
          end;
          //up, �ˬd�]�Ʈ榡
          if ii = 0 then
            sL_Para := sL_Para + '�]�ƧO:' +sL_EquipType + ',�դJ���q:' + sL_InCompCode + ',�Ƹ�:' + sL_MaterialNo
          else
            sL_Para := sL_Para + '-�]�ƧO:' + sL_EquipType + ',�դJ���q:' + sL_InCompCode + ',�Ƹ�:' + sL_MaterialNo;

          case StrToInt(sL_EquipType) of
            1: //�� STB
             begin
              sL_EquipTableName := 'SOAC0201A';
              sL_ErrStr1 := '-4: CC&B��Ʈw�L��STB�Ǹ�, �L�k�ռ� ,�Ǹ�=<'+sL_SeqNo+'>';
             end;
            2: //�� ICC
             begin
              sL_EquipTableName := 'SOAC0201B';
              sL_ErrStr1 := '-4: CC&B��Ʈw�L��ICC�Ǹ�, �L�k�ռ� ,�Ǹ�=<'+sL_SeqNo+'>';
             end;
          end;


          L_SeqNodeList := L_MaterialNoNode.selectNodes('./�Ǹ�');


          for kk:=0 to L_SeqNodeList.length -1 do
          begin
            L_SeqNode := L_SeqNodeList.item[kk];
            sL_SeqNo := L_SeqNode.text;

            //down, �ˬd�Ǹ��榡
            if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
            begin
              //�Y���椤�����D,�hRollback
              DbConnAryRollBack;

              Result := '-2: �Ǹ����׿��~, �Ǹ�=<'+sL_SeqNo + '>';
              exit;
            end;
            //up, �ˬd�Ǹ��榡

                //down, ����ݤ��q�O�ˬd�Ǹ��O�_�X�z
                sL_Function := '3';//�ռ�

                if not checkSeqNo2(sL_Function,sL_OutCompCode,sL_EquipType, sL_SeqNo,sL_Msg, bL_EquipExist) then
                begin
                  //�Y���椤�����D,�hRollback
                  DbConnAryRollBack;

                  Result := sL_Msg;
                  exit;
                end;
                //up, ����ݤ��q�O�ˬd�Ǹ��O�_�X�z
                
            { �ˬd�դJ�����q�O�O�_���إߦ��a�N�X, �ˮֽդJ�����q�O�Y�i }
            { 1.�����եX���q�O�� ModelNo }

            if ( sL_EquipType = '1' ) then
            begin
              aMsg := GetModelNo( sL_OutCompCode, sL_SeqNo, aModelNo );
              if ( aMsg <> EmptyStr ) then
              begin
                DbConnAryRollBack;
                Result := '-2:CC&B��Ʈw�����]�ƾ��ث������ŭ�, �Ƹ�=<'+sL_MaterialNo+'>';
                Exit;
              end;
            end;  

            aMsg := GetDescription( sL_EquipType, sL_OutCompCode, sL_SeqNo, aDescription );
            if ( aMsg <> EmptyStr ) then
            begin
              DbConnAryRollBack;
              Result := '-2:CC&B��Ʈw�����]�ƾ��ثȪA�~�W���ŭ�, �Ƹ�=<'+sL_MaterialNo+'>';
              Exit;
            end;
            

            if ( sL_InCompCode <> IntToStr( nG_TOTAL_WAREHOUSE_ID ) ) then
            begin
              { 3. �ˮֱ��դJ�����q�O�O�_���إߦ��N�X, �դJ���`�ܫh���ˮ� }
              if ( sL_EquipType = '1' ) then
              begin
                aMsg := VdCD043( sL_InCompCode , aModelNo );
                if ( aMsg <> EmptyStr ) then
                begin
                  DbConnAryRollBack;
                  Result := '-2:CC&B��Ʈw�����]�Ƭd�L�������ث���,����=<'+aModelNo+'> , �Ƹ�=<'+sL_MaterialNo+'>';
                  Exit;
                end;
              end;  
              { 4. �ˮֱ��դJ�����q�O�O�_���إߦ��ȪA�~�W, �դJ���`�ܫh���ˮ� }
              aMsg := VdCD022( sL_InCompCode , aDescription );
              if ( aMsg <> EmptyStr ) then
              begin
                DbConnAryRollBack;
                Result := '-2:CC&B��Ʈw�����]�Ƭd�L�����ȪA�~�W,�ȪA�~�W=<'+aDescription+'> , �Ƹ�=<'+sL_MaterialNo+'>';
                Exit;
              end;
            end;

            //�̤��P���p��s���
            //sL_Status
            if (sL_OutCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '0'; //�N��s�¤��q�O�ۦP,��������ʧ@
            end
            else if (sL_OutCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
               (sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '1'; //�`�ܳ]�ƽռ��ܤ��q�OA
            end
            else if (sL_OutCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '2'; //���q�OA�]�ƽռ��ܤ��q�OB
            end
            else if (sL_OutCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '3'; //���q�OA�]�ƽռ����`��
            end;


            if sL_Status <> '0' then   //�N��s�¤��q�O�ۦP,��������ʧ@
            begin
              //��s�ռ����

              allocateEquip(sL_Status,sL_OutCompCode,sL_InCompCode,sL_MaterialNo,
                            sL_EquipType,sL_SeqNo,sL_EquipTableName,sL_CurDateTime,sL_Msg);

            end;
          end;
        end;

        //Log�@���ܦ@�P��ưϪF�IMSS�I�s�O����(SOAC0204)
        logDataToSOAC0204(sL_CurDateTime,'','E2C_DLL2',sL_Para);

        Result := '';
      end
      else
        Result := '-1:�ǤJ���ռ� XML �榡���~!���˵�: ' + sL_TmpXmlFileName;


      //�Y���椤�S�����D,�hCommit
      for ii:=0 to High(G_DbConnAry) do
      begin
        G_DbConnAry[ii].CommitTrans;
      end;
      {
      for ii:=0 to G_InCompCodeStrList.Count-1 do
      begin
        nL_AreaDataNo := StrToInt(getAreaDataNo(G_InCompCodeStrList.Strings[ii]));
        G_DbConnAry[nL_AreaDataNo].CommitTrans;
      end;
      }

    except
      //�Y���椤�����D,�hRollback
      DbConnAryRollBack;

      {
      for ii:=0 to G_InCompCodeStrList.Count-1 do
      begin
        nL_AreaDataNo := StrToInt(getAreaDataNo(G_InCompCodeStrList.Strings[ii]));
        G_DbConnAry[nL_AreaDataNo].RollbackTrans;
      end;
      }
    end;


    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��
    for ii:=0 to G_InCompCodeStrList.Count-1 do
      deallocateSpace(G_InCompCodeStrList.Strings[ii]);

end;



//�B�z��h���
function TE2C.processXML2(sI_Xml: WideString): WideString;
var
  L_XMLDoc : IXMLDOMDocument;
  L_FirstNodeList,L_NodeList,L_SeqNodeList : IXMLDOMNodeList;
  L_FirstNode,L_EquipTypeNode,L_SeqNode : IXMLDOMNODE;
  ii,kk,nL_Seed : Integer;
  sL_TmpXmlFileName,sL_UpdEn,sL_CurDateTime,sL_EquipType,sL_Msg,sL_SeqNo : String;
  sL_EquipTableName,sL_TempTableName,sL_SQL,sL_ErrStr1,sL_Para,sL_Result : String;
  bL_EquipExist : Boolean;
  dL_ExecTime : TDate;

  sL_CompCode,sL_BatchNo,sL_OldBatchNo,sL_MaterialNo,sL_Function : String;
  sL_FaciSNo,sL_Description,sL_Spec,sL_VendorCode,sL_VendorName : String;
  L_AdoQry : TADOQuery;
  sL_Error,sL_ReturnCompCode : String;
  aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize: String;
begin
    try
      //�N�{���]�b Transaction �����
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].BeginTrans;

      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//�����J������,�{���~�|���U����
      L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�

      //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

      sL_UpdEn := '���ƨt��';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        L_FirstNodeList := L_XMLDoc.getElementsByTagName('��h');
        L_FirstNode := L_FirstNodeList.item[0];
        //���o��h�渹
        sL_BatchNo := TMyXML.getAttributeValue(L_FirstNode,'�渹');

        //���o��h���q�O
        sL_CompCode := TMyXML.getAttributeValue(L_FirstNode,'���q�N�X');
        //�u����h�b�`�ܪ��]��,�G�Y�ǤJ���q�O <> �`�ܥN�X�h������h
        if StrToInt(sL_CompCode) <> nG_TOTAL_WAREHOUSE_ID then
        begin
         if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
            G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

          Result := '-2: �ȯ���h�`�ܪ��]��,���~����h���q�O=< ' + sL_CompCode + ' >';
          exit;
        end;



        //down, �ˬd��h�渹�榡
        if sL_BatchNo = '' then
        begin
          if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
            G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

          Result := '-2: ��h�渹������';
          exit;
        end;
        //up, ��h�ˬd�渹�榡



        //�ˬd���q�N�X�O�_�s�b�� CD039 ��
        if sL_CompCode <> IntTOsTR(nG_TOTAL_WAREHOUSE_ID) then
        begin
          if not validCompCode(sL_CompCode,sL_Msg) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := sL_Msg;
            exit;
          end;
        end;
        //UP, �ˬd���q�O�N�X

  //�אּ�s��ini���Ҧ������q�O920927


        L_NodeList := L_XMLDoc.getElementsByTagName('�]��');     //���o��󤤤��Ҧ��]�� tag
        for ii:=0 to L_NodeList.length -1 do
        begin

          L_EquipTypeNode := L_NodeList.item[ii];
          sL_EquipType := L_EquipTypeNode.childNodes.item[0].text;// �]�ƧO

          //down, �ˬd�]�Ʈ榡
          if not (StrToInt(sL_EquipType) in [1,2]) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: ����h���]�ƧO����, �]�ƧO=<'+sL_EquipType+'>';
            exit;
          end;
          //up, �ˬd�]�Ʈ榡


          L_SeqNodeList := L_EquipTypeNode.selectNodes('./�Ǹ�');

          for kk:=0 to L_SeqNodeList.length -1 do
          begin
            L_SeqNode := L_SeqNodeList.item[kk];
            sL_SeqNo := L_SeqNode.text;


            //down, �ˬd�Ǹ��榡
            if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: ���Ǹ����׿��~, �Ǹ�=<'+sL_SeqNo + '>';
              exit;
            end;
            //up, �ˬd�Ǹ��榡

            case StrToInt(sL_EquipType) of
              1: //�� STB
               begin
                sL_EquipTableName := 'SOAC0201A';
                sL_TempTableName := 'SOAC0203A';
                sL_ErrStr1 := '-4: CC&B��Ʈw����STB���s�b, �L�k���P/��h, �Ǹ�=<'+sL_SeqNo+'>';
               end;
              2: //�� ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
                sL_TempTableName := 'SOAC0203B';
                sL_ErrStr1 := '-4: CC&B��Ʈw����ICC���s�b, �L�k���P/��h, �Ǹ�=<'+sL_SeqNo+'>';
               end;
            end;

            //�ˬd�ӳ]�ƧǸ��O�_�s�b���`��,�B���q�O�������`�ܥN�X
            sL_Error := isNotWareHouseCode(sL_EquipTableName,sL_SeqNo,sL_ReturnCompCode);
            if sL_Error <> '' then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              if sL_Error = ERROR_STRING then
                Result := '-2: ���Ǹ�=<' + sL_SeqNo + '>�]�Ƥw�ռ��ܤ��q�O=<' + sL_ReturnCompCode + '> ������h'
              else
                Result := sL_Error;
              exit;
            end;

            //��@�P��ư�Table�ˬd�ӳ]�ƧǸ��O�_�s�b down
            if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_EquipType,sL_SeqNo,sL_ErrStr1,sL_Msg) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := sL_Msg;
              exit;
            end
            else
            begin
              //�Y���ݤ��q�O�����`�ܤ��q�O,�h�ˬd���ݤ��q�O�ӳ]�ƬO�_�s�b
              if sL_CompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
              begin
                //down, �ˬd�Ǹ��O�_�X�z
                sL_Function := '2';//��h
                if not checkSeqNo2(sL_Function,sL_CompCode,sL_EquipType, sL_SeqNo,sL_Msg, bL_EquipExist) then
                begin
                  if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                    G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                  Result := sL_Msg;
                  exit;
                end;
                //up, �ˬd�Ǹ��O�_�X�z
              end;
            end;


            //down, �s�W���`�ܵ��P�Ƥ��ɤΧR��CC&B�������]�Ƹ����
            //���J�w����ɨ��X�ӧǸ������ ( SOAC0201A �� SOAC0201B )
            sL_SQL := 'SELECT * FROM ' + sL_EquipTableName + ' WHERE FACISNO=' +
                      STR_SEP + sL_SeqNo + STR_SEP;

            L_AdoQry := runSQL(SELECT_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);
            if L_AdoQry=nil then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + IntToStr(nG_TOTAL_WAREHOUSE_ID);
              Exit;
            end
            else
            begin
              dL_ExecTime := L_AdoQry.FieldByName('EXECTIME').AsDateTime;

              //��帹 = �J�w�帹
              sL_OldBatchNo := L_AdoQry.FieldByName('BATCHNO').AsString;

              sL_MaterialNo := L_AdoQry.FieldByName('MATERIALNO').AsString;
              sL_FaciSNo := L_AdoQry.FieldByName('FACISNO').AsString;
              sL_Description := L_AdoQry.FieldByName('DESCRIPTION').AsString;
              sL_Spec := L_AdoQry.FieldByName('SPEC').AsString;
              sL_VendorCode := L_AdoQry.FieldByName('VENDORCODE').AsString;
              sL_VendorName := L_AdoQry.FieldByName('VENDORNAME').AsString;
              {}
              aMaterialType := L_AdoQry.FieldByName( 'MATERIALTYPE' ).AsString;
              aHwVer := L_AdoQry.FieldByName( 'HWVER' ).AsString;
              aCallbackMode := L_AdoQry.FieldByName( 'CALLBACKMODE' ).AsString;
              aDirection := L_AdoQry.FieldByName( 'DIRECTION' ).AsString;
              aHdSize := L_AdoQry.FieldByName( 'HDSIZE' ).AsString;

              //Log�@����ƨ�@�P��ưϳ]�Ƶ��P�ƥ���
              if ( sL_EquipType = '1' ) then
              begin
                sL_SQL := 'INSERT INTO SOAC0203A ( COMPCODE, BATCHNO,    ' +
                          '  EXECTIME, OLDBATCHNO, MATERIALNO, FACISNO,  ' +
                          '  DESCRIPTION, SPEC, VENDORCODE, VENDORNAME,  ' +
                          '  UPDTIME, UPDEN,                             ' +
                          '  MATERIALTYPE, HWVER, CALLBACKMODE,          ' +
                          '  DIRECTION, HDSIZE )                         ' +
                          '  VALUES ( ' + sL_CompCode +
                          ',' + STR_SEP + sL_BatchNo + STR_SEP +
                          ',' + TUstr.getOracleSQLDateTimeStr(dL_ExecTime) +
                          ',' + STR_SEP + sL_OldBatchNo + STR_SEP +
                          ',' + STR_SEP + sL_MaterialNo + STR_SEP +
                          ',' + STR_SEP + sL_FaciSNo + STR_SEP +
                          ',' + STR_SEP + sL_Description + STR_SEP +
                          ',' + STR_SEP + sL_Spec + STR_SEP +
                          ',' + STR_SEP + sL_VendorCode + STR_SEP +
                          ',' + STR_SEP + sL_VendorName + STR_SEP +
                          ',' + STR_SEP + sL_CurDateTime + STR_SEP +
                          ',' + STR_SEP + '���ƨt��' + STR_SEP +
                          ',' + STR_SEP + aMaterialType + STR_SEP +
                          ',' + STR_SEP + aHwVer + STR_SEP +
                          ',' + STR_SEP + aCallbackMode + STR_SEP +
                          ',' + STR_SEP + aDirection + STR_SEP +
                          ',' + STR_SEP + aHdSize + STR_SEP + ')';
              end else
              begin
                sL_SQL := 'INSERT INTO SOAC0203B (COMPCODE,BATCHNO, ' +
                          'EXECTIME,OLDBATCHNO,MATERIALNO,FACISNO,DESCRIPTION,SPEC,' +
                          'VENDORCODE,VENDORNAME,UPDTIME,UPDEN) VALUES(' + sL_CompCode +
                          ',' + STR_SEP + sL_BatchNo + STR_SEP +
                          ',' + TUstr.getOracleSQLDateTimeStr(dL_ExecTime) +
                          ',' + STR_SEP + sL_OldBatchNo + STR_SEP +
                          ',' + STR_SEP + sL_MaterialNo + STR_SEP +
                          ',' + STR_SEP + sL_FaciSNo + STR_SEP +
                          ',' + STR_SEP + sL_Description + STR_SEP +
                          ',' + STR_SEP + sL_Spec + STR_SEP +
                          ',' + STR_SEP + sL_VendorCode + STR_SEP +
                          ',' + STR_SEP + sL_VendorName + STR_SEP +
                          ',' + STR_SEP + sL_CurDateTime + STR_SEP +
                          ',' + STR_SEP + '���ƨt��' + STR_SEP + ')';
              end;
              runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);
            end;

            //�R���@�P��ư�CC&B�������]�Ƹ����
            sL_SQL := 'DELETE ' + sL_EquipTableName + ' WHERE FACISNO=' +
                      STR_SEP + sL_SeqNo + STR_SEP;
            runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);

            //�Y�]�ƧǸ����ݤ��q�O�D�`�ܤ��q�O,�|���A�R���Ӥ��q�O��CC&B�������]�Ƹ����
            if sL_CompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
              runSQL(IUD_MODE,sL_CompCode,sL_SQL);

            //�R���@�P��ưϸӧǸ�������STB/ICC�t���� ( SOAC0201C )
            deletePairData2(sL_SeqNo);
          end;

        end;

        //Log�@���ܦ@�P��ưϪF�IMSS�I�s�O����(SOAC0204)
        sL_Para := '��h�渹=' + sL_BatchNo + ' ,���q�N�X=' + sL_CompCode;
        logDataToSOAC0204(sL_CurDateTime,sL_CompCode,'E2C_DLL2',sL_Para);
        Result := '';
      end
      else
        Result := '-1:�ǤJ����h XML �榡���~!���˵�: ' + sL_TmpXmlFileName;

      //�Y���椤�S�����D,�hCommit
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].CommitTrans;
    except
      //�Y���椤�����D,�hRollback
      if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
        G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
    end;
end;





function TE2C.getDbConnInfo(var I_DbAliasStrList, I_DbUserNameStrList, I_DbPasswordStrList : TStringList): WideString;
var
    sL_IniFileName,sL_CompCode : String;
    L_IniFile : TIniFile;
    ii, nL_TotalDataArea : Integer;
    sL_DbAlias, sL_DbUserName, sL_DbPassword, sL_Tag : String;
    sG_TOTAL_WAREHOUSE_ID,sL_ExeFileName,sL_ExePath : String;
    L_Obj : TMssIniData;
begin
    G_StrList := TStringList.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    //sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    sL_IniFileName := TMP_SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      result := '-1: Ū����Ʈw�s�u�Ѽ���<'+sL_IniFileName +'>����';
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);
    nL_TotalDataArea := L_IniFile.ReadInteger('SYSINFO','TOTALDATAAREASCOUNT',-1);
    if nL_TotalDataArea=-1 then
    begin
      if FileExists(sL_IniFileName) then
        DeleteFile(sL_IniFileName);
      result := '-1: ��Ʈw�s�u�Ѽ��ɤ���ư��`�Ƴ]�w���~!.���ˬd E2C_DLL.ini ���]�w.';
      exit;
    end;

    sG_TOTAL_WAREHOUSE_ID := L_IniFile.ReadString('SYSINFO','TOTAL_WAREHOUSE_ID','0');
    //�P�w�O�_���Ʀr
    if StrToIntDef(sG_TOTAL_WAREHOUSE_ID,ERROR_TOTAL_WAREHOUSE_ID) = ERROR_TOTAL_WAREHOUSE_ID then
    begin
      result := '-1: ��Ʈw�s�u�Ѽ��ɤ��`�ܥN�X�]�w���~!<'+sG_TOTAL_WAREHOUSE_ID +'>';
      exit;
    end;
    nG_TOTAL_WAREHOUSE_ID := StrToInt(sG_TOTAL_WAREHOUSE_ID);


    G_DbAliasStrList.Clear;
    G_DbUserNameStrList.Clear;
    G_DbPasswordStrList.Clear;
    try
      for ii:=0 to nL_TotalDataArea -1 do
      begin
        sL_Tag := DATA_AREA_HEADER + IntToStr(ii);
        sL_DbAlias := L_IniFile.ReadString(sL_Tag,'ALIAS','');
        sL_DbUserName := L_IniFile.ReadString(sL_Tag,'USERID','');
        sL_DbPassword := L_IniFile.ReadString(sL_Tag,'PASSWORD','');
        sL_CompCode := L_IniFile.ReadString(sL_Tag,'COMPCODE','');


        L_Obj := TMssIniData.Create;

        L_Obj.nDataAreaNo := ii;
        L_Obj.sDataArea := sL_Tag;
        L_Obj.sAlias := sL_DbAlias;
        L_Obj.sUserID := sL_DbUserName;
        L_Obj.sPassword := sL_DbPassword;
        L_Obj.sCompCode := sL_CompCode;

        G_DbAreaNameStrList.Add(sL_Tag);
        G_DbAliasStrList.Add(sL_DbAlias);
        G_DbUserNameStrList.Add(sL_DbUserName);
        G_DbPasswordStrList.Add(sL_DbPassword);

        G_StrList.AddObject(sL_CompCode, L_Obj);
      end;

      L_IniFile.Free;
    except
      result := '-1: ��Ʈw�s�u�Ѽ��ɤ���ưϳ]�w���~!<'+sL_IniFileName +'>';
      exit;
    end;
    Result := EmptyStr;
end;


function TE2C.isValidDateStr(sI_Date: String): boolean;
var
    sL_Year, sL_Month, sL_Day, sL_Date : String;
    nL_Year, nL_Month, nL_Day : Integer;
begin
    sL_Year := Copy(sI_Date,1,3);
    sL_Month := Copy(sI_Date,4,2);
    sL_Day := Copy(sI_Date,6,2);
    nL_Year := StrToInt(sL_Year) + 1911;
    nL_Month := StrToInt(sL_Month);
    nL_Day := StrToInt(sL_Day);
    sL_Date := Format('%.4d/%.2d/%.2d', [nL_Year, nL_Month, nL_Day]);
    if not TUdateTime.IsDateStr(sL_Date,'/') then
    begin
      result := false;
      exit;
    end;
      result := true;
end;

function TE2C.isValidTimeStr(sI_Time: String): boolean;
var
    nL_Hour, nL_Min : Integer;
    sL_Time : String;
begin
    nL_Hour := StrToInt(Copy(sI_Time, 1,2));
    nL_Min := StrToInt(Copy(sI_Time,3,2));

    sL_Time := Format('%.2d:%.2d:%.2d', [nL_Hour, nL_Min, 01]);
    if not TUdateTime.IsTimeStr(sL_Time,':') then
    begin
      result := false;
      exit;
    end;
      result := true;
end;

function TE2C.checkSeqNo1(sI_CompID, sI_EquipNo, sI_SeqNo: String;
  var sI_Msg: String; var bI_EquipExist:boolean ): boolean;
var
    sL_ErrStr : String;
    sL_EquipTableName, sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    sI_Msg := '';
    bI_EquipExist := false;
    case StrToInt(sI_EquipNo) of
      1: //�� STB
       begin
        sL_EquipTableName := 'SOAC0201A';
        sL_ErrStr := '-4: CC&B��Ʈw����STB�w�ϥΤ�, �L�k���s�J�w, �Ǹ�=<' + sI_SeqNo + '>'
       end;
      2: //�� ICC
       begin
        sL_EquipTableName := 'SOAC0201B';
        sL_ErrStr := '-4: CC&B��Ʈw����ICC�w�ϥΤ�, �L�k���s�J�w, �Ǹ�=<' + sI_SeqNo + '>';
       end;
    end;

    sL_SQL := 'select * from '+ sL_EquipTableName +' where  FACISNO='+ STR_SEP + sI_SeqNo + STR_SEP;
    //showmessage(sL_SQL);
    L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
    if L_AdoQry=nil then
    begin
      sI_Msg := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + sI_CompID;
      result := false;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount>0 then //��ܦ��]�Ƥw�g�s�b=>�ˬd�O�_���b�ϥΤ�
      begin
        bI_EquipExist := true;
        if (sI_CompID <> IntToStr(nG_TOTAL_WAREHOUSE_ID))then   //by Howard.. => �q�`�ܽեX�h����check
        begin
          sL_SQL := 'select PRDate from SO004 where FaciSNo=' + STR_SEP + sI_SeqNo+STR_SEP ;
          L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
          if L_AdoQry=nil then
          begin
            sI_Msg := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + sI_CompID;;
            result := false;
            Exit;
          end
          else
          begin
            if (L_AdoQry.RecordCount>0) and (L_AdoQry.FieldByName('PRDATE').AsString='') then  //�Y�L����,��ܦ��]�Ʃ|���Q�=>��ܦ��]�Ƥ��b�ϥΤ�
            begin
              L_AdoQry.Close;
              sI_Msg := sL_ErrStr;
              result := false;
              Exit;
            end;
          end;
        end;
      end;
      L_AdoQry.Close;
    end;
    result := true;
end;



function TE2C.checkSeqNo2(sI_Function,sI_CompID,sI_EquipNo, sI_SeqNo: String; var sI_Msg: String;
  var bI_EquipExist: boolean): boolean;
var
    sL_ErrStr1,sL_ErrStr2 : String;
    sL_EquipTableName, sL_SQL : String;
    L_AdoQry,L_AdoQry1 : TADOQuery;
begin
    //sI_Function = 2  �B�z��h
    //sI_Function = 3  �B�z�ռ�

    sI_Msg := '';
    bI_EquipExist := false;
    case StrToInt(sI_EquipNo) of
      1: //�� STB
       begin
         sL_EquipTableName := 'SOAC0201A';

         if sI_Function = '2' then
         begin
           sL_ErrStr1 := '-4: CC&B��Ʈw����STB���s�b, �L�k���P/��h, �Ǹ�=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B��Ʈw����STB�w�ϥΤ�, �L�k���P/��h, �Ǹ�=<'+sI_SeqNo+'>';
         end
         else if sI_Function = '3' then
         begin
           sL_ErrStr1 := '-4: CC&B��Ʈw�L��STB�Ǹ�,�Ǹ�=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B��Ʈw����STB�w�ϥΤ�, �L�k�ռ�, �Ǹ�=<'+sI_SeqNo+'>';
         end;
       end;
      2: //�� ICC
       begin
         sL_EquipTableName := 'SOAC0201B';
         if sI_Function = '2' then
         begin
           sL_ErrStr1 := '-4: CC&B��Ʈw����ICC���s�b, �L�k���P/��h, �Ǹ�=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B��Ʈw����ICC�w�ϥΤ�, �L�k���P/��h, �Ǹ�=<'+sI_SeqNo+'>';
         end
         else if sI_Function = '3' then
         begin
           sL_ErrStr1 := '-4: CC&B��Ʈw�L��ICC�Ǹ�,�Ǹ�=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B��Ʈw����ICC�w�ϥΤ�, �L�k�ռ�, �Ǹ�=<'+sI_SeqNo+'>';
         end;
       end;
    end;

    //sL_SQL := 'SELECT * FROM '+ sL_EquipTableName +' WHERE FACISNO='+ STR_SEP + sI_SeqNo + STR_SEP ;

    sL_SQL := Format( ' select * from %s where facisno = ''%s'' and compcode = ''%s''  ',
       [sL_EquipTableName, sI_SeqNo, sI_CompID] );

    L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
    if L_AdoQry=nil then
    begin
      sI_Msg := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + sI_CompID;;
      result := false;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount > 0 then //��ܦ��]�Ƥw�g�s�b=>�ˬd�O�_���b�ϥΤ�
      begin


//jackal0726
        bI_EquipExist := true;
        if (sI_CompID <> IntToStr(nG_TOTAL_WAREHOUSE_ID))then   //by Howard.. => �q�`�ܽեX�h����check
        begin
          sL_SQL := 'select PRDate from SO004 where FaciSNo=' + STR_SEP + sI_SeqNo+STR_SEP ;
          L_AdoQry1 := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
          if L_AdoQry1=nil then
          begin
            sI_Msg := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + sI_CompID;;
            result := false;
            Exit;
          end
          else
          begin
            if (L_AdoQry1.RecordCount>0) and (L_AdoQry1.FieldByName('PRDATE').AsString='') then  //�Y�L����,��ܦ��]�Ʃ|���Q�=>��ܦ��]�Ƥ��b�ϥΤ�
            begin
              L_AdoQry1.Close;
              sI_Msg := sL_ErrStr2;
              result := false;
              Exit;
            end;
          end;
        end;

      end
      else  //��ܦ��]�Ƥ��s�b
      begin
        L_AdoQry.Close;
        sI_Msg := sL_ErrStr1;
        result := false;
        Exit;
      end;
    end;

    L_AdoQry.Close;
    result := true;

end;

function TE2C.getExexDateTime(sI_Date, sI_Time: String): TDate;
var
    sL_Year, sL_Month, sL_Day, sL_Date : String;
    nL_Year, nL_Month, nL_Day : Integer;

    nL_Hour, nL_Min : Integer;
    sL_DateTime,sL_Time : String;

begin
    sL_Year := Copy(sI_Date,1,3);
    sL_Month := Copy(sI_Date,4,2);
    sL_Day := Copy(sI_Date,6,2);
    nL_Year := StrToInt(sL_Year) + 1911;
    nL_Month := StrToInt(sL_Month);
    nL_Day := StrToInt(sL_Day);
    sL_Date := Format('%.4d/%.2d/%.2d', [nL_Year, nL_Month, nL_Day]);

    nL_Hour := StrToInt(Copy(sI_Time, 1,2));
    nL_Min := StrToInt(Copy(sI_Time,3,2));
    sL_Time := Format('%.2d:%.2d:%.2d', [nL_Hour, nL_Min, 01]);

    sL_DateTime := sL_Date + ' ' + sL_Time;
    result := StrToDateTime(sL_DateTime);
    
end;





procedure TE2C.createDbInfoStrList;
begin
    G_DbAreaNameStrList := TStringList.Create;
    G_DbAliasStrList := TStringList.Create;
    G_DbUserNameStrList := TStringList.Create;
    G_DbPasswordStrList := TStringList.Create;
    G_InCompCodeStrList := TStringList.Create;
end;


function TE2C.getActiveQuery(sI_CompID: String): TADOQuery;
var
    L_AdoQry : TADOQuery;
    ii, nL_Pos : Integer;
    sL_AreaDataNo : String;
begin
    L_AdoQry := nil;
    for ii:=0 to High(G_AdoQueryAry) do
    begin
      if G_AdoQueryAry[ii].Tag = StrToInt(sI_CompID) then
      begin
        L_AdoQry := G_AdoQueryAry[ii];
      end;
    end;
    result := L_AdoQry;
end;


function TE2C.RunSQL(aMode: Integer; aCompID, aSQL: String): TADOQuery;
var
  aAdoQuery : TADOQuery;
begin
  try
    aAdoQuery := getActiveQuery( aCompID );
    if Assigned( aAdoQuery ) then
    begin
      aAdoQuery.SQL.Clear;
      aAdoQuery.SQL.Add( aSQL);
      case aMode of
        SELECT_MODE: aAdoQuery.Open;
        IUD_MODE: aAdoQuery.ExecSQL;
      end;
    end;
    Result := aAdoQuery;
  except
    {}
  end;
end;


function TE2C.validCompCode(sI_CompCode: String;var sI_Msg : String): Boolean;
var
  sL_SQL : String;
  L_AdoQry : TADOQuery;
begin
  sL_SQL := 'SELECT * FROM CD039 WHERE CODENO=' + sI_CompCode;

  L_AdoQry := runSQL(SELECT_MODE, sI_CompCode, sL_SQL);

  if L_AdoQry=nil then
  begin
    sI_Msg := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + sI_CompCode;
    Result := false;
  end
  else
  begin
    if L_AdoQry.RecordCount = 0 then
    begin
      sI_Msg := '-2: �����q�N�X���~, �N�X<' + sI_CompCode + '>';
      Result := false;
    end
    else
    begin
      sI_Msg := '';
      Result := true;
    end;
  end;
end;

function TE2C.allocateEquip(sI_Status, sI_OldCompCode, sI_InCompCode,
  sI_MaterialNo, sI_EquipType, sI_SeqNo,sI_EquipTableName,sI_CurDateTime : String;
  var sI_Msg: String): Boolean;
var
  sL_SQL,sL_Msg : String;
  sL_BatchNo,sL_ExecTime,sL_Description,sL_Spec,sL_VendorCode,sL_VendorName : String;
  bL_IsOk : Boolean;
  aModeFlag, aModelNo, aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize: String;
begin
  //���]�ƧO, �ק�@�P��ưϳ]�Ƹ����
  sL_SQL := 'UPDATE ' + sI_EquipTableName + ' SET COMPCODE=' + sI_InCompCode +
            ',MATERIALNO=' + STR_SEP + sI_MaterialNo + STR_SEP +
            ',UPDTIME=' + STR_SEP + sI_CurDateTime + STR_SEP +
            ',UPDEN=' + STR_SEP + '���ƨt��' + STR_SEP +
            ' WHERE FACISNO=' + STR_SEP + sI_SeqNo + STR_SEP;
            
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);


  if sI_Status = '1' then  //�`�ܳ]�ƽռ��ܤ��q�OA
  begin

    //���X��]�ƩҦ����
    bL_IsOk := queryOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo,
                                 sL_BatchNo,sL_ExecTime,sL_Description,sL_Spec,
                                 sL_VendorCode,sL_VendorName, aModeFlag, aModelNo,
                                 aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize,
                                 sL_Msg);
    if not bL_IsOk then
    begin
      sI_Msg := sL_Msg;
      Result := false;
      Exit;
    end;
    //���]�ƧO, �s�W�]�ƦܽդJ���q�O�]�Ƹ����
    insertEquipToCompany(sI_EquipTableName,sI_InCompCode,sL_BatchNo,sL_ExecTime,
                         sI_MaterialNo,sI_SeqNo,sL_Description,sL_Spec,
                         sL_VendorCode,sL_VendorName,sI_CurDateTime, aModeFlag, aModelNo,
                         aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize);
  end
  else if sI_Status = '2' then  //���q�OA�]�ƽռ��ܤ��q�OB
  begin
    //���X��]�ƩҦ����
    bL_IsOk := queryOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo,
                                 sL_BatchNo,sL_ExecTime,sL_Description,sL_Spec,
                                 sL_VendorCode,sL_VendorName, aModeFlag, aModelNo,
                                 aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize,
                                 sL_Msg);
    if not bL_IsOk then
    begin
      sI_Msg := sL_Msg;
      Result := false;
      Exit;
    end;

    //�R���]�ƭ즳���q�����
    deleteOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo);

    //���]�ƧO, �s�W�]�ƦܽդJ���q�O�]�Ƹ����
    insertEquipToCompany(sI_EquipTableName,sI_InCompCode,sL_BatchNo,sL_ExecTime,
                         sI_MaterialNo,sI_SeqNo,sL_Description,sL_Spec,
                         sL_VendorCode,sL_VendorName,sI_CurDateTime, aModeFlag, aModelNo,
                         aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize);


  end
  else if sI_Status = '3' then  //���q�OA�]�ƽռ����`��
  begin
    //�R���]�ƭ즳���q�����
    deleteOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo);
  end;

end;

function TE2C.queryOldEquipData(sI_EquipTableName,sI_OldCompCode, sI_SeqNo: String;
  var sI_BatchNo, sI_ExecTime, sI_Description, sI_Spec, sI_VendorCode,
  sI_VendorName, sI_ModeFlag, sI_ModelNo, 
  sI_MaterialType, sI_HwVer, sI_CallbackMode, sI_Direction, sI_HdSize,
  sI_Msg: String): Boolean;
var
  sL_SQL : String;
  L_AdoQry : TADOQuery;
begin
  sI_Msg := '';
  sL_SQL := 'SELECT * FROM ' + sI_EquipTableName + ' WHERE FACISNO=' +
            STR_SEP + sI_SeqNo + STR_SEP;


  L_AdoQry := runSQL(SELECT_MODE, sI_OldCompCode, sL_SQL);
  if L_AdoQry=nil then
  begin
    sI_Msg := NO_AVAILABLE_QUERY_COMP  + '  ���q�O=' + sI_OldCompCode;
    Result := false;
  end
  else
  begin
    sI_BatchNo := L_AdoQry.FieldByName('BATCHNO').AsString;
    sI_ExecTime := L_AdoQry.FieldByName('EXECTIME').AsString;
    sI_Description := L_AdoQry.FieldByName('DESCRIPTION').AsString;
    sI_Spec := L_AdoQry.FieldByName('SPEC').AsString;
    sI_VendorCode := L_AdoQry.FieldByName('VENDORCODE').AsString;
    sI_VendorName := L_AdoQry.FieldByName('VENDORNAME').AsString;
    if UpperCase( sI_EquipTableName ) = 'SOAC0201A' then
    begin
      sI_ModeFlag := L_AdoQry.FieldByName('MODEFLAG').AsString;
      sI_ModelNo :=  L_AdoQry.FieldByName('MODELNO').AsString;
      sI_MaterialType := L_AdoQry.FieldByName( 'MATERIALTYPE' ).AsString;
      sI_HwVer := L_AdoQry.FieldByName( 'HWVER' ).AsString;
      sI_CallbackMode := L_AdoQry.FieldByName( 'CALLBACKMODE' ).AsString;
      sI_Direction := L_AdoQry.FieldByName( 'DIRECTION' ).AsString;
      sI_HdSize := L_AdoQry.FieldByName( 'HDSIZE' ).AsString;
    end;
    Result := true;
  end;
end;

procedure TE2C.insertEquipToCompany(sI_EquipTableName, sI_InCompCode,
  sI_BatchNo, sI_ExecTime, sI_MaterialNo, sI_SeqNo, sI_Description,
  sI_Spec, sI_VendorCode, sI_VendorName, sI_CurDateTime, sI_ModeFlag, sI_ModelNo,
  sI_MaterialType, sI_HwVer, sI_CallbackMode, sI_Direction, sI_HdSize: String);
var
  sL_SQL : String;
begin
  if ( UpperCase( sI_EquipTableName ) = 'SOAC0201A' ) then
  begin
    sL_SQL := 'INSERT INTO SOAC0201A ( COMPCODE, BATCHNO, EXECTIME,      ' +
              '   MATERIALNO, FACISNO, DESCRIPTION, SPEC, VENDORCODE,    ' +
              '   VENDORNAME,UPDTIME, UPDEN, USEFLAG, MODEFLAG, MODELNO, ' +
              '   MATERIALTYPE, HWVER, CALLBACKMODE, DIRECTION, HDSIZE ) ' +
              ' VALUES( ' + sI_InCompCode + ',' + STR_SEP + sI_BatchNo + STR_SEP +
              ',' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_ExecTime)) +
              ',' + STR_SEP + sI_MaterialNo + STR_SEP +
              ',' + STR_SEP + sI_SeqNo + STR_SEP +
              ',' + STR_SEP + sI_Description + STR_SEP +
              ',' + STR_SEP + sI_Spec + STR_SEP +
              ',' + STR_SEP + sI_VendorCode + STR_SEP +
              ',' + STR_SEP + sI_VendorName + STR_SEP +
              ',' + STR_SEP + sI_CurDateTime + STR_SEP +
              ',' + STR_SEP + '���ƨt��' + STR_SEP +
              ',' + '0' +
              ',' + STR_SEP + sI_ModeFlag + STR_SEP +
              ',' + STR_SEP + sI_ModelNo + STR_SEP +
              ',' + STR_SEP + sI_MaterialType + STR_SEP +
              ',' + STR_SEP + sI_HwVer + STR_SEP +
              ',' + STR_SEP + sI_CallbackMode + STR_SEP +
              ',' + STR_SEP + sI_Direction + STR_SEP +
              ',' + STR_SEP + sI_HdSize + STR_SEP + ')';
  end else
  begin
    sL_SQL := 'INSERT INTO ' + sI_EquipTableName + '(COMPCODE,BATCHNO,EXECTIME,' +
              'MATERIALNO,FACISNO,DESCRIPTION,SPEC,VENDORCODE,VENDORNAME,UPDTIME,' +
              'UPDEN, USEFLAG) VALUES(' + sI_InCompCode + ',' + STR_SEP + sI_BatchNo + STR_SEP +
              ',' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_ExecTime)) +
              ',' + STR_SEP + sI_MaterialNo + STR_SEP +
              ',' + STR_SEP + sI_SeqNo + STR_SEP +
              ',' + STR_SEP + sI_Description + STR_SEP +
              ',' + STR_SEP + sI_Spec + STR_SEP +
              ',' + STR_SEP + sI_VendorCode + STR_SEP +
              ',' + STR_SEP + sI_VendorName + STR_SEP +
              ',' + STR_SEP + sI_CurDateTime + STR_SEP +
              ',' + STR_SEP + '���ƨt��' + STR_SEP + ',' + '0' + ')';
  end;

  runSQL(IUD_MODE,sI_InCompCode, sL_SQL);
end;

procedure TE2C.deleteOldEquipData(sI_EquipTableName, sI_OldCompCode,
  sI_SeqNo: String);
var
  sL_SQL : String;
begin
  sL_SQL := 'DELETE ' + sI_EquipTableName + ' WHERE FACISNO=' +
            STR_SEP + sI_SeqNo + STR_SEP;


  runSQL(IUD_MODE,sI_OldCompCode, sL_SQL);
end;

procedure TE2C.logDataToSOAC0204(sI_CurDateTime,sI_CompCode,sI_FunctionName,sI_Para : String);
var
  sL_SQL : String;
begin
  //Log�@���ܦ@�P��ưϪF�IMSS�I�s�O����(SOAC0204)
  sL_SQL := 'INSERT INTO SOAC0204(CALLTIME,COMPCODE,PROGNAME,PARA) VALUES(' +
            STR_SEP + sI_CurDateTime + STR_SEP + ',' +
            STR_SEP + sI_CompCode + STR_SEP + ',' +
            STR_SEP + sI_FunctionName + STR_SEP + ',' +
            STR_SEP + sI_Para + STR_SEP + ')';
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_SQL);
end;



function TE2C.checkEquipIsExist(sI_CompCode, sI_EquipType,
  sI_SeqNo,sI_ErrMsg: String; var sI_Msg: String): Boolean;
var
  sL_SQL,sL_EquipTableName : String;
  L_AdoQry : TADOQuery;
begin

  case StrToInt(sI_EquipType) of
    1: //�� STB
     begin
      sL_EquipTableName := 'SOAC0201A';
     end;
    2: //�� ICC
     begin
      sL_EquipTableName := 'SOAC0201B';
     end;
  end;


  //��@�P��ư�Table�ˬd�ӳ]�ƧǸ��O�_�s�b down
  sL_SQL :='SELECT * FROM ' + sL_EquipTableName +
           ' WHERE FACISNO=' + STR_SEP + sI_SeqNo + STR_SEP;
  L_AdoQry := runSQL(SELECT_MODE, sI_CompCode, sL_SQL);
  if L_AdoQry=nil then
  begin
    sI_Msg := NO_AVAILABLE_QUERY_COMP  + '  ���q�O=' + sI_CompCode;;
    Result := false;
    Exit;
  end
  else
  begin
    if L_AdoQry.RecordCount = 0 then
    begin
      sI_Msg := sI_ErrMsg;
      Result := false;
      exit;
    end
    else
    begin
      sI_Msg := '';
      Result := true;
      exit;
    end;
  end;

end;

procedure TE2C.insertPairData(sI_STBFaciSNo,sI_IccFaciSNo : String);
var
  sL_SQL : String;
begin
  sL_SQL := 'INSERT INTO SOAC0201C(STBSNO,SMARTCARDNO) VALUES(' +
             STR_SEP + sI_STBFaciSNo + STR_SEP +
            ',' + STR_SEP + sI_IccFaciSNo + STR_SEP + ')';
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_SQL);
end;

procedure TE2C.deletePairData1(sI_STBFaciSNo, sI_IccFaciSNo: String);
var
  sL_SQL : String;
begin
  sL_SQL := 'DELETE SOAC0201C WHERE STBSNO=' + STR_SEP + sI_STBFaciSNo + STR_SEP +
            ' OR SMARTCARDNO=' + STR_SEP + sI_IccFaciSNo + STR_SEP;
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_SQL);
end;

procedure TE2C.deletePairData2(sI_FaciSNo: String);
var
  sL_SQL : String;
begin
  sL_SQL := 'DELETE SOAC0201C WHERE STBSNO=' + STR_SEP + sI_FaciSNo + STR_SEP +
            ' OR SMARTCARDNO=' + STR_SEP + sI_FaciSNo + STR_SEP;
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_SQL);
end;


function TE2C.getRendomFileName: String;
var
    sL_TmpXmlFileName : String;
    nL_Seed : Integer;
begin
    Randomize;
    nL_Seed := 100000;
    sL_TmpXmlFileName := 'C:\CS' + IntToStr(Random(nL_Seed)) + '.xml';

    while FileExists(sL_TmpXmlFileName) do
      getRendomFileName;

    result := sL_TmpXmlFileName;

end;


procedure TE2C.TransTmpIniFile;
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath ,sL_IniFileName,sL_TempIniFileName : String;
    ii : Integer;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
    sL_Temp : WideString;
begin
    sL_EncKey := 'CS';
    L_Intf := CoEncrypt.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);


    sL_IniFileName := SYS_INI_FILE_NAME;

    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sL_IniFileName);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
          or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME') then
       begin
         L_TmpStrList.Add(L_StrList.Strings[ii]);
       end
       else
       begin
         sL_Temp := L_StrList.Strings[ii];
         L_TmpStrList.Add(L_Intf.Decrypt(sL_Temp));
       end;

    end;
    L_TmpStrList.SaveToFile(TMP_SYS_INI_FILE_NAME);
    L_TmpStrList.Free;
    L_Intf := nil;

end;

procedure TE2C.DeleteTmpIniFile;
var
    sL_ExeFileName,sL_ExePath,sL_IniFileName : String;
begin
    //�R���ѱK�᪺ini�ɮ�
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    //sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    sL_IniFileName := TMP_SYS_INI_FILE_NAME;
    DeleteFile(sL_IniFileName);
end;

function TE2C.getAreaDataNo(sI_CompCode: String): String;
var
    nL_Ndx,nL_AreaDatNo : Integer;
begin
    nL_Ndx := G_StrList.IndexOf(sI_CompCode);
//G_StrList.SaveToFile('c:\ccc.txt');
    if nL_Ndx<>-1 then
    begin
       nL_AreaDatNo := (G_StrList.Objects[nL_Ndx] as TMssIniData).nDataAreaNo;

       Result := IntToStr(nL_AreaDatNo);
    end;


end;

procedure TE2C.getIniData(sI_CompCode: String; var sI_DbPassword,
  sI_DbAlias, sI_DbUserName: String);
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_StrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       sI_DbPassword := (G_StrList.Objects[nL_Ndx] as TMssIniData).sPassword;
       sI_DbAlias := (G_StrList.Objects[nL_Ndx] as TMssIniData).sAlias;
       sI_DbUserName := (G_StrList.Objects[nL_Ndx] as TMssIniData).sUserID;
    end;
end;

function TE2C.checkIsConnectCompCode(sI_CompCode: String): Boolean;
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_InCompCodeStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
      Result := true
    else
      Result := false;
end;

procedure TE2C.allocateSpace(sI_CompID : String);
var
    sL_SQL : String;
begin
    sL_SQL := 'alter table SOAC0201A allocate extent (size 5M)';
    runSQL(IUD_MODE,sI_CompID, sL_SQL);

    sL_SQL := 'alter table SOAC0201B allocate extent (size 5M)';
    runSQL(IUD_MODE,sI_CompID, sL_SQL);
end;

procedure TE2C.deallocateSpace(sI_CompID: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'alter table SOAC0201A deallocate unused';
    runSQL(IUD_MODE,sI_CompID, sL_SQL);

    sL_SQL := 'alter table SOAC0201B deallocate unused';
    runSQL(IUD_MODE,sI_CompID, sL_SQL);
end;

procedure TE2C.DbConnAryRollBack;
var
    ii : Integer;
begin
    for ii:=0 to High(G_DbConnAry) do
    begin
      if G_DbConnAry[ii].InTransaction then
        G_DbConnAry[ii].RollbackTrans;
    end;
end;

function TE2C.isNotWareHouseCode(sI_TableName,sI_SeqNo: String;var sI_CompCode : String): String;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := 'SELECT * FROM ' + sI_TableName + ' WHERE FaciSNo=''' + sI_SeqNo + '''';

    L_AdoQry := runSQL(SELECT_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);

    if L_AdoQry=nil then
    begin
      Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + IntToStr(nG_TOTAL_WAREHOUSE_ID);
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount = 0 then
        Result := ''
      else
      begin
        sI_CompCode := L_AdoQry.FieldByName('COMPCODE').AsString;
        if sI_CompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
          Result := ERROR_STRING
        else
          Result := '';
      end;
    end;
end;


procedure TE2C.E2C_DLL3(vI_Data1: OleVariant);//��h�Ƥ�function...
var
    sL_Result, sL_Data1 : String;
    sL_DbConnStr: String;
begin
    setDateTimeFormat;
    sL_Data1 := VarToStr(vI_Data1);
    //�N�[�K��ini�ɮ׸ѱK
    TransTmpIniFile;

    sL_Result := '';

    if (UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (Trim(sL_Data1)<>'')then
    begin
      createDbInfoStrList;
      //down, �q ini ��Ū���s�u�� DB �� information
      sL_Result := getDbConnInfo(G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList);
      if sL_Result<>'' then
      begin
        Set_E2C_DLL3_RESULT(sL_Result);
        exit;
      end;
      //up, �q ini ��Ū���s�u�� DB �� information
    end;

    //down, �B�z��h��..
    if (UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (Trim(sL_Data1)<>'')then
    begin
      saveToFile(sL_Data1, 'c:\E2C_DLL3-1.txt');
      
      sL_Result := processXML5(sL_Data1);
      if sL_Result<>'' then
      begin
        DeleteTmpIniFile;
        Set_E2C_DLL3_RESULT(sL_Result);
        exit;
      end;

    end;
//    Set_E2C_DLL3_RESULT('');
    //up, �B�z��h��..

    //�R���ѱK�᪺ini�ɮ�
    DeleteTmpIniFile;
    G_StrList.Free;
    Set_E2C_DLL3_RESULT(sL_Result);

end;

function TE2C.processXML5(sI_Xml: WideString): WideString;
var
    bL_HasUpdateData : boolean;
    L_XMLDoc : IXMLDOMDocument;
    L_NodeList, L_MaterialNodeList, L_SeqNodeList : IXMLDOMNodeList;
    L_BatchNoNode, L_MaterialNode, L_SeqNode : IXMLDOMNODE;
    ii, jj, kk : Integer;

    sL_TmpXmlFileName, sL_EquipTableName,sL_Result : String;
    sL_BatchNo, sL_Date, sL_Time, sL_CurDateTime : String;
    sL_MakerID, sL_MakerName, sL_Status, sL_CompID : String;
    sL_MaterialNo,sL_Quantity, sL_ItemName, sL_Equip : String;
    sL_SeqNo, sL_Msg, sL_SQL, sL_UpdEn ,sL_CheckCompCode : String;
    nL_Seed, nL_Quantity  : Integer;
    dL_ExecDateTime : TDate;
    sL_Operation, sL_UseFlagValue, sL_Error,sL_ReturnCompCode, sL_ConnDbResult : String;
    nL_ActiveCompCode : Integer;
    sL_UseFlag,sL_SeqNoString : String;
    L_AdoQry : TADOQuery;
    aModeFlag, aModelNo, aMsg: String;
begin

    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��
    
    try


      bL_HasUpdateData := false;
      nL_ActiveCompCode := -1;
      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//�����J������,�{���~�|���U����
      L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�

      //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
      sL_TmpXmlFileName := getRendomFileName;

      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

      sL_UpdEn := '���ƨt��';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

//      sL_TmpXmlFileName := 'c:\CS11448.xml';
      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin


        
        L_NodeList := L_XMLDoc.getElementsByTagName('�帹');     //���o��󤤤��Ҧ��帹 tag

        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// �帹
          //down, �ˬd�帹�榡
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            Result := '-2: �帹���׿��~, �帹=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, �ˬd�帹�榡

          sL_Operation := TMyXML.getSingleNodeText(L_BatchNoNode, '�@�~'); // 1:���, 2:�h��
          if (sL_Operation<>'1') and (sL_Operation<>'2') then
          begin
            Result := '-6: ��h�Ƨ@�~�N�X���~.�@�~�N�X:'+ sL_Operation;
            exit;
          end;

          if sL_Operation = '1' then    //���
          begin
            sL_UseFlagValue := '1';
            sL_UseFlag := '��ƧǸ�: ';
          end
          else
          begin
            sL_UseFlagValue := '0';     //�h��
            sL_UseFlag := '�h�ƧǸ�: ';
          end;

          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '���');

          //down, �ˬd����榡
          if length(sL_Date)<>DATE_LENGTH then
          begin
            Result := '-2: ������׿��~, ���=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            Result := '-2: �������榡����: <' + sL_Date + '>';
            exit;
          end;
          //up, �ˬd����榡

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '�ɶ�');
          //down, �ˬd�ɶ��榡
          if length(sL_Time)<>TIME_LENGTH then
          begin
            Result := '-2: �ɶ����׿��~, �ɶ�=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            Result := '-2: ����ɶ��榡����: <' + sL_Time + '>';
            exit;
          end;
          //up, �ˬd�ɶ��榡
          
          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '���q�N�X');
//saveToFile(sL_CompID, 'c:\abc.txt');

          nL_ActiveCompCode := StrToIntDef(sL_CompID,-1);
          sL_ConnDbResult := connecttodb(not CONN_TO_WAREHOUSE, sL_CompID);
          if sL_ConnDbResult<>'' then
          begin
            Result := '-6: �L�k�s���W��Ʈw'+'-' + sL_ConnDbResult;
            exit;

          end;
//saveToFile('connect db ok', 'c:\aa1.txt');
          //howard123
          if nL_ActiveCompCode = -1 then
          begin
            Result := '-6: ��h�Ƥ��q�N�X���~(��h�Ƥ��q�N�X:' + sL_CompID + ')';
            exit;
          end;
//SaveToFile(inttostr(nL_ActiveCompCode) + '-' + inttostr(nG_TOTAL_WAREHOUSE_ID), 'c:\zzz12.txt');
          if nL_ActiveCompCode= nG_TOTAL_WAREHOUSE_ID then
          begin
            Result := '-6: ��h�Ƥ��q�N�X�������`�ܥN�X';
            exit;
          end;
          //�N�{���]�b Transaction �����
          G_DbConnAry[0].BeginTrans;
//SaveToFile('abc', 'c:\zzz13.txt');


  //saveToFile(sL_BatchNo+sL_Date+sL_Time+sL_MakerID+sL_MakerName+sL_Status+sL_CompID, 'c:\aa2.txt');
          L_MaterialNodeList := L_BatchNoNode.selectNodes('./�Ƹ�');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// �Ƹ�
            //down, �ˬd�Ƹ��榡
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, �ˬd�Ƹ��榡

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'�ƶq');
            //down, �ˬd�ƶq�榡
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: �ƶq���~, �ƶq=<'+sL_Quantity+'>';
              exit;
            end;
            //up, �ˬd�ƶq�榡

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode, '�ȪA�~�W');
            if ( Length( sL_ItemName ) > 20 ) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: �ȪA�~�W���׳̦h20�X, �Ƹ�=<'+sL_MaterialNo+'>, �ȪA�~�W=<'+sL_ItemName+'>';
              Exit;
            end;

            { 4. �ˮָӤ��q�O�O�_���إߦ��ȪA�~�W }
            aMsg := VdCD022( sL_CompID , sL_ItemName );
            if ( aMsg <> EmptyStr ) then
            begin
              DbConnAryRollBack;
              Result := '-2:CC&B��Ʈw�����]�Ƭd�L�����ȪA�~�W,�ȪA�~�W=<'+sL_ItemName+'> , �Ƹ�=<'+sL_MaterialNo+'>';
              Exit;
            end;

            

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'�]��');
            //down, �ˬd�]�Ʈ榡
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: ���Ƹ����]�ƧO����, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+sL_Equip+'>';
              exit;
            end;
            //up, �ˬd�]�Ʈ榡


            { �u�� STB �~�� MODEFLAG }
            aModeFlag := EmptyStr;
            if ( sL_Equip = '1' ) then
            begin
              aModeFlag := TMyXML.getAttributeValue(L_MaterialNode,'����' );
              if ( Trim( aModeFlag ) <> '0' ) and ( Trim( aModeFlag ) <> '1'  ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: ���Ƹ�����������, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aModelNo := TMyXML.getAttributeValue(L_MaterialNode,'����' );
              if ( Trim( aModelNo ) = EmptyStr ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: ���Ƹ������ج��ŭ�, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aMsg := VdCD043( sL_CompID, aModelNo );
              if ( aMsg <> EmptyStr ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: ���]�Ƭd�L�������ث���,����=<'+aModelNo+'> , �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+aModeFlag+'>';
                Exit;
              end;
              {} 
            end;

            L_SeqNodeList := L_MaterialNode.selectNodes('./�Ǹ�');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: ���Ƹ����ƶq�P���ӵ��Ƥ��X, �Ƹ�=<'+sL_MaterialNo+'>, �ƶq=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
  //showmessage('dL_ExecDateTime  ' + DateTimeToStr(dL_ExecDateTime));
            case StrToInt(sL_Equip) of
              1: //�� STB
               begin
                sL_EquipTableName := 'SOAC0201A';
               end;
              2: //�� ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
               end;
            end;



            sL_SeqNoString := '';   //��h��LOG��
            for kk:=0 to L_SeqNodeList.length -1 do
            begin
            
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, �ˬd�Ǹ��榡
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: �Ǹ����׿��~, �Ǹ�=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, �ˬd�Ǹ��榡



              //down, �ˬd���Ǹ��O�_�w�s�b

              sL_SQL := 'SELECT COUNT(*) COUNT FROM ' + sL_EquipTableName;
              sL_SQL := sL_SQL + ' where  FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;

              L_AdoQry := runSQL(SELECT_MODE,sL_CompID, sL_SQL);


              if L_AdoQry.FieldByName('COUNT').AsInteger = 0 then
              begin

                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: ���]�ƧǸ����s�b, �Ǹ�=<'+sL_SeqNo+'>';
                exit;
              end;

              //up, �ˬd���Ǹ��O�_�w�s�b


              //down, ��sCC&B�������]�Ƹ����
              //howard..

              sL_SQL := 'update '+ sL_EquipTableName +' set UseFlag=' + sL_UseFlagValue;
              if ( sL_Equip = '1' ) then
              begin
                sL_SQL := sL_SQL + ',MODEFLAG=' + STR_SEP + aModeFlag + STR_SEP + ', MODELNO=' + STR_SEP + aModelNo + STR_SEP;

              end;
              sL_SQL := sL_SQL + ',UPDEN=' + STR_SEP + '���ƨt��' + STR_SEP ;
              sL_SQL := sL_SQL + ',UpdTime=' + STR_SEP + sL_CurDateTime + STR_SEP ;
              sL_SQL := sL_SQL + ' where FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;


              try
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              finally
                bL_HasUpdateData := true;
              end;
              //up, ��sCC&B�������]�Ƹ����


              //��_��h�Ʀr��
              if sL_SeqNoString = '' then
                sL_SeqNoString := sL_SeqNo
              else
                sL_SeqNoString := sL_SeqNoString + ',' + sL_SeqNo;

            end;


            sL_SeqNoString := sL_UseFlag + sL_SeqNoString;
            //Log ��h�Ƹ�Ʀ� SOAC0204
            sL_SQL := 'INSERT INTO  SOAC0204 VALUES(' + STR_SEP + sL_CurDateTime +
                      STR_SEP + ',' + STR_SEP + sL_CompID +
                      STR_SEP + ',''E2C_DLL3'',' + STR_SEP + sL_SeqNoString + STR_SEP + ')';


            runSQL(IUD_MODE,sL_CompID, sL_SQL);

          end;
          //howard123
        end;

        Result := '';

      end else
      begin
        Result := '-1:�ǤJ����� XML �榡���~!���˵�: ' + sL_TmpXmlFileName;
        Exit;
      end;


        Result := EmptyStr;
      //�Y���椤�S�����D,�hCommit

      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].CommitTrans;
    except
      //�Y���椤�����D,�hRollback
      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].RollbackTrans;
    end;


end;

{
function TE2C.processXML5(sI_Xml: WideString): WideString;
var
    bL_HasUpdateData : boolean;
    L_XMLDoc : IXMLDOMDocument;
    L_NodeList, L_MaterialNodeList, L_SeqNodeList : IXMLDOMNodeList;
    L_BatchNoNode, L_MaterialNode, L_SeqNode : IXMLDOMNODE;
    ii, jj, kk : Integer;

    sL_TmpXmlFileName, sL_EquipTableName,sL_Result : String;
    sL_BatchNo, sL_Date, sL_Time, sL_CurDateTime : String;
    sL_MakerID, sL_MakerName, sL_Status, sL_CompID : String;
    sL_MaterialNo,sL_Quantity, sL_ItemName, sL_Equip : String;
    sL_SeqNo, sL_Msg, sL_SQL, sL_UpdEn ,sL_CheckCompCode : String;
    nL_Seed, nL_Quantity  : Integer;
    dL_ExecDateTime : TDate;
    sL_Operation, sL_UseFlagValue, sL_Error,sL_ReturnCompCode, sL_ConnDbResult : String;
    nL_ActiveCompCode : Integer;
begin

    //�]�g�b Try ���| AutoCommit , �ҥH�g�b Transaction �~��

    
    try

      bL_HasUpdateData := false;
      nL_ActiveCompCode := -1;
      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//�����J������,�{���~�|���U����
      L_XMLDoc.ValidateOnParse := False;//�O�_�����Ҥ��PDTD�O�_�۲�

      //down, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�
      sL_TmpXmlFileName := getRendomFileName;

      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, ���� loadXML �|�� error, �ҥH���N�� XML �s��Ȧs�ɤ�

      sL_UpdEn := '���ƨt��';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

//      sL_TmpXmlFileName := 'c:\CS11448.xml';
      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin


        
        L_NodeList := L_XMLDoc.getElementsByTagName('�帹');     //���o��󤤤��Ҧ��帹 tag

        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// �帹
          //down, �ˬd�帹�榡
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            Result := '-2: �帹���׿��~, �帹=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, �ˬd�帹�榡

          sL_Operation := TMyXML.getSingleNodeText(L_BatchNoNode, '�@�~'); // 1:���, 2:�h��
          if (sL_Operation<>'1') and (sL_Operation<>'2') then
          begin
            Result := '-6: ��h�Ƨ@�~�N�X���~.�@�~�N�X:'+ sL_Operation;
            exit;
          end;

          if sL_Operation = '1' then
            sL_UseFlagValue := '1' //���
          else
            sL_UseFlagValue := '0'; //�h��
            
          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '���');

          //down, �ˬd����榡
          if length(sL_Date)<>DATE_LENGTH then
          begin
            Result := '-2: ������׿��~, ���=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            Result := '-2: �������榡����: <' + sL_Date + '>';
            exit;
          end;
          //up, �ˬd����榡

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '�ɶ�');
          //down, �ˬd�ɶ��榡
          if length(sL_Time)<>TIME_LENGTH then
          begin
            Result := '-2: �ɶ����׿��~, �ɶ�=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            Result := '-2: ����ɶ��榡����: <' + sL_Time + '>';
            exit;
          end;
          //up, �ˬd�ɶ��榡
          
          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '���q�N�X');
//saveToFile(sL_CompID, 'c:\abc.txt');

          nL_ActiveCompCode := StrToIntDef(sL_CompID,-1);
          sL_ConnDbResult := connecttodb(not CONN_TO_WAREHOUSE, sL_CompID);
          if sL_ConnDbResult<>'' then
          begin
            Result := '-6: �L�k�s���W��Ʈw';
            exit;

          end;
//saveToFile('connect db ok', 'c:\aa1.txt');
          //howard123
          if nL_ActiveCompCode = -1 then
          begin
            Result := '-6: ��h�Ƥ��q�N�X���~(��h�Ƥ��q�N�X:' + sL_CompID + ')';
            exit;
          end;
//SaveToFile(inttostr(nL_ActiveCompCode) + '-' + inttostr(nG_TOTAL_WAREHOUSE_ID), 'c:\zzz12.txt');
          if nL_ActiveCompCode= nG_TOTAL_WAREHOUSE_ID then
          begin
            Result := '-6: ��h�Ƥ��q�N�X�������`�ܥN�X';
            exit;
          end;
          //�N�{���]�b Transaction �����
          G_DbConnAry[0].BeginTrans;
//SaveToFile('abc', 'c:\zzz13.txt');


  //saveToFile(sL_BatchNo+sL_Date+sL_Time+sL_MakerID+sL_MakerName+sL_Status+sL_CompID, 'c:\aa2.txt');
          L_MaterialNodeList := L_BatchNoNode.selectNodes('./�Ƹ�');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// �Ƹ�
            //down, �ˬd�Ƹ��榡
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: �Ƹ����׿��~, �Ƹ�=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, �ˬd�Ƹ��榡

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'�ƶq');
            //down, �ˬd�ƶq�榡
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: �ƶq���~, �ƶq=<'+sL_Quantity+'>';
              exit;
            end;
            //up, �ˬd�ƶq�榡

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode,'�~�W');

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'�]��');
            //down, �ˬd�]�Ʈ榡
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: ���Ƹ����]�ƧO����, �Ƹ�=<'+sL_MaterialNo+'>, �]�ƧO=<'+sL_Equip+'>';
              exit;
            end;
            //up, �ˬd�]�Ʈ榡


  //          showmessage(sL_MaterialNo+sL_Quantity + sL_ItemName + sL_Spec + sL_Equip);
  //saveToFile(sL_MaterialNo+sL_Quantity + sL_ItemName + sL_Spec + sL_Equip, 'c:\aa3.txt');

            L_SeqNodeList := L_MaterialNode.selectNodes('./�Ǹ�');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: ���Ƹ����ƶq�P���ӵ��Ƥ��X, �Ƹ�=<'+sL_MaterialNo+'>, �ƶq=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
  //showmessage('dL_ExecDateTime  ' + DateTimeToStr(dL_ExecDateTime));
            case StrToInt(sL_Equip) of
              1: //�� STB
               begin
                sL_EquipTableName := 'SOAC0201A';
               end;
              2: //�� ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
               end;
            end;


            for kk:=0 to L_SeqNodeList.length -1 do
            begin
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, �ˬd�Ǹ��榡
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: �Ǹ����׿��~, �Ǹ�=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, �ˬd�Ǹ��榡




              //down, ��sCC&B�������]�Ƹ����
              //howard..

              sL_SQL := 'update '+ sL_EquipTableName +' set UseFlag=' + sL_UseFlagValue;
              sL_SQL := sL_SQL + ',UPDEN=' + STR_SEP + '���ƨt��' + STR_SEP ;
              sL_SQL := sL_SQL + ' where  FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;
//SaveToFile(sL_SQL, 'c:\aaa11.txt');


              try
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              finally
                bL_HasUpdateData := true;
              end;  
              //up, ��sCC&B�������]�Ƹ����

            end;
          end;
          //howard123
        end;

        Result := '';

      end
      else
        Result := '-1:�ǤJ����� XML �榡���~!���˵�: ' + sL_TmpXmlFileName;


        Result := '';
      //�Y���椤�S�����D,�hCommit
//savetofile(inttostr(nL_ActiveCompCode),'c:\xxx.txt');

      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].CommitTrans;
//savetofile('ok...','c:\xxx1.txt');
    except
      //�Y���椤�����D,�hRollback
//savetofile('not ok...','c:\xxx2.txt');
      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].RollbackTrans;
    end;


end;
}

function TE2C.Get_E2C_DLL3_RESULT: WideString;
begin
    result := sG_E2C_DLL3_RESULT;
end;

procedure TE2C.Set_E2C_DLL3_RESULT(const Value: WideString);
begin
    sG_E2C_DLL3_RESULT := Value;
end;


procedure TE2C.setDateTimeFormat;
begin
  DateSeparator := '/';
  TimeSeparator := ':';
  LongDateFormat := 'yyyy/mm/dd';
  ShortDateFormat := 'yyyy/mm/dd';
  ShortTimeFormat :='hh:nn:ss';
  LongTimeFormat :='hh:nn:ss';
  TimeAMString := EmptyStr;
  TimePMString := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

function TE2C.GetModelNo(const aCompCode, aFacisNo: String; var aModelNo: String): String;
var
  aSql: String;
  aQuery: TADOQuery;
begin
  Result := EmptyStr;
  aModelNo := EmptyStr;
  aSql := Format(
    ' select modelno from soac0201a    ' +
    '  where compcode = ''%s''         ' +
    '    and facisno = ''%s''          ', [aCompCode, aFacisNo] );
  aQuery := runSQL(SELECT_MODE, aCompCode, aSql);
  if aQuery = nil then
  begin
    Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + aCompCode;
    Exit;
  end;
  aModelNo := aQuery.FieldByName( 'MODELNO' ).AsString;
  if ( aModelNo = EmptyStr ) then
  begin
    Result := 'Error';
  end;
  aQuery.Close;
  aQuery := nil;
end;

{ ---------------------------------------------------------------------------- }

function TE2C.GetDescription(const aType, aCompCode, aFacisNo: String;
  var aDescription: String): String;
var
  aSql: String;
  aQuery: TADOQuery;
begin
  Result := EmptyStr;
  aDescription := EmptyStr;
  if ( aType = '1' ) then
  begin
    aSql := Format(
      ' select description from soac0201a  ' +
      '  where compcode = ''%s''           ' +
      '    and facisno = ''%s''            ', [aCompCode, aFacisNo] );
  end else
  begin
    aSql := Format(
      ' select description from soac0201b  ' +
      '  where compcode = ''%s''           ' +
      '    and facisno = ''%s''            ', [aCompCode, aFacisNo] );
  end;    
  aQuery := runSQL(SELECT_MODE, aCompCode, aSql);
  if aQuery = nil then
  begin
    Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + aCompCode;
    Exit;
  end;
  aDescription := aQuery.FieldByName( 'DESCRIPTION' ).AsString;
  if ( aDescription = EmptyStr ) then
  begin
    Result := 'Error';
  end;
  aQuery.Close;
  aQuery := nil;
end;

{ ---------------------------------------------------------------------------- }

function TE2C.VdCD043(const aCompCode, aModelNo: String): String;
var
  aSql: String;
  aQuery: TADOQuery;
begin
  Result := EmptyStr;
  aSql := Format(
    ' select codeno from cd043     ' +
    '  where description = ''%s''  ', [aModelNo] );
  aQuery := runSQL( SELECT_MODE, aCompCode, aSql );
  if aQuery = nil then
  begin
    Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + aCompCode;
    Exit;
  end;
  if ( aQuery.IsEmpty ) or ( aQuery.RecordCount > 1 ) then
  begin
    Result := 'Error';
  end;
  aQuery.Close;
  aQuery := nil;
end;

{ ---------------------------------------------------------------------------- }

function TE2C.VdCD022(const aCompCode, aDescription: String): String;
var
  aSql: String;
  aQuery: TADOQuery;
begin
  Result := EmptyStr;
  aSql := Format(
    ' select codeno from cd022     ' +
    '  where description = ''%s''  ', [aDescription] );
  aQuery := RunSQL( SELECT_MODE, aCompCode, aSql );
  if aQuery = nil then
  begin
    Result := NO_AVAILABLE_QUERY_COMP + '  ���q�O=' + aCompCode;
    Exit;
  end;
  if ( aQuery.IsEmpty ) or ( aQuery.RecordCount > 1 ) then
  begin
    Result := 'Error';
  end;
  aQuery.Close;
  aQuery := nil;
end;

{ ---------------------------------------------------------------------------- }


initialization
  TAutoObjectFactory.Create(ComServer, TE2C, Class_E2C,
    ciMultiInstance, tmApartment);
end.

