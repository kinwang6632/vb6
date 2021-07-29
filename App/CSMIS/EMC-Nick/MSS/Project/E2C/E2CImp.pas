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


    BATCH_NO_LENGTH = 12;             //批號之最大長度
    DATE_LENGTH = 7;                  //日期之長度
    TIME_LENGTH = 4;                  //時間之長度
    MATERIAL_NO_LENGTH = 15;          //料號之最大長度
    SEQ_NO_LENGTH = 15;               //序號之最大長度
    NO_AVAILABLE_QUERY_COMP = '-9: 無法取得適當的元件reference.以致無法讀取資料庫之資料.';
    SELECT_MODE = 1;
    IUD_MODE = 2;
    ERROR_TOTAL_WAREHOUSE_ID = 54088; //判定傳入字串是否為純數字,若不是則等於此常數
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

    function processXML1(sI_Xml:WideString):WideString;  //處理入庫資料
    function processXML2(sI_Xml:WideString):WideString;  //處理驗退資料
    function processXML3(sI_Xml:WideString):WideString;  //處理調撥資料
    function processXML4(sI_Xml:WideString):WideString;  //處理配對資料
    function processXML5(sI_Xml:WideString):WideString;  //處理領退料資料    

    function getDbConnInfo(var I_DbAliasStrList, I_DbUserNameStrList, I_DbPasswordStrList : TStringList):WideString;
    function connectToDB(bI_ConnToWarehouse : boolean; sI_InCompCodeList : String) : String;

    function checkSeqNo1(sI_CompID, sI_EquipNo, sI_SeqNo:String;
             var sI_Msg:String; var bI_EquipExist:boolean ):boolean;  //入庫
    function checkSeqNo2(sI_Function, sI_CompID, sI_EquipNo, sI_SeqNo:String;
             var sI_Msg:String; var bI_EquipExist:boolean ):boolean;  //驗退

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

      if bI_ConnToWarehouse then //若要 connect 到總倉
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

        //取得傳入公司的資料庫連線資料
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
      Result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';
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
  //將加密的ini檔案解密
  TransTmpIniFile;
  aResult := EmptyStr;
  if ( UpperCase( trim( sL_Data1 ) ) <> XML_NO_DATA ) or
     ( UpperCase( trim( sL_Data2 ) ) <> XML_NO_DATA ) then
  begin
    createDbInfoStrList;
    // 從 ini 檔讀取連線至 DB 的 information
    aResult := getDbConnInfo( G_DbAliasStrList, G_DbUserNameStrList,
      G_DbPasswordStrList );
    if ( aResult<> EmptyStr ) then
    begin
      Set_E2C_DLL1_RESULT( aResult );
      Exit;
    end;
    // 建立DB的 connection
    aResult := connectToDB( CONN_TO_WAREHOUSE, EmptyStr );
    if ( aResult <> '' ) then
    begin
      Set_E2C_DLL1_RESULT( aResult );
      Exit;
    end;
  end;

  aStrList := TStringList.Create;
  try
    //處理入庫..
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

    //down, 處理驗退..
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
  //刪除解密後的ini檔案
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
    //將加密的ini檔案解密
    TransTmpIniFile;
    aResult := EmptyStr;
    if ((UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (trim(sL_Data1)<>''))
    or  ((UpperCase(trim(sL_Data2))<>XML_NO_DATA) and (trim(sL_Data2)<>''))then
    begin
      createDbInfoStrList;
      //down, 從 ini 檔讀取連線至 DB 的 information
      aResult := getDbConnInfo(G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList);
      if ( aResult <> EmptyStr ) then
      begin
        Set_E2C_DLL2_RESULT(aResult);
        exit;
      end;
      //up, 從 ini 檔讀取連線至 DB 的 information
      aResult := connectToDB(CONN_TO_WAREHOUSE, sL_Data3);
      if ( aResult <> EmptyStr ) then
      begin
        Set_E2C_DLL2_RESULT(aResult);
        exit;
      end;
    end;
    //處理調撥..
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
    //down, 處理配對..
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
    //刪除解密後的ini檔案
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



//處理入庫資料
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

    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面
    allocateSpace(IntToStr(nG_TOTAL_WAREHOUSE_ID));

    try
      //將程式包在 Transaction 交易中
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].BeginTrans;


      L_XMLDoc := CoDOMDocument.Create;



      L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
      L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符


      //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

      sL_UpdEn := '物料系統';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        //down, 檢查公司代碼是否為總倉
        L_NodeList := L_XMLDoc.selectNodes('//公司代碼');
        if L_NodeList<> nil then
        begin
          for ii:=0 to L_NodeList.length -1 do
          begin
            sL_CheckCompCode := L_NodeList.item[ii].text;
            //showmessage('總倉代碼:' + IntToStr(nG_TOTAL_WAREHOUSE_ID) + '   公司代碼:' + sL_CheckCompCode);
            if sL_CheckCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-5: 入庫公司代碼應為總倉代碼<' + IntToStr(nG_TOTAL_WAREHOUSE_ID) + '>,錯誤公司代碼=<'+sL_CheckCompCode + '>';
            Exit;
            end;
          end;
        end;
        //up, 檢查公司代碼是否為總倉



        L_NodeList := L_XMLDoc.getElementsByTagName('批號');     //取得文件中之所有批號 tag
        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// 批號
          //down, 檢查批號格式
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 批號長度錯誤, 批號=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, 檢查批號格式

          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '日期');
          //down, 檢查日期格式
          if length(sL_Date)<>DATE_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 日期長度錯誤, 日期=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 執行日期格式不對: <' + sL_Date + '>';
            exit;
          end;
          //up, 檢查日期格式

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '時間');
          //down, 檢查時間格式
          if length(sL_Time)<>TIME_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 時間長度錯誤, 時間=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 執行時間格式不對: <' + sL_Time + '>';
            exit;
          end;
          //up, 檢查時間格式

          sL_MakerID := TMyXML.getSingleNodeText(L_BatchNoNode, '廠商代碼');
          if length(sL_MakerID)>VENDORCODE_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 廠商代碼的長度最多 ' + IntToStr(VENDORCODE_LENGTH) + ' 碼';
            exit;
          end;

          sL_MakerName := TMyXML.getSingleNodeText(L_BatchNoNode, '廠商名稱');
          if length(sL_MakerName)>VENDORNAME_LENGTH then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 廠商名稱的長度最多 ' + IntToStr(VENDORNAME_LENGTH) + ' 碼';
            exit;
          end;


          // 根本沒用到, 拿掉, 介接的 xml 不傳也 ok
          //sL_Status := TMyXML.getSingleNodeText(L_BatchNoNode, '狀況');
          sL_Status := '1';
          //down, 檢查進出(入庫/驗退)狀況格式 => 此參數保留,一定要傳入 1, 留待日後擴充
          if not (StrToInt(sL_Status) in [1]) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 此批號之進出狀況不對, 批號=<'+sL_BatchNo+'>, 進出狀況=<'+sL_Status+'>';
            exit;
          end;
          //up, 檢查進出(入庫/驗退)狀況格式 => 此參數保留,一定要傳入 1, 留待日後擴充

          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '公司代碼');


          L_MaterialNodeList := L_BatchNoNode.selectNodes('./料號');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// 料號
            //down, 檢查料號格式
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 料號長度錯誤, 料號=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, 檢查料號格式

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'數量');
            //down, 檢查數量格式
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 數量錯誤, 數量=<'+sL_Quantity+'>';
              exit;
            end;
            //up, 檢查數量格式

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode, '客服品名' );
            if ( Length( sL_ItemName ) > 20 ) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 客服品名長度最多20碼, 料號=<'+sL_MaterialNo+'>, 客服品名=<'+sL_ItemName+'>';
              Exit;
            end;


            sL_Spec := TMyXML.getAttributeValue(L_MaterialNode,'規格');

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'設備');
            //down, 檢查設備格式
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 此料號之設備別不對, 料號=<'+sL_MaterialNo+'>, 設備別=<'+sL_Equip+'>';
              exit;
            end;
            //up, 檢查設備格式

            { 當為 STB 時, 才有此 Flag }
            
            aModeFlag := EmptyStr;
            aModelNo := EmptyStr;
            aMaterialType := EmptyStr;
            aHwVer := EmptyStr;
            aCallbackMode := EmptyStr;
            aDirection := EmptyStr;
            aHdSize := EmptyStr;

            if ( sL_Equip = '1' ) then
            begin
              aModeFlag := TMyXML.getAttributeValue(L_MaterialNode,'型式' );
              if ( Trim( aModeFlag ) <> '0' ) and ( Trim( aModeFlag ) <> '1'  ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 此料號之型式不對, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aModelNo := TMyXML.getAttributeValue(L_MaterialNode,'機種' );
              if ( Trim( aModelNo ) = EmptyStr ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 此料號之機種為空值, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              if ( Length( aModelNo ) > 20 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 機種名稱長度最多20碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aMaterialType := TMyXML.getAttributeValue( L_MaterialNode, '物料型號' );
              if ( Length( aMaterialType ) > 50 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 物料型號長度最多50碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aHwVer := TMyXML.getAttributeValue( L_MaterialNode, '硬體版本' );
              if ( Length( aHwVer ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 硬體版本長度最多10碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aCallbackMode := TMyXML.getAttributeValue( L_MaterialNode, '回傳模式' );
              if ( Length( aCallbackMode ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 回傳模式長度最多10碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aDirection := TMyXML.getAttributeValue( L_MaterialNode, '單雙向' );
              if ( Length( aDirection ) > 10 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 單雙向長度最多10碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aHdSize := TMyXML.getAttributeValue( L_MaterialNode, '硬碟容量' );
              if ( Length( aHdSize ) > 20 ) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
                Result := '-2: 硬碟容量長度最多20碼, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
            end;

            L_SeqNodeList := L_MaterialNode.selectNodes('./序號');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
              Result := '-2: 此料號之數量與明細筆數不合, 料號=<'+sL_MaterialNo+'>, 數量=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
            case StrToInt(sL_Equip) of
              1: sL_EquipTableName := 'SOAC0201A'; //為 STB
              2: sL_EquipTableName := 'SOAC0201B'; //為 ICC
            end;


            for kk:=0 to L_SeqNodeList.length -1 do
            begin
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, 檢查序號格式
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                Result := '-2: 序號長度錯誤, 序號=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, 檢查序號格式


              //檢查該設備序號是否存在於總倉,且公司別不等於總倉代碼
              sL_Error := isNotWareHouseCode(sL_EquipTableName,sL_SeqNo,sL_ReturnCompCode);
              if sL_Error <> '' then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                if sL_Error = ERROR_STRING then
                  Result := '-2: 此序號=<' + sL_SeqNo + '>設備已調撥至公司別=<' + sL_ReturnCompCode + '> 不能再入庫'
                else
                  Result := sL_Error;
                exit;
              end;


              //down, 檢查序號是否合理
              if not checkSeqNo1(sL_CompID, sL_Equip, sL_SeqNo,sL_Msg, bL_EquipExist) then
              begin
                if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                  G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                Result := sL_Msg;
                exit;
              end;
              //up, 檢查序號是否合理
              //down, 更新CC&B的相關設備資料檔
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
                { 當為 = 1 時(STB), 多7個欄位 MODEFLAG, MODELNO, MaterialType, HwVer, CallbackMode, Direction, HdSize }
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

              //up, 更新CC&B的相關設備資料檔
            end;
          end;
        end;

        Result := '';
      end
      else
        Result := '-1:傳入之入庫 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;


      //若執行中沒有問題,則Commit
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].CommitTrans;

    except
      //若執行中有問題,則Rollback
      on E: Exception do
      begin
        Set_E2C_DLL1_RESULT('ErrorMsg:' + E.Message);
        G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;
      end;
    end;


    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面
    deallocateSpace(IntToStr(nG_TOTAL_WAREHOUSE_ID));
end;


//處理配對資料
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

    L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
    L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符

    //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
    sL_TmpXmlFileName := getRendomFileName;
    saveToFile(sI_Xml, sL_TmpXmlFileName);

    sL_Para := sL_TmpXmlFileName + ';';
    //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

    sL_UpdEn := '物料系統';
    sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

    if L_XMLDoc.load(sL_TmpXmlFileName) then
    begin

      L_NodeList := L_XMLDoc.getElementsByTagName('配對');  //取得文件中之所有配對 tag

      for ii := 0 to L_NodeList.length-1 do
      begin
        L_PairNode := L_NodeList.item[ii];


        sL_STBMaterailNo := L_PairNode.attributes.item[0].text;
        sL_STBFaciSNo := L_PairNode.attributes.item[1].text;
        sL_ICCMaterailNo := L_PairNode.attributes.item[2].text;
        sL_ICCFaciSNo := L_PairNode.attributes.item[3].text;

        //down, 檢查料號格式
        if not (length(sL_STBMaterailNo)<=MATERIAL_NO_LENGTH) then
        begin
          Result := '-2: 料號長度錯誤, 料號=<'+sL_STBMaterailNo + '>';
          exit;
        end
        else if (not (length(sL_ICCMaterailNo)<=MATERIAL_NO_LENGTH))then
        begin
          Result := '-2: 料號長度錯誤, 料號=<'+sL_ICCMaterailNo + '>';
          exit;
        end;
        //up, 檢查料號格式

        //檢查STB設備序號是否存在
        sL_ErrStr1 := '-4: CC&B資料庫無此STB序號, 無法配對, 序號=<'+sL_STBFaciSNo+'>';
        if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),'1',sL_STBFaciSNo,sL_ErrStr1,sL_Msg) then
        begin
          Result := sL_Msg;
          exit;
        end;

        //檢查ICC設備序號是否存在
        sL_ErrStr2 := '-4: CC&B資料庫無此ICC序號, 無法配對, 序號=<'+sL_ICCFaciSNo+'>';
        if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),'2',sL_ICCFaciSNo,sL_ErrStr2,sL_Msg) then
        begin
          Result := sL_Msg;
          exit;
        end;

        //sL_CompCode := '';  //SOAC0201C 暫不儲存公司代碼
        //先刪除已存在的該序號,再新增,以確保資料
        deletePairData1(sL_STBFaciSNo,sL_IccFaciSNo);
        insertPairData(sL_STBFaciSNo,sL_IccFaciSNo);

      end;

      //Log一筆至共同資料區東富MSS呼叫記錄檔(SOAC0204)
      //SOAC0201C 暫不儲存公司代碼
      sL_Para := '配對檔 : ' + sL_Para;
      logDataToSOAC0204(sL_CurDateTime,'','E2C_DLL2',sL_Para);
      Result := '';

    end
    else
      Result := '-1:傳入之配對 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;


end;


//處理調撥資料
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

    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面
    for ii:=0 to G_InCompCodeStrList.Count-1 do
      allocateSpace(G_InCompCodeStrList.Strings[ii]);

    try
      //將程式包在 Transaction 交易中
      for ii:=0 to High(G_DbConnAry) do
      begin
        G_DbConnAry[ii].BeginTrans;
      end;

      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
      L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符

      //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);

      sL_Para := sL_TmpXmlFileName + ';';
      //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

      sL_UpdEn := '物料系統';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        L_NodeList := L_XMLDoc.getElementsByTagName('料號');  //取得文件中之所有料號 tag
        //檢查XML內要調撥的公司別,是否有不存在有連線的資料庫
        for ii := 0 to L_NodeList.length-1 do
        begin
          L_MaterialNoNode := L_NodeList.item[ii];
          sL_MaterialNo :=  L_MaterialNoNode.childNodes.item[0].text;
          sL_InCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'調入公司');


          bL_IsConnectCompCode := checkIsConnectCompCode(sL_InCompCode);
          if not bL_IsConnectCompCode then
          begin
            //若執行中有問題,則Rollback
            DbConnAryRollBack;

            Result := '-3: , 要調入的公司別,不包含在有連線的公司中' + sL_InCompCode;
            exit;
          end;


          sL_OutCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'調出公司');
          bL_IsConnectCompCode := checkIsConnectCompCode(sL_OutCompCode);
          if not bL_IsConnectCompCode then
          begin
            //若執行中有問題,則Rollback
            DbConnAryRollBack;

            Result := '-3: , 要調出的公司別,不包含在有連線的公司中' + sL_OutCompCode;
            exit;
          end;
        end;


        for ii := 0 to L_NodeList.length-1 do
        begin

          L_MaterialNoNode := L_NodeList.item[ii];
          sL_MaterialNo :=  L_MaterialNoNode.childNodes.item[0].text;

          sL_EquipType := TMyXML.getAttributeValue(L_MaterialNoNode,'設備別');
          sL_InCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'調入公司');
          sL_OutCompCode := TMyXML.getAttributeValue(L_MaterialNoNode,'調出公司');

          //down, 檢查料號格式
          if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
          begin
            //若執行中有問題,則Rollback
            DbConnAryRollBack;

            Result := '-2: 料號長度錯誤, 料號=<'+sL_MaterialNo + '>';
            exit;
          end;
          //up, 檢查料號格式

          //down, 檢查公司別代碼
          if sL_InCompCode = '' then
          begin
            //若執行中有問題,則Rollback
            DbConnAryRollBack;

            Result := '-2: 此調撥之調入公司代碼不能為空值';
            exit;
          end;

          //檢查調入公司代碼是否存在於 CD039 中
          if sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
          begin
            if not validCompCode(sL_InCompCode,sL_Msg) then
            begin
              //若執行中有問題,則Rollback
              DbConnAryRollBack;

              Result := sL_Msg;
              exit;
            end;
          end;

          //UP, 檢查公司別代碼


          //down, 檢查設備格式
          if not (StrToInt(sL_EquipType) in [1,2]) then
          begin
            //若執行中有問題,則Rollback
            DbConnAryRollBack;

            Result := '-2: 此調撥之設備別不對, 設備別=<'+sL_EquipType+'>';
            exit;
          end;
          //up, 檢查設備格式
          if ii = 0 then
            sL_Para := sL_Para + '設備別:' +sL_EquipType + ',調入公司:' + sL_InCompCode + ',料號:' + sL_MaterialNo
          else
            sL_Para := sL_Para + '-設備別:' + sL_EquipType + ',調入公司:' + sL_InCompCode + ',料號:' + sL_MaterialNo;

          case StrToInt(sL_EquipType) of
            1: //為 STB
             begin
              sL_EquipTableName := 'SOAC0201A';
              sL_ErrStr1 := '-4: CC&B資料庫無此STB序號, 無法調撥 ,序號=<'+sL_SeqNo+'>';
             end;
            2: //為 ICC
             begin
              sL_EquipTableName := 'SOAC0201B';
              sL_ErrStr1 := '-4: CC&B資料庫無此ICC序號, 無法調撥 ,序號=<'+sL_SeqNo+'>';
             end;
          end;


          L_SeqNodeList := L_MaterialNoNode.selectNodes('./序號');


          for kk:=0 to L_SeqNodeList.length -1 do
          begin
            L_SeqNode := L_SeqNodeList.item[kk];
            sL_SeqNo := L_SeqNode.text;

            //down, 檢查序號格式
            if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
            begin
              //若執行中有問題,則Rollback
              DbConnAryRollBack;

              Result := '-2: 序號長度錯誤, 序號=<'+sL_SeqNo + '>';
              exit;
            end;
            //up, 檢查序號格式

                //down, 於原屬公司別檢查序號是否合理
                sL_Function := '3';//調撥

                if not checkSeqNo2(sL_Function,sL_OutCompCode,sL_EquipType, sL_SeqNo,sL_Msg, bL_EquipExist) then
                begin
                  //若執行中有問題,則Rollback
                  DbConnAryRollBack;

                  Result := sL_Msg;
                  exit;
                end;
                //up, 於原屬公司別檢查序號是否合理
                
            { 檢查調入的公司別是否有建立此帶代碼, 檢核調入的公司別即可 }
            { 1.先取調出公司別的 ModelNo }

            if ( sL_EquipType = '1' ) then
            begin
              aMsg := GetModelNo( sL_OutCompCode, sL_SeqNo, aModelNo );
              if ( aMsg <> EmptyStr ) then
              begin
                DbConnAryRollBack;
                Result := '-2:CC&B資料庫中此設備機種型號為空值, 料號=<'+sL_MaterialNo+'>';
                Exit;
              end;
            end;  

            aMsg := GetDescription( sL_EquipType, sL_OutCompCode, sL_SeqNo, aDescription );
            if ( aMsg <> EmptyStr ) then
            begin
              DbConnAryRollBack;
              Result := '-2:CC&B資料庫中此設備機種客服品名為空值, 料號=<'+sL_MaterialNo+'>';
              Exit;
            end;
            

            if ( sL_InCompCode <> IntToStr( nG_TOTAL_WAREHOUSE_ID ) ) then
            begin
              { 3. 檢核掉調入的公司別是否有建立此代碼, 調入為總倉則不檢核 }
              if ( sL_EquipType = '1' ) then
              begin
                aMsg := VdCD043( sL_InCompCode , aModelNo );
                if ( aMsg <> EmptyStr ) then
                begin
                  DbConnAryRollBack;
                  Result := '-2:CC&B資料庫中此設備查無對應機種型號,機種=<'+aModelNo+'> , 料號=<'+sL_MaterialNo+'>';
                  Exit;
                end;
              end;  
              { 4. 檢核掉調入的公司別是否有建立此客服品名, 調入為總倉則不檢核 }
              aMsg := VdCD022( sL_InCompCode , aDescription );
              if ( aMsg <> EmptyStr ) then
              begin
                DbConnAryRollBack;
                Result := '-2:CC&B資料庫中此設備查無對應客服品名,客服品名=<'+aDescription+'> , 料號=<'+sL_MaterialNo+'>';
                Exit;
              end;
            end;

            //依不同狀況更新資料
            //sL_Status
            if (sL_OutCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '0'; //代表新舊公司別相同,不做任何動作
            end
            else if (sL_OutCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
               (sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '1'; //總倉設備調撥至公司別A
            end
            else if (sL_OutCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '2'; //公司別A設備調撥至公司別B
            end
            else if (sL_OutCompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID)) and
                    (sL_InCompCode = IntToStr(nG_TOTAL_WAREHOUSE_ID)) then
            begin
              sL_Status := '3'; //公司別A設備調撥至總倉
            end;


            if sL_Status <> '0' then   //代表新舊公司別相同,不做任何動作
            begin
              //更新調撥資料

              allocateEquip(sL_Status,sL_OutCompCode,sL_InCompCode,sL_MaterialNo,
                            sL_EquipType,sL_SeqNo,sL_EquipTableName,sL_CurDateTime,sL_Msg);

            end;
          end;
        end;

        //Log一筆至共同資料區東富MSS呼叫記錄檔(SOAC0204)
        logDataToSOAC0204(sL_CurDateTime,'','E2C_DLL2',sL_Para);

        Result := '';
      end
      else
        Result := '-1:傳入之調撥 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;


      //若執行中沒有問題,則Commit
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
      //若執行中有問題,則Rollback
      DbConnAryRollBack;

      {
      for ii:=0 to G_InCompCodeStrList.Count-1 do
      begin
        nL_AreaDataNo := StrToInt(getAreaDataNo(G_InCompCodeStrList.Strings[ii]));
        G_DbConnAry[nL_AreaDataNo].RollbackTrans;
      end;
      }
    end;


    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面
    for ii:=0 to G_InCompCodeStrList.Count-1 do
      deallocateSpace(G_InCompCodeStrList.Strings[ii]);

end;



//處理驗退資料
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
      //將程式包在 Transaction 交易中
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].BeginTrans;

      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
      L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符

      //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
      sL_TmpXmlFileName := getRendomFileName;
      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

      sL_UpdEn := '物料系統';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin

        L_FirstNodeList := L_XMLDoc.getElementsByTagName('驗退');
        L_FirstNode := L_FirstNodeList.item[0];
        //取得驗退單號
        sL_BatchNo := TMyXML.getAttributeValue(L_FirstNode,'單號');

        //取得驗退公司別
        sL_CompCode := TMyXML.getAttributeValue(L_FirstNode,'公司代碼');
        //只能驗退在總倉的設備,故若傳入公司別 <> 總倉代碼則不能驗退
        if StrToInt(sL_CompCode) <> nG_TOTAL_WAREHOUSE_ID then
        begin
         if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
            G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

          Result := '-2: 僅能驗退總倉的設備,錯誤的驗退公司別=< ' + sL_CompCode + ' >';
          exit;
        end;



        //down, 檢查驗退單號格式
        if sL_BatchNo = '' then
        begin
          if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
            G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

          Result := '-2: 驗退單號須有值';
          exit;
        end;
        //up, 驗退檢查單號格式



        //檢查公司代碼是否存在於 CD039 中
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
        //UP, 檢查公司別代碼

  //改為連接ini內所有的公司別920927


        L_NodeList := L_XMLDoc.getElementsByTagName('設備');     //取得文件中之所有設備 tag
        for ii:=0 to L_NodeList.length -1 do
        begin

          L_EquipTypeNode := L_NodeList.item[ii];
          sL_EquipType := L_EquipTypeNode.childNodes.item[0].text;// 設備別

          //down, 檢查設備格式
          if not (StrToInt(sL_EquipType) in [1,2]) then
          begin
            if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
              G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

            Result := '-2: 此驗退之設備別不對, 設備別=<'+sL_EquipType+'>';
            exit;
          end;
          //up, 檢查設備格式


          L_SeqNodeList := L_EquipTypeNode.selectNodes('./序號');

          for kk:=0 to L_SeqNodeList.length -1 do
          begin
            L_SeqNode := L_SeqNodeList.item[kk];
            sL_SeqNo := L_SeqNode.text;


            //down, 檢查序號格式
            if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 此序號長度錯誤, 序號=<'+sL_SeqNo + '>';
              exit;
            end;
            //up, 檢查序號格式

            case StrToInt(sL_EquipType) of
              1: //為 STB
               begin
                sL_EquipTableName := 'SOAC0201A';
                sL_TempTableName := 'SOAC0203A';
                sL_ErrStr1 := '-4: CC&B資料庫中此STB不存在, 無法註銷/驗退, 序號=<'+sL_SeqNo+'>';
               end;
              2: //為 ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
                sL_TempTableName := 'SOAC0203B';
                sL_ErrStr1 := '-4: CC&B資料庫中此ICC不存在, 無法註銷/驗退, 序號=<'+sL_SeqNo+'>';
               end;
            end;

            //檢查該設備序號是否存在於總倉,且公司別不等於總倉代碼
            sL_Error := isNotWareHouseCode(sL_EquipTableName,sL_SeqNo,sL_ReturnCompCode);
            if sL_Error <> '' then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              if sL_Error = ERROR_STRING then
                Result := '-2: 此序號=<' + sL_SeqNo + '>設備已調撥至公司別=<' + sL_ReturnCompCode + '> 不能驗退'
              else
                Result := sL_Error;
              exit;
            end;

            //於共同資料區Table檢查該設備序號是否存在 down
            if not checkEquipIsExist(IntToStr(nG_TOTAL_WAREHOUSE_ID),sL_EquipType,sL_SeqNo,sL_ErrStr1,sL_Msg) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := sL_Msg;
              exit;
            end
            else
            begin
              //若所屬公司別不為總倉公司別,則檢查所屬公司別該設備是否存在
              if sL_CompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
              begin
                //down, 檢查序號是否合理
                sL_Function := '2';//驗退
                if not checkSeqNo2(sL_Function,sL_CompCode,sL_EquipType, sL_SeqNo,sL_Msg, bL_EquipExist) then
                begin
                  if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                    G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

                  Result := sL_Msg;
                  exit;
                end;
                //up, 檢查序號是否合理
              end;
            end;


            //down, 新增至總倉註銷備分檔及刪除CC&B的相關設備資料檔
            //載入庫資料檔取出該序號之資料 ( SOAC0201A 或 SOAC0201B )
            sL_SQL := 'SELECT * FROM ' + sL_EquipTableName + ' WHERE FACISNO=' +
                      STR_SEP + sL_SeqNo + STR_SEP;

            L_AdoQry := runSQL(SELECT_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);
            if L_AdoQry=nil then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + IntToStr(nG_TOTAL_WAREHOUSE_ID);
              Exit;
            end
            else
            begin
              dL_ExecTime := L_AdoQry.FieldByName('EXECTIME').AsDateTime;

              //原批號 = 入庫批號
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

              //Log一筆資料到共同資料區設備註銷備份檔
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
                          ',' + STR_SEP + '物料系統' + STR_SEP +
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
                          ',' + STR_SEP + '物料系統' + STR_SEP + ')';
              end;
              runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);
            end;

            //刪除共同資料區CC&B的相關設備資料檔
            sL_SQL := 'DELETE ' + sL_EquipTableName + ' WHERE FACISNO=' +
                      STR_SEP + sL_SeqNo + STR_SEP;
            runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);

            //若設備序號所屬公司別非總倉公司別,尚須再刪除該公司別的CC&B的相關設備資料檔
            if sL_CompCode <> IntToStr(nG_TOTAL_WAREHOUSE_ID) then
              runSQL(IUD_MODE,sL_CompCode,sL_SQL);

            //刪除共同資料區該序號對應之STB/ICC配對檔 ( SOAC0201C )
            deletePairData2(sL_SeqNo);
          end;

        end;

        //Log一筆至共同資料區東富MSS呼叫記錄檔(SOAC0204)
        sL_Para := '驗退單號=' + sL_BatchNo + ' ,公司代碼=' + sL_CompCode;
        logDataToSOAC0204(sL_CurDateTime,sL_CompCode,'E2C_DLL2',sL_Para);
        Result := '';
      end
      else
        Result := '-1:傳入之驗退 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;

      //若執行中沒有問題,則Commit
      G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].CommitTrans;
    except
      //若執行中有問題,則Rollback
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
      result := '-1: 讀取資料庫連線參數檔<'+sL_IniFileName +'>失敗';
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);
    nL_TotalDataArea := L_IniFile.ReadInteger('SYSINFO','TOTALDATAAREASCOUNT',-1);
    if nL_TotalDataArea=-1 then
    begin
      if FileExists(sL_IniFileName) then
        DeleteFile(sL_IniFileName);
      result := '-1: 資料庫連線參數檔之資料區總數設定錯誤!.請檢查 E2C_DLL.ini 之設定.';
      exit;
    end;

    sG_TOTAL_WAREHOUSE_ID := L_IniFile.ReadString('SYSINFO','TOTAL_WAREHOUSE_ID','0');
    //判定是否為數字
    if StrToIntDef(sG_TOTAL_WAREHOUSE_ID,ERROR_TOTAL_WAREHOUSE_ID) = ERROR_TOTAL_WAREHOUSE_ID then
    begin
      result := '-1: 資料庫連線參數檔之總倉代碼設定錯誤!<'+sG_TOTAL_WAREHOUSE_ID +'>';
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
      result := '-1: 資料庫連線參數檔之資料區設定錯誤!<'+sL_IniFileName +'>';
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
      1: //為 STB
       begin
        sL_EquipTableName := 'SOAC0201A';
        sL_ErrStr := '-4: CC&B資料庫中此STB已使用中, 無法重新入庫, 序號=<' + sI_SeqNo + '>'
       end;
      2: //為 ICC
       begin
        sL_EquipTableName := 'SOAC0201B';
        sL_ErrStr := '-4: CC&B資料庫中此ICC已使用中, 無法重新入庫, 序號=<' + sI_SeqNo + '>';
       end;
    end;

    sL_SQL := 'select * from '+ sL_EquipTableName +' where  FACISNO='+ STR_SEP + sI_SeqNo + STR_SEP;
    //showmessage(sL_SQL);
    L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
    if L_AdoQry=nil then
    begin
      sI_Msg := NO_AVAILABLE_QUERY_COMP + '  公司別=' + sI_CompID;
      result := false;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount>0 then //表示此設備已經存在=>檢查是否仍在使用中
      begin
        bI_EquipExist := true;
        if (sI_CompID <> IntToStr(nG_TOTAL_WAREHOUSE_ID))then   //by Howard.. => 從總倉調出則不需check
        begin
          sL_SQL := 'select PRDate from SO004 where FaciSNo=' + STR_SEP + sI_SeqNo+STR_SEP ;
          L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
          if L_AdoQry=nil then
          begin
            sI_Msg := NO_AVAILABLE_QUERY_COMP + '  公司別=' + sI_CompID;;
            result := false;
            Exit;
          end
          else
          begin
            if (L_AdoQry.RecordCount>0) and (L_AdoQry.FieldByName('PRDATE').AsString='') then  //若無拆除日期,表示此設備尚未被拆除=>表示此設備仍在使用中
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
    //sI_Function = 2  處理驗退
    //sI_Function = 3  處理調撥

    sI_Msg := '';
    bI_EquipExist := false;
    case StrToInt(sI_EquipNo) of
      1: //為 STB
       begin
         sL_EquipTableName := 'SOAC0201A';

         if sI_Function = '2' then
         begin
           sL_ErrStr1 := '-4: CC&B資料庫中此STB不存在, 無法註銷/驗退, 序號=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B資料庫中此STB已使用中, 無法註銷/驗退, 序號=<'+sI_SeqNo+'>';
         end
         else if sI_Function = '3' then
         begin
           sL_ErrStr1 := '-4: CC&B資料庫無此STB序號,序號=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B資料庫中此STB已使用中, 無法調撥, 序號=<'+sI_SeqNo+'>';
         end;
       end;
      2: //為 ICC
       begin
         sL_EquipTableName := 'SOAC0201B';
         if sI_Function = '2' then
         begin
           sL_ErrStr1 := '-4: CC&B資料庫中此ICC不存在, 無法註銷/驗退, 序號=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B資料庫中此ICC已使用中, 無法註銷/驗退, 序號=<'+sI_SeqNo+'>';
         end
         else if sI_Function = '3' then
         begin
           sL_ErrStr1 := '-4: CC&B資料庫無此ICC序號,序號=<'+sI_SeqNo+'>';
           sL_ErrStr2 := '-4: CC&B資料庫中此ICC已使用中, 無法調撥, 序號=<'+sI_SeqNo+'>';
         end;
       end;
    end;

    //sL_SQL := 'SELECT * FROM '+ sL_EquipTableName +' WHERE FACISNO='+ STR_SEP + sI_SeqNo + STR_SEP ;

    sL_SQL := Format( ' select * from %s where facisno = ''%s'' and compcode = ''%s''  ',
       [sL_EquipTableName, sI_SeqNo, sI_CompID] );

    L_AdoQry := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
    if L_AdoQry=nil then
    begin
      sI_Msg := NO_AVAILABLE_QUERY_COMP + '  公司別=' + sI_CompID;;
      result := false;
      Exit;
    end
    else
    begin
      if L_AdoQry.RecordCount > 0 then //表示此設備已經存在=>檢查是否仍在使用中
      begin


//jackal0726
        bI_EquipExist := true;
        if (sI_CompID <> IntToStr(nG_TOTAL_WAREHOUSE_ID))then   //by Howard.. => 從總倉調出則不需check
        begin
          sL_SQL := 'select PRDate from SO004 where FaciSNo=' + STR_SEP + sI_SeqNo+STR_SEP ;
          L_AdoQry1 := runSQL(SELECT_MODE, sI_CompID, sL_SQL);
          if L_AdoQry1=nil then
          begin
            sI_Msg := NO_AVAILABLE_QUERY_COMP + '  公司別=' + sI_CompID;;
            result := false;
            Exit;
          end
          else
          begin
            if (L_AdoQry1.RecordCount>0) and (L_AdoQry1.FieldByName('PRDATE').AsString='') then  //若無拆除日期,表示此設備尚未被拆除=>表示此設備仍在使用中
            begin
              L_AdoQry1.Close;
              sI_Msg := sL_ErrStr2;
              result := false;
              Exit;
            end;
          end;
        end;

      end
      else  //表示此設備不存在
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
    sI_Msg := NO_AVAILABLE_QUERY_COMP + '  公司別=' + sI_CompCode;
    Result := false;
  end
  else
  begin
    if L_AdoQry.RecordCount = 0 then
    begin
      sI_Msg := '-2: 此公司代碼錯誤, 代碼<' + sI_CompCode + '>';
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
  //視設備別, 修改共同資料區設備資料檔
  sL_SQL := 'UPDATE ' + sI_EquipTableName + ' SET COMPCODE=' + sI_InCompCode +
            ',MATERIALNO=' + STR_SEP + sI_MaterialNo + STR_SEP +
            ',UPDTIME=' + STR_SEP + sI_CurDateTime + STR_SEP +
            ',UPDEN=' + STR_SEP + '物料系統' + STR_SEP +
            ' WHERE FACISNO=' + STR_SEP + sI_SeqNo + STR_SEP;
            
  runSQL(IUD_MODE,IntToStr(nG_TOTAL_WAREHOUSE_ID), sL_SQL);


  if sI_Status = '1' then  //總倉設備調撥至公司別A
  begin

    //取出原設備所有資料
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
    //視設備別, 新增設備至調入公司別設備資料檔
    insertEquipToCompany(sI_EquipTableName,sI_InCompCode,sL_BatchNo,sL_ExecTime,
                         sI_MaterialNo,sI_SeqNo,sL_Description,sL_Spec,
                         sL_VendorCode,sL_VendorName,sI_CurDateTime, aModeFlag, aModelNo,
                         aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize);
  end
  else if sI_Status = '2' then  //公司別A設備調撥至公司別B
  begin
    //取出原設備所有資料
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

    //刪除設備原有公司的資料
    deleteOldEquipData(sI_EquipTableName,sI_OldCompCode,sI_SeqNo);

    //視設備別, 新增設備至調入公司別設備資料檔
    insertEquipToCompany(sI_EquipTableName,sI_InCompCode,sL_BatchNo,sL_ExecTime,
                         sI_MaterialNo,sI_SeqNo,sL_Description,sL_Spec,
                         sL_VendorCode,sL_VendorName,sI_CurDateTime, aModeFlag, aModelNo,
                         aMaterialType, aHwVer, aCallbackMode, aDirection, aHdSize);


  end
  else if sI_Status = '3' then  //公司別A設備調撥至總倉
  begin
    //刪除設備原有公司的資料
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
    sI_Msg := NO_AVAILABLE_QUERY_COMP  + '  公司別=' + sI_OldCompCode;
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
              ',' + STR_SEP + '物料系統' + STR_SEP +
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
              ',' + STR_SEP + '物料系統' + STR_SEP + ',' + '0' + ')';
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
  //Log一筆至共同資料區東富MSS呼叫記錄檔(SOAC0204)
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
    1: //為 STB
     begin
      sL_EquipTableName := 'SOAC0201A';
     end;
    2: //為 ICC
     begin
      sL_EquipTableName := 'SOAC0201B';
     end;
  end;


  //於共同資料區Table檢查該設備序號是否存在 down
  sL_SQL :='SELECT * FROM ' + sL_EquipTableName +
           ' WHERE FACISNO=' + STR_SEP + sI_SeqNo + STR_SEP;
  L_AdoQry := runSQL(SELECT_MODE, sI_CompCode, sL_SQL);
  if L_AdoQry=nil then
  begin
    sI_Msg := NO_AVAILABLE_QUERY_COMP  + '  公司別=' + sI_CompCode;;
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
    //刪除解密後的ini檔案
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
      Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + IntToStr(nG_TOTAL_WAREHOUSE_ID);
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


procedure TE2C.E2C_DLL3(vI_Data1: OleVariant);//領退料之function...
var
    sL_Result, sL_Data1 : String;
    sL_DbConnStr: String;
begin
    setDateTimeFormat;
    sL_Data1 := VarToStr(vI_Data1);
    //將加密的ini檔案解密
    TransTmpIniFile;

    sL_Result := '';

    if (UpperCase(trim(sL_Data1))<>XML_NO_DATA) and (Trim(sL_Data1)<>'')then
    begin
      createDbInfoStrList;
      //down, 從 ini 檔讀取連線至 DB 的 information
      sL_Result := getDbConnInfo(G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList);
      if sL_Result<>'' then
      begin
        Set_E2C_DLL3_RESULT(sL_Result);
        exit;
      end;
      //up, 從 ini 檔讀取連線至 DB 的 information
    end;

    //down, 處理領退料..
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
    //up, 處理領退料..

    //刪除解密後的ini檔案
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

    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面
    
    try


      bL_HasUpdateData := false;
      nL_ActiveCompCode := -1;
      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
      L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符

      //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
      sL_TmpXmlFileName := getRendomFileName;

      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

      sL_UpdEn := '物料系統';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

//      sL_TmpXmlFileName := 'c:\CS11448.xml';
      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin


        
        L_NodeList := L_XMLDoc.getElementsByTagName('批號');     //取得文件中之所有批號 tag

        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// 批號
          //down, 檢查批號格式
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            Result := '-2: 批號長度錯誤, 批號=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, 檢查批號格式

          sL_Operation := TMyXML.getSingleNodeText(L_BatchNoNode, '作業'); // 1:領料, 2:退料
          if (sL_Operation<>'1') and (sL_Operation<>'2') then
          begin
            Result := '-6: 領退料作業代碼錯誤.作業代碼:'+ sL_Operation;
            exit;
          end;

          if sL_Operation = '1' then    //領料
          begin
            sL_UseFlagValue := '1';
            sL_UseFlag := '領料序號: ';
          end
          else
          begin
            sL_UseFlagValue := '0';     //退料
            sL_UseFlag := '退料序號: ';
          end;

          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '日期');

          //down, 檢查日期格式
          if length(sL_Date)<>DATE_LENGTH then
          begin
            Result := '-2: 日期長度錯誤, 日期=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            Result := '-2: 執行日期格式不對: <' + sL_Date + '>';
            exit;
          end;
          //up, 檢查日期格式

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '時間');
          //down, 檢查時間格式
          if length(sL_Time)<>TIME_LENGTH then
          begin
            Result := '-2: 時間長度錯誤, 時間=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            Result := '-2: 執行時間格式不對: <' + sL_Time + '>';
            exit;
          end;
          //up, 檢查時間格式
          
          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '公司代碼');
//saveToFile(sL_CompID, 'c:\abc.txt');

          nL_ActiveCompCode := StrToIntDef(sL_CompID,-1);
          sL_ConnDbResult := connecttodb(not CONN_TO_WAREHOUSE, sL_CompID);
          if sL_ConnDbResult<>'' then
          begin
            Result := '-6: 無法連接上資料庫'+'-' + sL_ConnDbResult;
            exit;

          end;
//saveToFile('connect db ok', 'c:\aa1.txt');
          //howard123
          if nL_ActiveCompCode = -1 then
          begin
            Result := '-6: 領退料公司代碼錯誤(領退料公司代碼:' + sL_CompID + ')';
            exit;
          end;
//SaveToFile(inttostr(nL_ActiveCompCode) + '-' + inttostr(nG_TOTAL_WAREHOUSE_ID), 'c:\zzz12.txt');
          if nL_ActiveCompCode= nG_TOTAL_WAREHOUSE_ID then
          begin
            Result := '-6: 領退料公司代碼不應為總倉代碼';
            exit;
          end;
          //將程式包在 Transaction 交易中
          G_DbConnAry[0].BeginTrans;
//SaveToFile('abc', 'c:\zzz13.txt');


  //saveToFile(sL_BatchNo+sL_Date+sL_Time+sL_MakerID+sL_MakerName+sL_Status+sL_CompID, 'c:\aa2.txt');
          L_MaterialNodeList := L_BatchNoNode.selectNodes('./料號');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// 料號
            //down, 檢查料號格式
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 料號長度錯誤, 料號=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, 檢查料號格式

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'數量');
            //down, 檢查數量格式
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 數量錯誤, 數量=<'+sL_Quantity+'>';
              exit;
            end;
            //up, 檢查數量格式

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode, '客服品名');
            if ( Length( sL_ItemName ) > 20 ) then
            begin
              if G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].InTransaction then
                G_DbConnAry[nG_TOTAL_WAREHOUSE_ID].RollbackTrans;

              Result := '-2: 客服品名長度最多20碼, 料號=<'+sL_MaterialNo+'>, 客服品名=<'+sL_ItemName+'>';
              Exit;
            end;

            { 4. 檢核該公司別是否有建立此客服品名 }
            aMsg := VdCD022( sL_CompID , sL_ItemName );
            if ( aMsg <> EmptyStr ) then
            begin
              DbConnAryRollBack;
              Result := '-2:CC&B資料庫中此設備查無對應客服品名,客服品名=<'+sL_ItemName+'> , 料號=<'+sL_MaterialNo+'>';
              Exit;
            end;

            

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'設備');
            //down, 檢查設備格式
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 此料號之設備別不對, 料號=<'+sL_MaterialNo+'>, 設備別=<'+sL_Equip+'>';
              exit;
            end;
            //up, 檢查設備格式


            { 只有 STB 才有 MODEFLAG }
            aModeFlag := EmptyStr;
            if ( sL_Equip = '1' ) then
            begin
              aModeFlag := TMyXML.getAttributeValue(L_MaterialNode,'型式' );
              if ( Trim( aModeFlag ) <> '0' ) and ( Trim( aModeFlag ) <> '1'  ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 此料號之型式不對, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aModelNo := TMyXML.getAttributeValue(L_MaterialNode,'機種' );
              if ( Trim( aModelNo ) = EmptyStr ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 此料號之機種為空值, 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {}
              aMsg := VdCD043( sL_CompID, aModelNo );
              if ( aMsg <> EmptyStr ) then
              begin
                if G_DbConnAry[0].InTransaction then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 此設備查無對應機種型號,機種=<'+aModelNo+'> , 料號=<'+sL_MaterialNo+'>, 設備別=<'+aModeFlag+'>';
                Exit;
              end;
              {} 
            end;

            L_SeqNodeList := L_MaterialNode.selectNodes('./序號');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 此料號之數量與明細筆數不合, 料號=<'+sL_MaterialNo+'>, 數量=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
  //showmessage('dL_ExecDateTime  ' + DateTimeToStr(dL_ExecDateTime));
            case StrToInt(sL_Equip) of
              1: //為 STB
               begin
                sL_EquipTableName := 'SOAC0201A';
               end;
              2: //為 ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
               end;
            end;



            sL_SeqNoString := '';   //領退料LOG用
            for kk:=0 to L_SeqNodeList.length -1 do
            begin
            
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, 檢查序號格式
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 序號長度錯誤, 序號=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, 檢查序號格式



              //down, 檢查此序號是否已存在

              sL_SQL := 'SELECT COUNT(*) COUNT FROM ' + sL_EquipTableName;
              sL_SQL := sL_SQL + ' where  FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;

              L_AdoQry := runSQL(SELECT_MODE,sL_CompID, sL_SQL);


              if L_AdoQry.FieldByName('COUNT').AsInteger = 0 then
              begin

                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 此設備序號不存在, 序號=<'+sL_SeqNo+'>';
                exit;
              end;

              //up, 檢查此序號是否已存在


              //down, 更新CC&B的相關設備資料檔
              //howard..

              sL_SQL := 'update '+ sL_EquipTableName +' set UseFlag=' + sL_UseFlagValue;
              if ( sL_Equip = '1' ) then
              begin
                sL_SQL := sL_SQL + ',MODEFLAG=' + STR_SEP + aModeFlag + STR_SEP + ', MODELNO=' + STR_SEP + aModelNo + STR_SEP;

              end;
              sL_SQL := sL_SQL + ',UPDEN=' + STR_SEP + '物料系統' + STR_SEP ;
              sL_SQL := sL_SQL + ',UpdTime=' + STR_SEP + sL_CurDateTime + STR_SEP ;
              sL_SQL := sL_SQL + ' where FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;


              try
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              finally
                bL_HasUpdateData := true;
              end;
              //up, 更新CC&B的相關設備資料檔


              //串起領退料字串
              if sL_SeqNoString = '' then
                sL_SeqNoString := sL_SeqNo
              else
                sL_SeqNoString := sL_SeqNoString + ',' + sL_SeqNo;

            end;


            sL_SeqNoString := sL_UseFlag + sL_SeqNoString;
            //Log 領退料資料至 SOAC0204
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
        Result := '-1:傳入之領料 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;
        Exit;
      end;


        Result := EmptyStr;
      //若執行中沒有問題,則Commit

      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].CommitTrans;
    except
      //若執行中有問題,則Rollback
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

    //因寫在 Try 內會 AutoCommit , 所以寫在 Transaction 外面

    
    try

      bL_HasUpdateData := false;
      nL_ActiveCompCode := -1;
      L_XMLDoc := CoDOMDocument.Create;

      L_XMLDoc.async := False;//文件載入完成後,程式才會往下執行
      L_XMLDoc.ValidateOnParse := False;//是否需驗證文件與DTD是否相符

      //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
      sL_TmpXmlFileName := getRendomFileName;

      saveToFile(sI_Xml, sL_TmpXmlFileName);
      //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

      sL_UpdEn := '物料系統';
      sL_CurDateTime := TUdateTime.GetPureDateTimeStr(now);

//      sL_TmpXmlFileName := 'c:\CS11448.xml';
      if L_XMLDoc.load(sL_TmpXmlFileName) then
      begin


        
        L_NodeList := L_XMLDoc.getElementsByTagName('批號');     //取得文件中之所有批號 tag

        for ii:=0 to L_NodeList.length -1 do
        begin

          L_BatchNoNode := L_NodeList.item[ii];
          sL_BatchNo := L_BatchNoNode.childNodes.item[0].text;// 批號
          //down, 檢查批號格式
          if not (length(sL_BatchNo)<=BATCH_NO_LENGTH) then
          begin
            Result := '-2: 批號長度錯誤, 批號=<'+sL_BatchNo + '>';
            exit;
          end;
          //up, 檢查批號格式

          sL_Operation := TMyXML.getSingleNodeText(L_BatchNoNode, '作業'); // 1:領料, 2:退料
          if (sL_Operation<>'1') and (sL_Operation<>'2') then
          begin
            Result := '-6: 領退料作業代碼錯誤.作業代碼:'+ sL_Operation;
            exit;
          end;

          if sL_Operation = '1' then
            sL_UseFlagValue := '1' //領料
          else
            sL_UseFlagValue := '0'; //退料
            
          sL_Date := TMyXML.getSingleNodeText(L_BatchNoNode, '日期');

          //down, 檢查日期格式
          if length(sL_Date)<>DATE_LENGTH then
          begin
            Result := '-2: 日期長度錯誤, 日期=<'+sL_Date + '>';
            exit;
          end;

          if not isValidDateStr(sL_Date) then
          begin
            Result := '-2: 執行日期格式不對: <' + sL_Date + '>';
            exit;
          end;
          //up, 檢查日期格式

          sL_Time := TMyXML.getSingleNodeText(L_BatchNoNode, '時間');
          //down, 檢查時間格式
          if length(sL_Time)<>TIME_LENGTH then
          begin
            Result := '-2: 時間長度錯誤, 時間=<'+sL_Time + '>';
            exit;
          end;

          if not isValidTimeStr(sL_Time) then
          begin
            Result := '-2: 執行時間格式不對: <' + sL_Time + '>';
            exit;
          end;
          //up, 檢查時間格式
          
          sL_CompID := TMyXML.getSingleNodeText(L_BatchNoNode, '公司代碼');
//saveToFile(sL_CompID, 'c:\abc.txt');

          nL_ActiveCompCode := StrToIntDef(sL_CompID,-1);
          sL_ConnDbResult := connecttodb(not CONN_TO_WAREHOUSE, sL_CompID);
          if sL_ConnDbResult<>'' then
          begin
            Result := '-6: 無法連接上資料庫';
            exit;

          end;
//saveToFile('connect db ok', 'c:\aa1.txt');
          //howard123
          if nL_ActiveCompCode = -1 then
          begin
            Result := '-6: 領退料公司代碼錯誤(領退料公司代碼:' + sL_CompID + ')';
            exit;
          end;
//SaveToFile(inttostr(nL_ActiveCompCode) + '-' + inttostr(nG_TOTAL_WAREHOUSE_ID), 'c:\zzz12.txt');
          if nL_ActiveCompCode= nG_TOTAL_WAREHOUSE_ID then
          begin
            Result := '-6: 領退料公司代碼不應為總倉代碼';
            exit;
          end;
          //將程式包在 Transaction 交易中
          G_DbConnAry[0].BeginTrans;
//SaveToFile('abc', 'c:\zzz13.txt');


  //saveToFile(sL_BatchNo+sL_Date+sL_Time+sL_MakerID+sL_MakerName+sL_Status+sL_CompID, 'c:\aa2.txt');
          L_MaterialNodeList := L_BatchNoNode.selectNodes('./料號');
          for jj:=0 to L_MaterialNodeList.length-1 do
          begin

            L_MaterialNode := L_MaterialNodeList.item[jj];
            sL_MaterialNo := L_MaterialNode.childNodes.item[0].text;// 料號
            //down, 檢查料號格式
            if not (length(sL_MaterialNo)<=MATERIAL_NO_LENGTH) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 料號長度錯誤, 料號=<'+sL_MaterialNo + '>';
              exit;
            end;
            //up, 檢查料號格式

            sL_Quantity := TMyXML.getAttributeValue(L_MaterialNode,'數量');
            //down, 檢查數量格式
            nL_Quantity := StrToIntDef(sL_Quantity,-1);
            if nL_Quantity<=0 then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 數量錯誤, 數量=<'+sL_Quantity+'>';
              exit;
            end;
            //up, 檢查數量格式

            sL_ItemName := TMyXML.getAttributeValue(L_MaterialNode,'品名');

            sL_Equip := TMyXML.getAttributeValue(L_MaterialNode,'設備');
            //down, 檢查設備格式
            if not (StrToInt(sL_Equip) in [1,2]) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 此料號之設備別不對, 料號=<'+sL_MaterialNo+'>, 設備別=<'+sL_Equip+'>';
              exit;
            end;
            //up, 檢查設備格式


  //          showmessage(sL_MaterialNo+sL_Quantity + sL_ItemName + sL_Spec + sL_Equip);
  //saveToFile(sL_MaterialNo+sL_Quantity + sL_ItemName + sL_Spec + sL_Equip, 'c:\aa3.txt');

            L_SeqNodeList := L_MaterialNode.selectNodes('./序號');
            if L_SeqNodeList.length<>StrToInt(sL_Quantity) then
            begin
              if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                G_DbConnAry[0].RollbackTrans;

              Result := '-2: 此料號之數量與明細筆數不合, 料號=<'+sL_MaterialNo+'>, 數量=<'+sL_Quantity+'>';
              Exit;
            end;

            dL_ExecDateTime := getExexDateTime(sL_Date, sL_Time);
  //showmessage('dL_ExecDateTime  ' + DateTimeToStr(dL_ExecDateTime));
            case StrToInt(sL_Equip) of
              1: //為 STB
               begin
                sL_EquipTableName := 'SOAC0201A';
               end;
              2: //為 ICC
               begin
                sL_EquipTableName := 'SOAC0201B';
               end;
            end;


            for kk:=0 to L_SeqNodeList.length -1 do
            begin
              L_SeqNode := L_SeqNodeList.item[kk];
              sL_SeqNo := L_SeqNode.text;

              //down, 檢查序號格式
              if not (length(sL_SeqNo)<=SEQ_NO_LENGTH) then
              begin
                if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
                  G_DbConnAry[0].RollbackTrans;

                Result := '-2: 序號長度錯誤, 序號=<'+sL_SeqNo + '>';
                exit;
              end;
              //up, 檢查序號格式




              //down, 更新CC&B的相關設備資料檔
              //howard..

              sL_SQL := 'update '+ sL_EquipTableName +' set UseFlag=' + sL_UseFlagValue;
              sL_SQL := sL_SQL + ',UPDEN=' + STR_SEP + '物料系統' + STR_SEP ;
              sL_SQL := sL_SQL + ' where  FACISNO=' + STR_SEP + sL_SeqNo + STR_SEP;
//SaveToFile(sL_SQL, 'c:\aaa11.txt');


              try
                runSQL(IUD_MODE,sL_CompID, sL_SQL);
              finally
                bL_HasUpdateData := true;
              end;  
              //up, 更新CC&B的相關設備資料檔

            end;
          end;
          //howard123
        end;

        Result := '';

      end
      else
        Result := '-1:傳入之領料 XML 格式有誤!請檢視: ' + sL_TmpXmlFileName;


        Result := '';
      //若執行中沒有問題,則Commit
//savetofile(inttostr(nL_ActiveCompCode),'c:\xxx.txt');

      if (bL_HasUpdateData) and (nL_ActiveCompCode<>-1) and (G_DbConnAry[0].Connected) and(G_DbConnAry[0].InTransaction) then
        G_DbConnAry[0].CommitTrans;
//savetofile('ok...','c:\xxx1.txt');
    except
      //若執行中有問題,則Rollback
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
    Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + aCompCode;
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
    Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + aCompCode;
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
    Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + aCompCode;
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
    Result := NO_AVAILABLE_QUERY_COMP + '  公司別=' + aCompCode;
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

