unit UCommonU;

interface

uses
    SysUtils, comctrls, classes, UObjectu, dbctrls, stdctrls, controls, inifiles, Forms, DB,
    DBClient;

const
    COMBINE_CODE_AND_DESC = true;
    FIELD_TYPE_SEP = '$';

    DATE_PERIOD_SEP = '~';
    AA4440_AMT_KEY1 = '應收款';
    AA4440_AMT_KEY2 = '沖銷款';
    UNIT_STRING = '單位:新台幣(元)';
    ZERO_AMT = '0.00';
    UDP_PORT = 8090;
    USE_CHINESE_FORMAT = true;
    SHOW_DATE_HEADER = true;

    DO_NOT_CARE_COMP = '-1';
    HANG_DATA_FLAG = 1;
    NONE_HANG_DATA_FLAG = 0;
    NOT_LAST_REC_FLAG='N';
    LAST_REC_FLAG='L';
    VOU_REPORT_PER_PAGE_DATA_COUNT= 6;
    AMT_TAIL = '.00';
    NO_LAST_REPORT_DATA = '無上次統計結果!';
    NO_REPORT_NAME = '報表格式檔不存在! 檔名:';
    CODE_VALUE_SEP = '-';
    QUERY_NO_RECORD = '查無資料,請重新查詢!';
    VOUKEY_LENGTH = 12;
    STR_SEP = '''';
    EMPTY_DATE_STR = '    /  /  ';
    DEBIT_FLAG='D';
    CREDIT_FLAG='C';
    SYSTEM_CONST_VALUE = 'FFFF';
    DEFAULT_ALL_DATA_ITEM_CODE_VALUE = '999';
    DEFAULT_ALL_DATA_ITEM_DESC_VALUE = '所有';
    ACC_CODE_SOURCE = 'X';    
    ACC_CODE_TARGET = 'a';
    DEFAULT_ACCCODE_MASK = 'aaaa-aaa-aaa;0';
    TEMP_VOUKEY2='XXXXX';
    SYS_INI_FILE_NAME = 'act.ini';

    INSERT_ALL_DATA_ITEM = true;
    INNER_ACC=1;
    BOTH_ACC=2;
    OUTER_ACC=3;

    MODIFY_OPERATION = 101;
    DELETE_OPERATION = 102;
    INSERT_OPERATION = 103;    

type
  TUCommonFun = class
    public
      class procedure setWaitingCursor;
      class procedure setDefaultCursor;      
      class procedure AddObjectToComboBox(I_StrList:TStrings; I_DS:TDataSet; bI_InsertAllDataItem:boolean);overload;
      class procedure AddObjectToComboBox(I_StrList:TStrings; I_DS:TDataSet; bI_InsertAllDataItem, bI_CombineCD:boolean);overload;
      class procedure setComboDefaultNdx(I_Combo:TComboBox; sI_ID:String);
      class function getComboCode(I_Combo:TComboBox):String;      
      class function getSpecialAcc(I_CDS:TClientDataSet):TDataSet;
      
      class procedure GetVesselID(I_WinControl:TWinControl; sL_FieldName:String);
      class procedure GetVoyage(I_WinControl:TWinControl; sL_FieldName:String);
      class procedure GetJobNo(I_WinControl:TWinControl);
//      class procedure GetVoyage(var I_Comp:TComponent);
      class procedure SaveVesselID(I_WinControl:TWinControl);
      class procedure SaveVoyage(I_WinControl:TWinControl);
      class procedure SaveJobNo(I_WinControl:TWinControl);
      //down, 將各種幣別的值轉換成台幣
//      class function TransCoin(I_StrList:TStrings; sI_CoinID:String;  fI_Value:Double):Double;
      //up, 將各種幣別的值轉換成台幣

      //down, 依照各船期航次的匯率,將台幣轉換成美金
//      class function TransNTDollarsToUS(I_StrList:TStrings; fI_Value:Double):Double;
      //down, 依照各船期航次的匯率,將台幣轉換成美金

    private
      class function RWIniFiles(nI_RW:Integer; Section, sI_Ident, sI_Value: string):String;
    end;
var
   sC_OutTaxItemCD,sC_InTaxItemCD : String;
   sC_OutNTSIncomeItemCD, sC_InNTSIncomeItemCD : String;
   sC_OutNTSHCItemCD,sC_InNTSHCItemCD : String;
   sC_CashItemCD : String;
   sC_MajorItemCD : String;//乙存
   sC_SouldPayItemCD : String;//應付帳款
   sC_ShouldChargeItemCD : String;//應收帳款
   
   nC_IM : integer;//用以區分進出口(0:進口; 1:出口) 
   sC_ActiveVesselID : String;
   sC_ActiveVoyage : String;
   sC_ActiveJobNo : String;   
   sC_UserID : String;
   sC_UserName : String;
   sC_Password : String;
   sC_UserGrp : String;
   fC_CompTax : Double;

implementation

uses dtmMainU, Ustru;

{ TUCommonFun }


{ TUCommonFun }


{ TUCommonFun }


class procedure TUCommonFun.AddObjectToComboBox(I_StrList: TStrings;
  I_DS: TDataSet; bI_InsertAllDataItem: boolean);
var
    L_Obj : TNormalObj;
    sL_CodeNo, sL_Desc : String;
begin
    I_StrList.Clear;
    with I_DS do
    begin
      if state = dsInactive then
        active := true;
      first;
      if bI_InsertAllDataItem then
      begin
        L_Obj := TNormalObj.Create;
        sL_CodeNo := DEFAULT_ALL_DATA_ITEM_CODE_VALUE;
        sL_Desc := DEFAULT_ALL_DATA_ITEM_DESC_VALUE;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj);
      end;

      while not eof do
      begin
        sL_CodeNo := FieldByName('CodeNo').AsString;
        sL_Desc := FieldByName('Description').AsString;
        L_Obj := TNormalObj.Create;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj);
        Next;
      end;
    end;
end;


class procedure TUCommonFun.AddObjectToComboBox(I_StrList: TStrings;
  I_DS: TDataSet; bI_InsertAllDataItem, bI_CombineCD: boolean);
var
    L_Obj : TNormalObj;
    sL_CodeNo, sL_Desc : String;
begin
    //bI_CombineCD => 是否要將 code and description 結合在一起
    I_StrList.Clear;
    with I_DS do
    begin
      if state = dsInactive then
        active := true;
      first;
      if bI_InsertAllDataItem then
      begin
        L_Obj := TNormalObj.Create;
        sL_CodeNo := DEFAULT_ALL_DATA_ITEM_CODE_VALUE;
        sL_Desc := DEFAULT_ALL_DATA_ITEM_DESC_VALUE;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj);
      end;

      while not eof do
      begin
        sL_CodeNo := FieldByName('CodeNo').AsString;
        sL_Desc := FieldByName('Description').AsString;
        L_Obj := TNormalObj.Create;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        if bI_CombineCD then
          I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj)
        else
          I_StrList.AddObject(sL_Desc,L_Obj);          
        Next;
      end;
    end;
end;

class function TUCommonFun.getComboCode(I_Combo: TComboBox): String;
var
    sL_result : String;
begin
    sL_result := (I_Combo.Items.Objects[I_Combo.ItemIndex] as TNormalObj).s_Code;

    result := sL_result;
end;

class procedure TUCommonFun.GetJobNo(I_WinControl: TWinControl);
begin

   sC_ActiveJobNo := TUCommonFun.RWIniFiles(0,'RUNTIMEINFO','工作序號','');
   if I_WinControl IS TEdit then
     (I_WinControl AS TEdit).Text := sC_ActiveJobNo;
end;

class function TUCommonFun.getSpecialAcc(I_CDS:TClientDataSet): TDataSet;
var
    sL_SQL : String;
begin
    with I_CDS do
    begin
      Close;
      sL_SQL := 'select a.CODENO,a.DESCRIPTION,a.DECRCODE from AA007 a, AA006 b';
      sL_SQL := sL_SQL + ' where b.REFNO1=3 and b.CODENO=SubStr(a.CODENO,1,length(b.CODENO))';
      CommandText := sL_SQL;
      Open;
    end;
    result := I_CDS;
end;

class procedure TUCommonFun.GetVesselID(I_WinControl:TWinControl; sL_FieldName:String);
begin
    if not I_WinControl.visible then exit;
    if not I_WinControl.Enabled then exit;
    sC_ActiveVesselID := TUCommonFun.RWIniFiles(0,'RUNTIMEINFO','船舶代碼','');
    I_WinControl.SetFocus;
    if I_WinControl IS TDBLookupComboBox then
    begin
      (I_WinControl AS TDBLookupComboBox).DropDown;
      (I_WinControl AS TDBLookupComboBox).ListSource.DataSet.Locate(sL_FieldName,sC_ActiveVesselID,[]);
//     dblVessel.CloseUp(True);
    end
    else if I_WinControl IS TEdit then
      (I_WinControl AS TEdit).Text := sC_ActiveVesselID;
end;

class procedure TUCommonFun.GetVoyage(I_WinControl: TWinControl;
  sL_FieldName: String);
begin
    sC_ActiveVoyage := TUCommonFun.RWIniFiles(0,'RUNTIMEINFO','航次','');
    
    I_WinControl.SetFocus;
    if I_WinControl IS TDBLookupComboBox then
    begin
      (I_WinControl AS TDBLookupComboBox).DropDown;
      (I_WinControl AS TDBLookupComboBox).ListSource.DataSet.Locate(sL_FieldName,sC_ActiveVoyage,[]);
//     dblVessel.CloseUp(True);
    end
    else if I_WinControl IS TEdit then
      (I_WinControl AS TEdit).Text := sC_ActiveVoyage;
end;

class function TUCommonFun.RWIniFiles(nI_RW: Integer; Section, sI_Ident,
  sI_Value: string):String;
var
   L_IniFile : TIniFile;
   ExeFileName, ExePath : String;
begin
   ExeFileName := Application.ExeName;
   ExePath := ExtractFileDir(ExeFileName);

   L_IniFile := TIniFile.Create(ExePath + '\' + 'Carriage.ini');
   case nI_RW of
    0://read
      Result := L_IniFile.ReadString(Section,sI_Ident,sI_Value);
    1://write
      L_IniFile.WriteString(Section,sI_Ident,sI_Value);
   end;

   L_IniFile.Free;
end;

class procedure TUCommonFun.SaveJobNo(I_WinControl:TWinControl);
var
   sL_JobNo : String;

begin
    if I_WinControl IS TEdit then
      sL_JobNo := (I_WinControl AS TEdit).Text
    else if I_WinControl IS TDBEdit then
      sL_JobNo := (I_WinControl AS TDBEdit).Text;


    sC_ActiveJobNo := sL_JobNo;
    RWIniFiles(1,'RUNTIMEINFO','工作序號', sC_ActiveJobNo);
end;

class procedure TUCommonFun.SaveVesselID(I_WinControl:TWinControl);
var
    sL_VesselID : String;
begin
    if I_WinControl IS TDBLookupComboBox then
      sL_VesselID := (I_WinControl AS TDBLookupComboBox).Text
    else if I_WinControl IS TComboBox then
      sL_VesselID := (I_WinControl AS TComboBox).Text
    else if I_WinControl IS TEdit then
      sL_VesselID := (I_WinControl AS TEdit).Text
    else if I_WinControl IS TDBEdit then
      sL_VesselID := (I_WinControl AS TDBEdit).Text;

    sC_ActiveVesselID := sL_VesselID;

    RWIniFiles(1,'RUNTIMEINFO','船舶代碼', sC_ActiveVesselID);
end;

class procedure TUCommonFun.SaveVoyage(I_WinControl:TWinControl);
var
    sL_Voyage : String;
begin
    if I_WinControl IS TDBLookupComboBox then
      sL_Voyage := (I_WinControl AS TDBLookupComboBox).Text
    else if I_WinControl IS TComboBox then
      sL_Voyage := (I_WinControl AS TComboBox).Text
    else if I_WinControl IS TEdit then
      sL_Voyage := (I_WinControl AS TEdit).Text
    else if I_WinControl IS TDBEdit then
      sL_Voyage := (I_WinControl AS TDBEdit).Text;

    sC_ActiveVoyage := sL_Voyage;

    RWIniFiles(1,'RUNTIMEINFO','航次', sC_ActiveVoyage);
end;

{
class function TUCommonFun.TransCoin(I_StrList: TStrings;
  sI_CoinID: String; fI_Value: Double): Double;
var
   ii : Integer;
   fL_Result, fL_Rate : Double;
begin
     fL_Result := 0;
     for ii:=0 to I_StrList.Count-1 do
     begin
       if TCoinRate(I_StrList.Objects[ii]).sCoinID=sI_CoinID then
       begin
         fL_Rate := TCoinRate(I_StrList.Objects[ii]).fRate;
         fL_Result := fL_Rate*fI_Value;
         break;
       end;
     end;
     Result := fL_Result;
end;
}
{
class function TUCommonFun.TransNTDollarsToUS(I_StrList: TStrings;
  fI_Value: Double): Double;
var
   ii : Integer;
   fL_Result, fL_Rate : Double;
begin
     fL_Result := 0;
     for ii:=0 to I_StrList.Count-1 do
     begin
       if TCoinRate(I_StrList.Objects[ii]).sCoinID='C01' then//為美金者
       begin
         fL_Rate := TCoinRate(I_StrList.Objects[ii]).fRate;
         fL_Result := fI_Value/fL_Rate;
         break;
       end;
     end;
     Result := fL_Result;
end;
}
class procedure TUCommonFun.setComboDefaultNdx(I_Combo: TComboBox;
  sI_ID: String);
var
    ii : Integer;
    L_StrList : TStringList;
    sL_ComboValue, sL_Code, sL_Value : String;
begin
    for ii:=0 to I_Combo.Items.Count -1 do
    begin
      sL_ComboValue := UpperCase(I_Combo.Items[ii]);
      L_StrList := TUStr.ParseStrings(sL_ComboValue,CODE_VALUE_SEP);
      if L_StrList.Count=2 then
      begin
        sL_Code := L_StrList.Strings[0];
        sL_Value := L_StrList.Strings[1];
        if UpperCase(sL_Code)= UpperCase(sI_Id) then
        begin
          I_Combo.ItemIndex := ii;
          break;
        end;
      end;
    end;
end;

class procedure TUCommonFun.setDefaultCursor;
begin
    Screen.Cursor := crDefault;
end;

class procedure TUCommonFun.setWaitingCursor;
begin
    Screen.Cursor := crSQLWait;
end;

end.

