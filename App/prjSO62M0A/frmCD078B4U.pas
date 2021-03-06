unit frmCD078B4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, DBClient, ADODB, Menus, Buttons,
  frmCD078B1U,
  cbDBController, cbStyleController, cbDataLookup,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cxCurrencyEdit, cxCheckBox, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TOldValue = record
    InstCode: String;
    ServiceType: String;
    GroupNo: String;
    PRCode: String;
    RefNo: String;
    PrRefNo: String;
  end;

  TfrmCD078B4 = class(TForm)
    Panel1: TPanel;
    edtDML: TcxTextEdit;
    Bevel1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    Panel2: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    chkInterdepend: TcxCheckBox;
    InstCode: TDataLookup;
    ServiceType: TDataLookup;
    GroupNo: TDataLookup;
    edtWorkUnit: TcxCurrencyEdit;
    edtRefNo: TcxTextEdit;
    Label6: TLabel;
    PRCode: TDataLookup;
    Label7: TLabel;
    btnInstCode2: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure edtWorkUnitPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure InstCodeCodeNoPropertiesChange(Sender: TObject);
    procedure GroupNoCodeNoPropertiesChange(Sender: TObject);
    procedure InstCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure InstCodeCodeNamePropertiesChange(Sender: TObject);
    procedure GroupNoCodeNamePropertiesInitPopup(Sender: TObject);
    procedure InstCodeCodeNoExit(Sender: TObject);
    procedure PRCodeCodeNoPropertiesChange(Sender: TObject);
    procedure PRCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure edtRefNoPropertiesChange(Sender: TObject);
    procedure btnInstCode2Click(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataFrom: String;
    FDataSet: TClientDataSet;
    FLockInstCodeChange: Boolean;
    FReader: TADOQuery;
    FOldValue: TOldValue;
    FInterDataSet: TClientDataSet;
    FaciDataSet: TClientDataSet;
    FCD078DDataSet: TClientDataSet;
    FCD078ADataSet: TClientDataSet;
    FCD078CDataSet: TClientDataSet;
    FCD078JDataSet: TClientDataSet;
    FCD078BRecord: TCD078BRecord;
    FCD078B1DataSet: TClientDataSet;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    procedure LoadRelationItem;
    procedure UpdateCD078B1KeyValue;
    procedure AutoAppendOneRecord;
    procedure EmptyCD078B1;
    function GetPrCodeRefNo(const ACode: String): String;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    function EditorToData: Boolean;
    function VdMustInput: Boolean;
    function VdDuplicate: Boolean;
    function VdCanShowMultiInstCode2: Boolean;
    function AppendCD078D(const AInstCode: String): Boolean;
    function AppendCD078A(const AInstCode: String): Boolean;
    function AppendCD078C(const AInstCode: String): Boolean;
    function AppendCD078J(const AInstCode: String): Boolean;
    function IsEnableCD078A1: Boolean; overload;
    function IsEnableCD078A1(const ARefNo: String): Boolean; overload;
    function IsSpecialNoCheck: Boolean; overload;
    function IsSpecialNoCheck(const ARefNo: String): Boolean; overload;
  public
    { Public declarations }
    constructor Create(AMode: TDML; AKeyValue, ADataFrom: String; ADataSet: TClientDataSet); reintroduce;
    property FaciItemDataSet: TClientDataSet read FaciDataSet write FaciDataSet;
    property CD078DDataSet: TClientDataSet read FCD078DDataSet write FCD078DDataSet;
    property CD078ADataSet: TClientDataSet read FCD078ADataSet write FCD078ADataSet;
    property CD078CDataSet: TClientDataSet read FCD078CDataSet write FCD078CDataSet;
    property CD078B1DataSet: TClientDataSet read FCD078B1DataSet write FCD078B1DataSet;
    property CD078JDataSet: TClientDataSet read FCD078JDataSet write FCD078JDataSet;
  end;

var
  frmCD078B4: TfrmCD078B4;

implementation

uses cbUtilis, frmCD078B4_1U;

{$R *.dfm}

{ TfrmCD078B2 }

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B4.Create(AMode: TDML; AKeyValue, ADataFrom: String; ADataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := AMode;
  FKeyValue := AKeyValue;
  FDataFrom := ADataFrom;
  FDataSet := ADataSet;
  FInterDataSet := TClientDataSet.Create( nil );
  FInterDataSet.Data := FDataSet.Data;
  FReader := DBController.CodeReader;
  FOldValue.InstCode := EmptyStr;
  FOldValue.ServiceType := EmptyStr;
  FOldValue.GroupNo := EmptyStr;
  FOldValue.PRCode := EmptyStr;
  FOldValue.RefNo := EmptyStr;
  FOldValue.PrRefNo := EmptyStr;
  FLockInstCodeChange := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.FormCreate(Sender: TObject);
begin
  InstCode.Initializa;
  ServiceType.Initializa;
  GroupNo.Initializa;
  PRCode.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    Application.ProcessMessages;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.FormDestroy(Sender: TObject);
begin
  FReader := nil;
  FDataSet := nil;
  FaciDataSet := nil;
  InstCode.Finaliza;
  ServiceType.Finaliza;
  GroupNo.Finaliza;
  PRCode.Finaliza;
  FInterDataSet.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.ClearEditor;
begin
  InstCode.Clear;
  ServiceType.Clear;
  GroupNo.Clear;
  edtWorkUnit.Text := EmptyStr;
  edtRefNo.Clear;
  PRCode.Clear;
  chkInterdepend.Checked := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.DataToEditor;
begin
  InstCode.Value := FDataSet.FieldByName( 'INSTCODE' ).AsString;
  InstCode.ValueName := FDataSet.FieldByName( 'INSTNAME' ).AsString;
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  GroupNo.Value := FDataSet.FieldByName( 'GROUPNO' ).AsString;
  GroupNo.ValueName := FDataSet.FieldByName( 'GROUPNAME' ).AsString;
  edtWorkUnit.Value := FDataSet.FieldByName( 'WORKUNIT' ).AsFloat;
  edtRefNo.Text := FDataSet.FieldByName( 'REFNO' ).AsString;
  chkInterdepend.Checked := ( FDataSet.FieldByName( 'INTERDEPEND' ).AsInteger = 1 );
  PRCode.Value := FDataSet.FieldByName( 'PRCode' ).AsString;
  PRCode.ValueName := FDataSet.FieldByName( 'PRName' ).AsString;
  {}
  FOldValue.InstCode := InstCode.Value;
  FOldValue.ServiceType := ServiceType.Value;
  FOldValue.GroupNo := GroupNo.Value;
  FOldValue.PRCode := PRCode.Value;
  FOldValue.RefNo := edtRefNo.Text;
  FOldValue.PrRefNo := EmptyStr;
  if ( PRCode.Value <> EmptyStr ) then
    FOldValue.PrRefNo := GetPrCodeRefNo( PRCode.Value );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.EditorToData: Boolean;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FDataSet.FieldByName( 'NEWBPCODE' ).AsString := FDataFrom;
    end;
    FDataSet.FieldByName( 'INSTCODE' ).AsString := InstCode.Value;
    FDataSet.FieldByName( 'INSTNAME' ).AsString := InstCode.ValueName;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'GROUPNO' ).AsString := GroupNo.Value;
    FDataSet.FieldByName( 'GROUPNAME' ).AsString :=  GroupNo.ValueName;
    FDataSet.FieldByName( 'WORKUNIT' ).AsFloat := edtWorkUnit.Value;
    FDataSet.FieldByName( 'REFNO' ).AsString := edtRefNo.Text;
    FDataSet.FieldByName( 'INTERDEPEND' ).AsInteger := 0;
    if chkInterdepend.Checked then
      FDataSet.FieldByName( 'INTERDEPEND' ).AsInteger := 1;
    FDataSet.FieldByName( 'PRCODE' ).AsString := PRCode.Value;
    FDataSet.FieldByName( 'PRNAME' ).AsString := PRCode.ValueName;
    if ( not IsEnableCD078A1( FDataSet.FieldByName( 'REFNO' ).AsString ) ) then
      EmptyCD078B1;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '存檔失敗, 原因:%s。', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;


{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.VdMustInput: Boolean;
var
  aErrMsg: String;
  aControl: TWinControl;
begin
  aErrMsg := EmptyStr;
  aControl := nil;
  try
    if ( InstCode.Value = EmptyStr ) then
    begin
      aErrMsg := '請輸入【裝機類別】。';
      aControl := InstCode.CodeNo;
      Exit;
    end;
    if ( ServiceType.Value = EmptyStr ) then
    begin
      aErrMsg := '請輸入【服務類別】。';
      aControl := ServiceType.CodeNo;
      Exit;
    end;
    if ( GroupNo.Value = EmptyStr ) then
    begin
      aErrMsg := '請輸入【管派類別】。';
      aControl := GroupNo.CodeNo;
      Exit;
    end;
    if ( edtWorkUnit.Text = EmptyStr ) then
    begin
      aErrMsg := '請輸入【派工點數】。';
      aControl := edtWorkUnit;
      Exit;
    end;
    if ( edtRefNo.Text = EmptyStr ) then
    begin
      aErrMsg := '輸入【參考號】。';
      aControl := edtRefNo;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      WarningMsg( aErrMsg );
      if Assigned( aControl ) then
        if aControl.CanFocus then aControl.SetFocus;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.VdDuplicate: Boolean;
var
  aCurrentPrRefNo: String;
  aCompareCount: Integer;
begin
  Result := False;
  { 重覆輸入檢查, BPCode + InstCode + ServiceType }
  if ( FOldValue.InstCode <> InstCode.Value ) or
     ( FOldValue.ServiceType <> ServiceType.Value ) then
  begin
    FInterDataSet.Filtered := False;
    FInterDataSet.Filter := Format( 'INSTCODE=''%s'' and SERVICETYPE=''%s''',
      [InstCode.Value, ServiceType.Value] );
    FInterDataSet.Filtered := True;
    try
      if ( FInterDataSet.RecordCount >= 1 ) then
      begin
        WarningMsg( '此【裝機類別】已存在(重覆), 請重新輸入。' );
        if InstCode.CodeNo.CanFocusEx then InstCode.CodeNo.SetFocus;
        Exit;
      end;
    finally
      FInterDataSet.Filtered := False;
      FInterDataSet.Filter := EmptyStr;
    end;
  end;
  { 同一服務類別中所選定之派工類別的參考號不能重覆,
    Ex : 服務別=I , 參考號=1的派工類別不能超過1 筆 }
  if ( ( FOldValue.ServiceType <> ServiceType.Value ) or
       ( FOldValue.GroupNo <> GroupNo.Value ) or
       ( FOldValue.RefNo <> edtRefNo.Text ) ) then
  begin
    { 2008/5/28: TACO: 若參考號為 16 則不須檢核 }
    { 2009/10/13: Penny: 若參考號為 19 則不須檢核 }
    if ( not IsEnableCD078A1 ) and ( not IsSpecialNoCheck ) then
    begin
      FInterDataSet.Filtered := False;
      FInterDataSet.Filter := Format( 'SERVICETYPE=''%s'' and REFNO=''%s''',
        [ServiceType.Value, edtRefNo.Text] );
      FInterDataSet.Filtered := True;
      try
        if ( FInterDataSet.RecordCount >= 1 ) then
        begin
          WarningMsg( '此【服務類別】的裝機類別【參考號】已重覆, 請重新輸入。');
          if edtRefNo.CanFocusEx then edtRefNo.SetFocus;
          Exit;
        end;
      finally
        FInterDataSet.Filtered := False;
        FInterDataSet.Filter := EmptyStr;
      end;
    end;
  end;
  aCurrentPrRefNo := GetPrCodeRefNo( PrCode.Value );
    { 特殊規則, 若拆機類別參考號為 10,
      則檢核該對應的拆機類別不可重覆 }
  if ( PRCode.Value <> EmptyStr ) and ( aCurrentPrRefNo = '10' ) then
  begin
    aCompareCount := 1;
    if ( FOldValue.PRCode <> PRCode.Value ) then
      aCompareCount := 0;
    FInterDataSet.Filtered := False;
    FInterDataSet.Filter := Format( 'SERVICETYPE=''%s'' and PRCODE=''%s''',
      [ServiceType.Value, PRCode.Value] );
    FInterDataSet.Filtered := True;
    try
      if ( FInterDataSet.RecordCount > aCompareCount ) then
      begin
        WarningMsg( '此【裝機類別】下對應的【拆機類別】已重覆, 請重新輸入。');
        if PRCode.CodeNo.CanFocusEx then PRCode.CodeNo.SetFocus;
        Exit;
      end;
    finally
      FInterDataSet.Filtered := False;
      FInterDataSet.Filter := EmptyStr;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.VdCanShowMultiInstCode2: Boolean;
begin
  Result := False;
  if ( InstCode.Value = EmptyStr ) then
  begin
    WarningMsg( '請先設定【裝機類別】。' );
    if ( InstCode.CodeNo.CanFocusEx ) then InstCode.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '請先設定【服務類別】。' );
    if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      Exit;
  end;
  {}
  if ( not IsEnableCD078A1 ) then
  begin
    WarningMsg( '此【裝機類別】所屬參考號, 不可設定【可更換頻道】。' );
    if edtRefNo.CanFocusEx then edtRefNo.SetFocus;
    Exit;
  end;
  {}
  if ( not VdDuplicate ) then
  begin
    if InstCode.CodeNo.CanFocusEx then InstCode.CodeNo.SetFocus;
    Exit;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '結束(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.EditorStateChange;
begin
  case FMode of
    dmInsert:
      begin
        ClearEditor;
        InstCode.Enabled := True;
        ServiceType.Enabled := True;
        GroupNo.Enabled := False;
        edtWorkUnit.Enabled := True;
        edtRefNo.Enabled := True;
        chkInterdepend.Enabled := True;
        if ServiceType.CodeNo.CanFocus then ServiceType.CodeNo.SetFocus;
        PRCode.Enabled := True;
        btnInstCode2.Enabled := False;
      end;
    dmUpdate:
      begin
        InstCode.Enabled := False;
        ServiceType.Enabled := False;
        GroupNo.Enabled := False;
        edtWorkUnit.Enabled := True;
        edtRefNo.Enabled := True;
        chkInterdepend.Enabled := True;
        if edtWorkUnit.CanFocus then edtWorkUnit.SetFocus;
        PRCode.Enabled := True;
        { 參考號為 16.頻道恢復授權 才可以設定 可更換頻道 }
        btnInstCode2.Enabled := IsEnableCD078A1;
      end;
    dmBrowser:
      begin
        InstCode.Enabled := False;
        ServiceType.Enabled := False;
        GroupNo.Enabled := False;
        edtWorkUnit.Enabled := False;
        edtRefNo.Enabled := False;
        chkInterdepend.Enabled := False;
        if btnCancel.CanFocus then btnCancel.SetFocus;
        PRCode.Enabled := False;
        btnInstCode2.Enabled := False;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.GetPrCodeRefNo(const ACode: String): String;
begin
  FReader.Close;
  FReader.SQL.Text := Format(
    ' SELECT REFNO FROM %s.CD007 A ' +
    '  WHERE A.CODENO = ''%S''     ',
    [DBController.LoginInfo.DbAccount, ACode] );
  FReader.Open;
  Result := FReader.Fields[0].AsString;
  FReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.DataSetStateChange(const aMode: TDML): Boolean;

  { ------------------------------------------- }

  procedure DoSubDetailSave;
  begin
    if ( not FCD078B1DataSet.IsEmpty ) then
      UpdateCD078B1KeyValue;
    if ( IsEnableCD078A1 ) then
      AutoAppendOneRecord;
    frmCD078B1.CloneCD078B1Data;
  end;

  { ------------------------------------------- }

  procedure DoSubDetailAutoAppendTables;
  var
    aResult: Boolean;
  begin
    FCD078B1DataSet.First;
    while not FCD078B1DataSet.Eof do
    begin
      { CD078B 設定的 InstCode 會自動寫一筆到 CD078B1, 所以 CD078B1 的
        InstCode 若有在 CD078B1 找到的話, 不用再做一次, 只須要做不一樣的 }
      if ( FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString <> InstCode.Value ) then
      begin
        aResult := AppendCD078D( FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString );
        if aResult then AppendCD078A( FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString );
        if aResult then AppendCD078C( FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString );
      end;
      FCD078B1DataSet.Next;
    end;
  end;

  { ------------------------------------------- }


begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        ClearEditor;
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := AppendCD078D( InstCode.Value );
        if Result then Result := AppendCD078A( InstCode.Value );
        if Result then Result := AppendCD078C( InstCode.Value );
        if Result then Result := AppendCD078J( InstCode.Value );
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
            DoSubDetailSave;
            { 子項裝機類別, 也要自動新增到相關 Table }
            DoSubDetailAutoAppendTables;
            { 檢核用的 DataSet 要更新 }
            FInterDataSet.Data := FDataSet.Data;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
              Result := False;
              Exit;
            end;
          end;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.DataValidate: Boolean;
begin
  Result := VdMustInput;
  if not Result then Exit;
  {}
  Result := VdDuplicate;
  if not Result then Exit;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.AppendCD078D(const AInstCode: String): Boolean;
var
  aErrMsg: String;
begin
  FCD078BRecord.BPCode := FKeyValue;
  FCD078BRecord.InstCode := AInstCode;
  FCD078BRecord.ServiceType := ServiceType.Value;
  FCD078BRecord.FaciItem := EmptyStr;
  Result := DBController.AutoAppendCD078D( FCD078BRecord, FCD078DDataSet, aErrMsg );
  if not Result then ErrorMsg( aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.AppendCD078A(const AInstCode: String): Boolean;
var
  aErrMsg: String;
begin
  FCD078BRecord.BPCode := FKeyValue;
  FCD078BRecord.InstCode := AInstCode;
  FCD078BRecord.ServiceType := ServiceType.Value;
  { 這一行不能給值, 因為 Append 到 CD078D 時以經取值了 }
  //FCD078BRecord.FaciItem := EmptyStr;
  Result := DBController.AutoAppendCD078A( FCD078BRecord, aErrMsg );
  if not Result then ErrorMsg( aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.AppendCD078C(const AInstCode: String): Boolean;
var
  aErrMsg: String;
begin
  FCD078BRecord.BPCode := FKeyValue;
  FCD078BRecord.InstCode := AInstCode;
  FCD078BRecord.ServiceType := ServiceType.Value;
  { 這一行不能給值, 因為 Append 到 CD078D 時以經取值了 }
  //FCD078BRecord.FaciItem := EmptyStr;
  Result := DBController.AutoAppendCD078C( FCD078BRecord, FCD078CDataSet, aErrMsg );
  if not Result then ErrorMsg( aErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.IsEnableCD078A1: Boolean;
begin
  Result := IsEnableCD078A1( edtRefNo.Text );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.IsEnableCD078A1(const ARefNo: String): Boolean;
begin
  Result := ( ARefNo = '16' );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.IsSpecialNoCheck: Boolean;
begin
  Result := IsSpecialNoCheck( edtRefNo.Text );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.IsSpecialNoCheck(const ARefNo: String): Boolean;
begin
  Result := ( ARefNo = '19' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );

  if ( not FLockInstCodeChange ) then
  begin
    InstCode.Clear;
    GroupNo.Clear;
    edtWorkUnit.Value := 0;
    edtRefNo.Clear;
    PRCode.Clear;
    chkInterdepend.Checked := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.edtRefNoPropertiesChange(Sender: TObject);
begin
  btnInstCode2.Enabled := IsEnableCD078A1;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.edtWorkUnitPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if not Error then Exit;
  if VarToStrDef( DisplayValue, EmptyStr ) = EmptyStr then Exit;
  if ( TcxCurrencyEdit( Sender ).Properties.DecimalPlaces = 0 ) then
  begin
     ErrorText := Format( '輸入值有誤, 值必須介於%d~%d',
      [Trunc( TcxCurrencyEdit( Sender ).Properties.MinValue ),
       Trunc( TcxCurrencyEdit( Sender ).Properties.MaxValue )] );
  end else
  begin
     ErrorText := Format( '輸入值有誤, 值必須介於%5.2f~%5.2f',
      [TcxCurrencyEdit( Sender ).Properties.MinValue,
       TcxCurrencyEdit( Sender ).Properties.MaxValue] );
  end;
  WarningMsg( ErrorText );
  ErrorText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.GroupNoCodeNoPropertiesChange(Sender: TObject);
begin
  GroupNo.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.InstCodeCodeNamePropertiesInitPopup(Sender: TObject);
var
  aSql: String;
begin
  Screen.Cursor := crSQLWait;
  try
    aSql := Format(
      ' SELECT A.CODENO, A.DESCRIPTION    ' +
      '   FROM %s.CD005 A                 ' +
      '  WHERE A.STOPFLAG = 0             ',
      [DBController.LoginInfo.DbAccount] );
    {}
    if ( ServiceType.Value <> EmptyStr ) then
      aSql := aSql + Format( ' AND A.SERVICETYPE = ''%s'' ', [ServiceType.Value] );
    aSql := aSql + ' ORDER BY A.CODENO ';
    {}
    InstCode.SQL.Text := aSql;
    InstCode.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.InstCodeCodeNoPropertiesChange(Sender: TObject);
begin
  LoadRelationItem;
  InstCode.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.InstCodeCodeNamePropertiesChange(Sender: TObject);
begin
  InstCode.CodeNamePropertiesChange( Sender );
  LoadRelationItem;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  InstCode.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION   ' +
      '   FROM %s.CD046 A                ' +
      '  ORDER BY A.CODENO               ',
      [DBController.LoginInfo.DbAccount] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.GroupNoCodeNamePropertiesInitPopup(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    GroupNo.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION         ' +
      '   FROM %s.CD044 A, %s.CD005 B          ' +
      '  WHERE A.CODENO = B.GROUPNO            ' +
      '    AND A.SERVICETYPE = B.SERVICETYPE   ' +
      '    AND B.CODENO = ''%s''               ' +
      '    AND B.SERVICETYPE = ''%s''          ' +
      '  ORDER BY A.CODENO                     ',
      [DBController.LoginInfo.DbAccount, DBController.LoginInfo.DbAccount,
       InstCode.Value, ServiceType.Value] );
    GroupNo.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.LoadRelationItem;
begin
  FReader.Close;
  Screen.Cursor := crSQLWait;
  try
    FLockInstCodeChange := True;
    try
      FReader.SQL.Text := Format(
        ' SELECT A.CODENO,           ' +
        '        A.DESCRIPTION,      ' +
        '        A.SERVICETYPE,      ' +
        '        A.GROUPNO,          ' +
        '        A.WORKUNIT,         ' +
        '        A.REFNO,            ' +
        '        A.INTERDEPEND       ' +
        '   FROM %s.CD005 A          ' +
        '  WHERE A.CODENO = ''%s''   ' ,
        [DBController.LoginInfo.DbAccount, InstCode.value] );
      FReader.Open;
      if ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> EmptyStr ) and
         ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> ServiceType.Value ) then
      begin
        ServiceType.Value := FReader.FieldByName( 'SERVICETYPE' ).AsString;
        PRCode.Clear;
      end;
      GROUPNO.Value := FReader.FieldByName( 'GROUPNO' ).AsString;
      edtWorkUnit.Value := FReader.FieldByName( 'WORKUNIT' ).AsFloat;
      edtRefNo.Text := FReader.FieldByName( 'REFNO' ).AsString;
      chkInterdepend.Checked := FReader.FieldByName( 'INTERDEPEND' ).AsInteger = 1 ;
      FReader.Close;
    finally
      FLockInstCodeChange := False;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.UpdateCD078B1KeyValue;
begin
  FCD078B1DataSet.First;
  while not FCD078B1DataSet.Eof do
  begin
    FCD078B1DataSet.Edit;
    FCD078B1DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    FCD078B1DataSet.FieldByName( 'INSTCODE' ).AsString := InstCode.Value;
    FCD078B1DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FCD078B1DataSet.Post;
    FCD078B1DataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.AutoAppendOneRecord;
begin
  { 當本次裝機類別, 為必須輸入 可更換頻道 時, 在存檔時
    該筆 CD078B 輸入的資料, 必須自動帶一筆到 CD078B1 }
  if FCD078B1DataSet.Locate( 'BPCODE;INSTCODE;INSTCODE2', VarArrayOf(
    [FKeyValue,
     FDataSet.FieldByName( 'INSTCODE' ).AsString,
     FDataSet.FieldByName( 'INSTCODE' ).AsString] ), [] ) then
  begin
    { 只更新拆機類別 }
    FCD078B1DataSet.Edit;
    FCD078B1DataSet.FieldByName( 'PRCODE2' ).AsString := FDataSet.FieldByName( 'PRCODE' ).AsString;
    FCD078B1DataSet.FieldByName( 'PRNAME2' ).AsString := FDataSet.FieldByName( 'PRNAME' ).AsString;
    FCD078B1DataSet.Post;
  end else
  begin
    FCD078B1DataSet.First;
    FCD078B1DataSet.InsertRecord( [
      FKeyValue,
      FDataSet.FieldByName( 'INSTCODE' ).AsString,
      FDataSet.FieldByName( 'SERVICETYPE' ).AsString,
      FDataSet.FieldByName( 'INSTCODE' ).AsString,
      FDataSet.FieldByName( 'INSTNAME' ).AsString,
      FDataSet.FieldByName( 'PRCODE' ).AsString,
      FDataSet.FieldByName( 'PRNAME' ).AsString] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.EmptyCD078B1;
begin
  FCD078B1DataSet.First;
  while not FCD078B1DataSet.Eof do
    FCD078B1DataSet.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.InstCodeCodeNoExit(Sender: TObject);
begin
  InstCode.CodeNoExit( Sender );
  LoadRelationItem;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.PRCodeCodeNoPropertiesChange(Sender: TObject);
begin
  PRCode.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.PRCodeCodeNamePropertiesInitPopup(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    PRCode.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION,  ' +
      '        A.REFNO                   ' +
      '   FROM %s.CD007 A                ' +
      '  WHERE A.SERVICETYPE =''%s''     ' +
      '    AND A.STOPFLAG = 0            ' +
      '  ORDER BY A.CODENO               ',
      [DBController.LoginInfo.DbAccount, ServiceType.Value] );
    PRCode.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4.btnInstCode2Click(Sender: TObject);
var
  aKeys: array [1..4] of String;
begin
  if not VdCanShowMultiInstCode2 then Exit;
  {}
  aKeys[1] := FKeyValue;
  aKeys[2] := InstCode.Value;
  aKeys[3] := ServiceType.Value;
{}
  if ( FMode = dmInsert ) then
  begin
    FCD078B1DataSet.Filtered := False;
    FCD078B1DataSet.Filter := Format(
      'BPCODE=''%s'' AND INSTCODE=''%s'' AND SERVICETYPE=''%s''',
      [aKeys[1], aKeys[2], aKeys[3]] );
    FCD078B1DataSet.Filtered := True;
  end;
  {}
  frmCD078B4_1 := TfrmCD078B4_1.Create( FMode, FCD078B1DataSet, aKeys );
  try
    frmCD078B4_1.Caption := '派工參數子項設定[frmCD078B4_1]';
    frmCD078B4_1.ShowModal;
  finally
    frmCD078B4_1.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4.AppendCD078J(const AInstCode: String): Boolean;
  var aErrMsg : String;
begin
  FCD078BRecord.BPCode := FKeyValue;
  FCD078BRecord.InstCode := AInstCode;
  FCD078BRecord.ServiceType := ServiceType.Value;
  { 這一行不能給值, 因為 Append 到 CD078D 時以經取值了 }
  //FCD078BRecord.FaciItem := EmptyStr;
  Result := DBController.AutoAppendCD078J( FCD078BRecord, FCD078JDataSet, aErrMsg );
  if not Result then ErrorMsg( aErrMsg );
end;

end.
