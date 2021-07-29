unit frmCD078B5U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, DBClient, ADODB, Menus,
  frmCD078B0U, frmCD078B1U,
  cbDBController, cbStyleController,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cbDataLookup, cxCurrencyEdit, cxCheckBox, cxLookAndFeelPainters,
  cxButtons, dxSkinsCore, dxSkinsDefaultPainters, cxGraphics, dxSkinscxPCPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;


type
  TfrmCD078B5 = class(TForm)
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
    Label5: TLabel;
    chkPunish: TcxCheckBox;
    Label4: TLabel;
    Label6: TLabel;
    btnInstCodeStr: TButton;
    Citem: TDataLookup;
    ServiceType: TDataLookup;
    FaciItem: TDataLookup;
    Label9: TLabel;
    cxButton1: TcxButton;
    tmpCD078CDataSet: TClientDataSet;
    edtAmount: TcxCurrencyEdit;
    edtDiscountAmt: TcxCurrencyEdit;
    cmbPenalType: TcxComboBox;
    edtInstCodeStr: TcxTextEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CitemCodeNoExit(Sender: TObject);
    procedure CitemCodeNoPropertiesChange(Sender: TObject);
    procedure CitemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure CitemCodeNamePropertiesChange(Sender: TObject);
    procedure FaciItemCodeNoPropertiesChange(Sender: TObject);
    procedure FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure chkPunishPropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataSet: TClientDataSet;
    FReader: TADOQuery;
    FServiceType: String;
    FDataFrom: String;
    FaciDataSet: TClientDataSet;
    FInstCodeDataSet: TClientDataSet;
    FQueryParam: TQueryParam;
    FLockCItemChange: Boolean;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    procedure LoadRelationItem;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    function EditorToData: Boolean;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue, aService, aDataFrom: String;
      aDataSet: TClientDataSet); reintroduce;
    property FaciItemDataSet: TClientDataSet read FaciDataSet write FaciDataSet;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
  end;

var
  frmCD078B5: TfrmCD078B5;

implementation

uses cbUtilis, frmMultiSelectU;

{$R *.dfm}

{ TfrmCD078B2 }

{ ---------------------------------------------------------------------------- }

function QuotedServiceType(aService: String): String;
begin
  Result := EmptyStr;
  if ( aService <> EmptyStr ) then
  begin
    repeat
       Result := Result + QuotedStr( ExtractValue( aService ) ) + ',';
    until ( aService = EmptyStr );
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );  
end;

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B5.Create(aMode: TDML; aKeyValue, aService, aDataFrom: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FServiceType :=  QuotedServiceType( aService );
  FDataFrom := aDataFrom;
  FDataSet := aDataSet;
  FReader := DBController.CodeReader;
  FQueryParam.Param1 := EmptyStr;
  FQueryParam.Param2 := EmptyStr;
  FQueryParam.Param3 := EmptyStr;
  FQueryParam.Param4 := EmptyStr;
  FQueryParam.Param5 := EmptyStr;
  FQueryParam.Param6 := EmptyStr;          
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FormCreate(Sender: TObject);
begin
  Citem.Initializa;
  ServiceType.Initializa;
  FaciItem.Initializa;
  FLockCItemChange := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FaciItem.DataSet := FaciDataSet;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FormDestroy(Sender: TObject);
begin
  FReader := nil;
  FDataSet := nil;
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := EmptyStr;
  FaciDataSet := nil;
  Citem.Finaliza;
  ServiceType.Finaliza;
  FaciItem.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.LoadRelationItem;
begin
  FReader.Close;
  edtAmount.Value := 0;
  Screen.Cursor := crSQLWait;
  try
    FReader.SQL.Text := Format(
      ' SELECT A.CODENO,           ' +
      '        A.DESCRIPTION,      ' +
      '        A.SERVICETYPE,      ' +
      '        A.SIGN,             ' + 
      '        A.AMOUNT            ' +
      '   FROM %s.CD019 A          ' +
      '  WHERE A.CODENO = ''%s''   ' ,
      [DBController.LoginInfo.DbAccount, Citem.Value] );
    FReader.Open;
    {}
    FLockCItemChange := True;
    try
      if ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> EmptyStr ) and
         ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> ServiceType.Value ) then
        ServiceType.Value := FReader.FieldByName( 'SERVICETYPE' ).AsString;
    finally
      FLockCItemChange := False;
    end;
    {}
    edtAmount.Value := FReader.FieldByName( 'AMOUNT' ).AsInteger;
    if ( FReader.FieldByName( 'SIGN' ).AsString = '-' ) then
      edtAmount.Value := ( 0 - edtAmount.Value );
    {}
    edtDiscountAmt.Value := edtAmount.Value;
    {}
    FReader.Close;
    FaciDataSet.Filtered := False;
    FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
    FaciDataSet.Filtered := True;
    if not FaciDataSet.IsEmpty then
      FaciItem.Value := FaciDataSet.FieldByName( 'CODENO' ).AsString;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.ClearEditor;
begin
  Citem.Clear;
  ServiceType.Clear;
  FaciItem.Clear;
  edtDiscountAmt.Value := 0;
  edtAmount.Value := 0;
  chkPunish.Checked := False;
  cmbPenalType.ItemIndex := 0;
  edtInstCodeStr.Text := EmptyStr;
  cmbPenalType.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.DataToEditor;
begin
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  Citem.Value := FDataSet.FieldByName( 'CITEMCODE' ).AsString;
  Citem.ValueName := FDataSet.FieldByName( 'CITEMNAME' ).AsString;
  FaciItem.Value := FDataSet.FieldByName( 'FACIITEM' ).AsString;
  edtAmount.Value := FDataSet.FieldByName( 'AMOUNT' ).AsInteger;
  edtDiscountAmt.Value := FDataSet.FieldByName( 'DISCOUNTAMT' ).AsInteger;
  chkPunish.Checked := ( FDataSet.FieldByName( 'PUNISH' ).AsInteger = 1 );
  if ( FDataSet.FieldByName( 'PENALTYPE' ).AsString = '1' ) or
     ( FDataSet.FieldByName( 'PENALTYPE' ).AsString = '2' ) then
    cmbPenalType.ItemIndex := FDataSet.FieldByName( 'PENALTYPE' ).AsInteger;
  FQueryParam.Param1 := FDataSet.FieldByName( 'INSTCODESTR' ).AsString;
  edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B5.EditorToData: Boolean;
begin
  Result := True;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FDataSet.FieldByName( 'NEWBPCODE' ).AsString := FDataFrom;
    end;
    FDataSet.FieldByName( 'CITEMCODE' ).AsString := Citem.Value;
    FDataSet.FieldByName( 'CITEMNAME' ).AsString := Citem.ValueName;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'FACIITEM' ).AsString := FaciItem.Value;
    FDataSet.FieldByName( 'AMOUNT' ).AsInteger := Trunc( edtAmount.Value );
    FDataSet.FieldByName( 'DISCOUNTAMT' ).AsInteger := Trunc( edtDiscountAmt.Value );
    FDataSet.FieldByName( 'PUNISH' ).AsInteger := 0;
    if ( chkPunish.Checked ) then
      FDataSet.FieldByName( 'PUNISH' ).AsInteger := 1;
    FDataSet.FieldByName( 'PENALTYPE' ).AsString := EmptyStr;
    if ( cmbPenalType.ItemIndex > 0 ) then
      FDataSet.FieldByName( 'PENALTYPE' ).AsInteger := cmbPenalType.ItemIndex; 
    FDataSet.FieldByName( 'REFCITEMCODE' ).AsString := EmptyStr;
    FDataSet.FieldByName( 'REFCITEMNAME' ).AsString := EmptyStr;
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString :=
      TrimChar( FQueryParam.Param1, [''''] );
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '存檔失敗,原因:%s', [E.Message] ) );
      Result := False;
    end;
  end;
end;


{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end
  else  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.ButtonStateChange;
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

procedure TfrmCD078B5.EditorStateChange;
begin
  case FMode of
    dmInsert:
      begin
        ClearEditor;
        if ServiceType.CodeNo.CanFocus then
          ServiceType.CodeNo.SetFocus;
      end;
    dmUpdate:
      begin
        Citem.Enabled := True;
        ServiceType.Enabled := True;
        FaciItem.Enabled := True;
        edtAmount.Enabled := True;
        edtDiscountAmt.Enabled := True;
        chkPunish.Enabled := True;
        if ServiceType.CodeNo.CanFocus then ServiceType.CodeNo.SetFocus;
      end;
    dmBrowser:
      begin
        Citem.Enabled := False;
        ServiceType.Enabled := False;
        FaciItem.Enabled := False;
        edtAmount.Enabled := False;
        edtDiscountAmt.Enabled := False;
        chkPunish.Enabled := False;
        cmbPenalType.Enabled := False;
        if btnCancel.CanFocus then btnCancel.SetFocus;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B5.DataSetStateChange(const aMode: TDML): Boolean;
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
        ClearEditor;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;  
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '存檔有誤, 原因:%s。', [E.Message] ) );
              Result := False;
            end;
          end;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B5.DataValidate: Boolean;
var
  aErrMsg: String;
  aControl: TWinControl;
begin
  aErrMsg := EmptyStr;
  aControl := nil;
  try
    if ( Citem.Value = EmptyStr ) then
    begin
      aErrMsg := '請輸入【收費項目】。';
      aControl := Citem.CodeNo;
      Exit;
    end;
    if ( ServiceType.Value = EmptyStr ) then
    begin
      aErrMsg := '請輸入【服務類別】。';
      aControl := ServiceType.CodeNo;
      Exit;
    end;
    if ( edtAmount.EditingText = EmptyStr ) then
    begin
      aErrMsg := '請輸入【銷售牌價】。';
      aControl := edtAmount;
      Exit;
    end;
    if ( chkPunish.Checked ) then
    begin
      if ( cmbPenalType.ItemIndex = 0 ) then
      begin
        aErrMsg := '請輸入【違約時計價依據】。';
        aControl := cmbPenalType;
        Exit;
      end;
    end;
    if ( FQueryParam.Param1 = EmptyStr ) then
    begin
      aErrMsg := '請輸入【指定派工類別】。';
      aControl := btnInstCodeStr;
      Exit;
    end;
    {}
    if not DBController.VdInstCodeExitisInCD078B( FQueryParam.Param1,
      FInstCodeDataSet ) then
    begin
      aErrMsg := '【指定派工類別】中的裝機類別在【智慧型派工參數】中並不存在, 請重新設定。';
      aControl := btnInstCodeStr;
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

procedure TfrmCD078B5.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  Citem.Clear;
  FaciItem.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  ServiceType.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION  ' +
      '   FROM %s.CD046 A               ' +
      '  WHERE A.CODENO In (%s)         ' +
      '  ORDER BY A.CODENO              ' ,
      [DBController.LoginInfo.DbAccount, FServiceType] );
  ServiceType.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  if not FLockCItemChange then
  begin
    Citem.Clear;
    FaciItem.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FaciItemCodeNoPropertiesChange(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.CitemCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    Citem.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION   ' +
      '   FROM %s.CD019 A                ' +
      '  WHERE A.PERIODFLAG = 0          ' +
      '    AND A.STOPFLAG = 0            ',
      [DBController.LoginInfo.DbAccount] );
    if ( ServiceType.value <> EmptyStr )then
    begin
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.SERVICETYPE=''' + ServiceType.Value+''' ';
    end;
    CItem.SQL.Text := CItem.SQL.Text +
      '  ORDER BY A.CODENO             ';
    Citem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.CitemCodeNoExit(Sender: TObject);
begin
  Citem.CodeNoExit(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.CitemCodeNoPropertiesChange(Sender: TObject);
begin
  LoadRelationItem;
  Citem.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.CitemCodeNamePropertiesChange(Sender: TObject);
begin
  Citem.CodeNamePropertiesChange( Sender );
  LoadRelationItem;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.chkPunishPropertiesChange(Sender: TObject);
begin
  cmbPenalType.Enabled := chkPunish.Checked;
  if chkPunish.Checked then
  begin
    cmbPenalType.Enabled := False;
    cmbPenalType.ItemIndex := 1;
  end else
  begin
    cmbPenalType.ItemIndex := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,裝機代碼,INSTNAME,裝機類別名稱';
      frmMultiSelect.ResultFields := 'INSTNAME';
      FInstCodeDataSet.Filtered := False;
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filtered := True;
      try
        frmMultiSelect.DataSet := FInstCodeDataSet;
        if frmMultiSelect.ShowModal = mrOk then
        begin
          if ( FMode in [dmInsert, dmUpdate] ) then
          begin
            FQueryParam.Param1 := frmMultiSelect.SelectedValue;
            edtInstCodeStr.Text := frmMultiSelect.SelectedDisplay;
          end;
        end;
      finally
        FInstCodeDataSet.Filtered := False;
        FInstCodeDataSet.Filter := EmptyStr;
      end;
    finally
      frmMultiSelect.Free;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B5.cxButton1Click(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    FQueryParam.Param1 := EmptyStr;
    edtInstCodeStr.Text := EmptyStr;
  end;  
end;

{ ---------------------------------------------------------------------------- }

end.
