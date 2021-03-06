unit frmCD078IU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Menus, cxLookAndFeelPainters, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxButtons, StdCtrls,
  cbDataLookup, ExtCtrls, frmCD078B1U, DBClient, frmCD078B0U, cxCheckBox,
  cxCurrencyEdit, ActnList;

type
  TfrmCD078I = class(TForm)
    Panel1: TPanel;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    ServiceType: TDataLookup;
    CItem: TDataLookup;
    FaciItem: TDataLookup;
    btnInstCodeStr: TButton;
    btnDelInstCode: TcxButton;
    edtInstCodeStr: TcxTextEdit;
    chkSelectData: TcxCheckBox;
    edtDiscountAmt: TcxCurrencyEdit;
    Label4: TLabel;
    Panel3: TPanel;
    btnUpdate: TButton;
    btnSave: TButton;
    btnCancel: TButton;
    edtDML: TcxTextEdit;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    actUpdate: TAction;
    lbl1: TLabel;
    edtAmount: TcxCurrencyEdit;
    lbl2: TLabel;
    edtCalcDisountAmunt: TcxCurrencyEdit;
    lbl3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure CItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure btnDelInstCodeClick(Sender: TObject);
    procedure FaciItemCodeNamePropertiesChange(Sender: TObject);
    procedure FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure FaciItemCodeNoPropertiesChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CItemCodeNamePropertiesChange(Sender: TObject);
    procedure CItemCodeNoPropertiesChange(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure EnabledCTL;
    procedure chkSelectDataPropertiesChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure edtDiscountAmtPropertiesEditValueChanged(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataSet: TClientDataSet;
    FAlreadySetCItemValue: String;
    FInstCodeDataSet : TClientDataSet;
    FaciDataSet: TClientDataSet;
    FQueryParam: TQueryParam;
    procedure ClearEditor;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function VdMustInput: Boolean;
    procedure DMLModeChange(aMode: TDML);
    procedure ShowSO562Amt;
    procedure ButtonStateChange;
    procedure EditorStateChange;
    procedure DataToEditor;
    function EditorToData: Boolean;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue: String; aDataSet:TClientDataSet );reintroduce;
    property AlreadySetCItemValue: String read FAlreadySetCItemValue write FAlreadySetCItemValue;
    property FaciItemDataSet: TClientDataSet read FaciDataSet write FaciDataSet;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
  end;

var
  frmCD078I: TfrmCD078I;

implementation
  uses cbUtilis, frmMultiSelectU, cbDBController;
{$R *.dfm}

constructor TfrmCD078I.Create(aMode: TDML; aKeyValue: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FDataSet := aDataSet;
end;

procedure TfrmCD078I.FormCreate(Sender: TObject);
begin
  CItem.Initializa;
  ServiceType.Initializa;
  FaciItem.Initializa;

end;

procedure TfrmCD078I.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  CItem.Clear;
  FaciItem.Clear;
end;

procedure TfrmCD078I.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  { ?A?????O }
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ' +
    '  WHERE CODENO=''%s''          ' ,
    [DBController.LoginInfo.DbAccount,'D'] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;

end;

procedure TfrmCD078I.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  CItem.Clear;
  FaciItem.Clear;
end;

procedure TfrmCD078I.CItemCodeNamePropertiesInitPopup(Sender: TObject);
var
  aCItems: String;
begin
  Screen.Cursor := crSQLWait;
  try
    CItem.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION      ' +
      '  FROM %s.CD019 A                    ' +
      ' WHERE A.SIGN = ''-''                ' +
      '   AND Nvl(A.STOPFLAG,0) = 0         ',
      [DBController.LoginInfo.DbAccount] );
    {}
    if ( ServiceType.value <> EmptyStr )then
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.SERVICETYPE=''' + ServiceType.Value+''' ';
    {}
    if ( FMode in [dmInsert] ) and ( FAlreadySetCItemValue <> EmptyStr ) then
    begin

      aCItems := QuotedValue( FAlreadySetCItemValue );
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.CODENO NOT IN ( ' + aCItems + ' ) ';
    end;
    {}
    CItem.SQL.Text := CItem.SQL.Text +
      '  ORDER BY A.CODENO  ';
    CItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078I.ClearEditor;
begin
  ServiceType.Clear;
  CItem.Clear;
  FaciItem.Clear;
  chkSelectData.Checked := False;
  FQueryParam.Param1 := EmptyStr;
  edtInstCodeStr.Text := EmptyStr;
  edtDiscountAmt.Text := EmptyStr;
end;

function TfrmCD078I.DataSetStateChange(const aMode: TDML): Boolean;
begin

  Result := True;
  case aMode of
    dmInsert:
      begin
//        FilterCD078G1DataWhenNew;
        ClearEditor;
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        EnabledCTL;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser,dmCannotModify:
      begin
        ClearEditor;
        DataToEditor;
      end;  
    dmSave:
      begin
        Result := VdMustInput;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
//            FAlreadySetCItemValue := frmCD078B1.GetCD078GCItemValue;
            Self.Close;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '?s?????~, ???]:%s?C', [E.Message] ) );
              Result := False;
            end;  
          end;
        end;
      end;
  end;

end;

function TfrmCD078I.VdMustInput: Boolean;
begin

  Result := False;
  edtDiscountAmt.PostEditValue;
  if not chkSelectData.Checked then
  begin
    Result := False;
  end;
  {
  if ( CItem.Value = EmptyStr ) and ( chkSelectData.Checked ) then
  begin
    WarningMsg( '?????J?i???O?????j?C' );
    if CItem.CodeNo.CanFocusEx then CItem.CodeNo.SetFocus;
    Exit;
  end;
  }
  {}
  if ( ServiceType.Value = EmptyStr ) and ( chkSelectData.Checked ) then
  begin
    WarningMsg( '?????J?i?A?????O?j?C' );
    if ServiceType.CodeNo.CanFocusEx then ServiceType.CodeNo.SetFocus;
    Exit;
  end;
  if ( FQueryParam.Param1 = EmptyStr ) and ( chkSelectData.Checked ) then
  begin
    WarningMsg( '?????w???u???O?I');
    if btnInstCodeStr.CanFocus then btnInstCodeStr.Click;
    Exit;
  end;


  if chkSelectData.Checked then
  begin
    if ( edtDiscountAmt.Value = 0 )  then
    begin
      WarningMsg( '?????J?u?f???B?I');
      if edtDiscountAmt.CanFocus then edtDiscountAmt.SetFocus;
      Exit;
    end else
    begin
      if edtDiscountAmt.Value < 0 then
      begin
        //edtDiscountAmt.Value := edtDiscountAmt.Value * -1;
        WarningMsg( '?u?f???B???o???t?I');
        if edtDiscountAmt.CanFocus then edtDiscountAmt.SetFocus;
        Exit;
      end else
      begin
        if edtDiscountAmt.Value > FDataSet.FieldByName( 'AMOUNT' ).AsFloat then
        begin
          WarningMsg( Format('?u?f???B???o?j???P?? [ %s ] ?I',
                          [FDataSet.FieldByName( 'AMOUNT' ).AsString]));
          if edtDiscountAmt.CanFocus then edtDiscountAmt.SetFocus;
          Exit;
        end;
      end;
    end;
  end;

  Result := True;
end;

procedure TfrmCD078I.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,?????N?X,INSTNAME,???????O?W??';
      frmMultiSelect.ResultFields := 'INSTNAME';
      FInstCodeDataSet.Filtered := False;
//      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ ServiceType.Value ] );
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

procedure TfrmCD078I.btnDelInstCodeClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    FQueryParam.Param1 := EmptyStr;
    edtInstCodeStr.Text := EmptyStr;
  end;
end;

procedure TfrmCD078I.FaciItemCodeNamePropertiesChange(Sender: TObject);
begin
  FaciItem.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078I.FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNamePropertiesInitPopup( Sender );
end;

procedure TfrmCD078I.FaciItemCodeNoPropertiesChange(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078I.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FaciItem.DataSet := FaciDataSet;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078I.CItemCodeNamePropertiesChange(Sender: TObject);
begin
  CItem.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078I.CItemCodeNoPropertiesChange(Sender: TObject);
begin
  CItem.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078I.DMLModeChange(aMode: TDML);
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
  ShowSO562Amt;
  edtDML.Text := TDMLString[FMode];
  {
  if not FDataSet.IsEmpty then
  begin
    frmCD078B1.btnUpdate.Enabled := True;
    frmCD078B1.actUpdate.Enabled := True;
    frmCD078B1.btnDelete.Enabled := True;
    frmCD078B1.actDelete.Enabled := True;
  end;
  }
end;

procedure TfrmCD078I.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
    actCancel.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '????(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

procedure TfrmCD078I.EditorStateChange;
var aIndex : Integer;
begin
  for aIndex := 0 to ControlCount - 1 do
  begin
    if ( Controls[aIndex].Name <> 'btnSave' ) and
       ( Controls[aIndex].Name <> 'Panel3' )  then
      Controls[aIndex].Enabled := not ( FMode in [dmBrowser,dmCannotModify] );
  end;
  {#6385 ?nShow?P???P???????B,???????\???J By Kin 2012/11/29}
  edtAmount.Enabled := False;
  edtCalcDisountAmunt.Enabled := False;
  btnUpdate.Enabled := ( FMode in [dmBrowser] );
  btnCancel.Enabled := ( FMode in [dmInsert,dmUpdate] );
  case FMode of
    dmInsert:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        CItem.Enabled := True;
        ServiceType.Enabled := True;
      end;
    dmUpdate:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        //CItem.Enabled := True;
        ServiceType.Enabled := False;

      end;
    dmBrowser,dmCannotModify:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      end;
  end;
end;

procedure TfrmCD078I.DataToEditor;
begin
  chkSelectData.Checked := FDataSet.FieldByName( 'CHOICE' ).AsString = '1';
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  CItem.Value := Trim( FDataSet.FieldByName( 'CITEMCODE2' ).AsString );
  CItem.ValueName := Trim( FDataSet.FieldByName( 'CITEMNAME2' ).AsString );
  if ( CItem.Value = '-1' ) or ( CItem.Value = '0' ) then
  begin
    CItem.Value := EmptyStr;
    CItem.ValueName := EmptyStr;

  end;

  FaciItem.Value := Trim( FDataSet.FieldByName( 'FACIITEM' ).AsString );
  if ( FaciItem.Value = '-1' ) or ( FaciItem.Value = '0' ) then
  begin
    FaciItem.Value := EmptyStr;
  end;

  FQueryParam.Param1 :=  Trim( FDataSet.FieldByName( 'INSTCODESTR' ).AsString );
  edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
  edtDiscountAmt.Text := FDataSet.FieldByName( 'DISCOUNTAMT' ).AsString;

//  cdsCD078G1AfterScroll( cdsCD078G1 );
end;

procedure TfrmCD078I.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

procedure TfrmCD078I.btnUpdateClick(Sender: TObject);
begin
  FMode := dmUpdate;
  DMLModeChange(dmUpdate);  
end;

procedure TfrmCD078I.btnCancelClick(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

function TfrmCD078I.EditorToData: Boolean;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    end;
    FDataSet.FieldByName( 'CHOICE' ).AsString := '0';
    if chkSelectData.Checked then FDataSet.FieldByName( 'CHOICE' ).AsString := '1';

    FDataSet.FieldByName( 'CITEMCODE2' ).AsString := CItem.Value;
    FDataSet.FieldByName( 'CITEMNAME2' ).AsString := CItem.ValueName;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'FACIITEM' ).AsString := FaciItem.Value;
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString := TrimChar( FQueryParam.Param1, [''''] );
    FDataSet.FieldByName( 'DISCOUNTAMT' ).AsFloat := edtDiscountAmt.Value;
   except
    on E: Exception do
    begin
      ErrorMsg( Format( '?s?????~, ???]:%s', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;

procedure TfrmCD078I.EnabledCTL;
var aEnabled:Boolean;

begin
  {#6385 ?????????????i?s?? By Kin 2012/11/29}
  aEnabled := chkSelectData.Checked;
  aEnabled := True;
  ServiceType.Enabled := aEnabled;
  CItem.Enabled := aEnabled;
  FaciItem.Enabled := aEnabled;
  edtDiscountAmt.Enabled := aEnabled;
  btnInstCodeStr.Enabled := aEnabled;
  btnDelInstCode.Enabled := aEnabled;

end;

procedure TfrmCD078I.chkSelectDataPropertiesChange(Sender: TObject);
begin
  EnabledCTL;
end;

procedure TfrmCD078I.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

procedure TfrmCD078I.FormDestroy(Sender: TObject);
begin
  ServiceType.Finaliza;
  CItem.Finaliza;
  FaciItem.Finaliza;
end;

procedure TfrmCD078I.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

procedure TfrmCD078I.edtDiscountAmtPropertiesEditValueChanged(
  Sender: TObject);
begin
  edtCalcDisountAmunt.Value := edtAmount.Value - edtDiscountAmt.Value;
  {
  if edtDiscountAmt.Text <> EmptyStr then
  begin
    if edtDiscountAmt.Value > 0 then
    begin
      edtDiscountAmt.Value := 0 - edtDiscountAmt.Value;
    end;
  end;
  }
end;

procedure TfrmCD078I.ShowSO562Amt;
begin
  edtAmount.Text := FDataSet.FieldByName( 'Amount' ).AsString;
  if edtAmount.Text = EmptyStr then
    edtAmount.Value := 0;
end;

end.
