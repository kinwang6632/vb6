unit frmInvA07_5U;

interface

uses
 Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, dtmMainU, ADODB, DBClient, DB,
  cxControls, cxContainer,  cxEdit, cxLabel, Menus, cxLookAndFeelPainters,
  cxButtons, cxGraphics, cxDropDownEdit, cxRadioGroup, cxGroupBox, cxTextEdit,
  cxMaskEdit, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TfrmInvA07_5 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    lblDisType: TLabel;
    Label4: TLabel;
    lblMsg: TLabel;
    btnExit: TcxButton;
    btnOK: TcxButton;
    txtDisNoSt: TcxMaskEdit;
    cxGroupBox1: TcxGroupBox;
    Label2: TLabel;
    rdoDrop: TcxRadioButton;
    rdoRecover: TcxRadioButton;
    cmbReason: TcxComboBox;
    txtDisNoEd: TcxMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    function isDataOK : Boolean;
    procedure txtDisNoStExit(Sender: TObject);
    procedure btnOKClick(Sender: TObject);


  private
    { Private declarations }
    FDisNo: String;
    FList: TStringList;

    procedure InitialComboBox;
    procedure UnPrepareMISDataSet;
    procedure SetDisNo(const Value: String);
    function IsDisObsolete(aDisNo: String): Boolean;
    procedure GetRelationDisList(const aDisNoSt, aDisNoEd: String);
  public
    { Public declarations }
    property DisNo: String read FDisNo write SetDisNo;
  end;

var
  frmInvA07_5: TfrmInvA07_5;
  const DisnoType: String = 'ALLOWANCENO';

implementation
uses cbUtilis, frmMainU, dtmMainJU, dtmSOU;
{$R *.dfm}

{ TForm1 }

procedure TfrmInvA07_5.InitialComboBox;
var
  aParam: TComboBoxCreateParam;
begin
  aParam.Sql := ' SELECT ITEMID, DESCRIPTION FROM INV006 ORDER BY ITEMID ';
  aParam.KeyField := 'ITEMID';
  aParam.DescField := 'DESCRIPTION';
  aParam.AddAllText := False;
  CreateCxComboBoxItem( cmbReason, aParam );

  if ( cmbReason.Properties.Items.Count > 0 ) then cmbReason.ItemIndex := 0;
end;

procedure TfrmInvA07_5.FormShow(Sender: TObject);
begin
  InitialComboBox;
  lblMsg.Caption := EmptyStr;
  lblDisType.Caption := '?o?????X : ';
  if UpperCase( DisnoType ) = UpperCase( 'ALLOWANCENO' ) then
  begin
    lblDisType.Caption := '?????????X : ';
  end;

  if ( FDisNo = EmptyStr ) then
  begin
    if ( txtDisNoSt.CanFocusEx ) then txtDisNoSt.SetFocus;
    rdoDrop.Checked := True;
  end;
end;

procedure TfrmInvA07_5.UnPrepareMISDataSet;
begin
  {
  if not VarIsNull( SODataSet.Data ) then
    SODataSet.EmptyDataSet;
  SODataSet.Data := Null;
  }
end;

procedure TfrmInvA07_5.SetDisNo(const Value: String);
begin
  if ( FDisNo <> Value ) then FDisNo := Value;
  txtDisNoSt.Text := FDisNo;
  txtDisNoSt.Enabled := ( txtDisNoSt.Text = EmptyStr );
  {}
  txtDisNoEd.Text := txtDisNoSt.Text;
  txtDisNoEd.Enabled := txtDisNoSt.Enabled;
  {}
  if ( txtDisNoEd.Enabled ) then Exit;
  {}
  if IsDisObsolete( FDisNo ) then
  begin
    rdoDrop.Enabled := False;
    cmbReason.Enabled := False;
    rdoRecover.Enabled := True;
    rdoRecover.Checked := True;
  end else
  begin
    rdoDrop.Enabled := True;
    cmbReason.Enabled := True;
    rdoRecover.Enabled := False;
    rdoDrop.Checked := True;
  end;
end;

function TfrmInvA07_5.IsDisObsolete(aDisNo: String): Boolean;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT ISOBSOLETE            ' +
    '   FROM INV014                ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID = ''%s''       ' +
    '    AND %s = ''%s''  ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,DisnoType, aDisNo] );
  dtmMain.adoComm.Open;
  Result := ( dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' );
  dtmMain.adoComm.Close;
end;

procedure TfrmInvA07_5.FormCreate(Sender: TObject);
begin
    FList := TStringList.Create;
end;

procedure TfrmInvA07_5.FormDestroy(Sender: TObject);
begin
  FList.Free;

end;

procedure TfrmInvA07_5.btnExitClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
  Close;
end;

function TfrmInvA07_5.isDataOK: Boolean;
var
  aMsg: String;
  aFocusObj: TWinControl;
begin
  Result := False;
  try
    if Trim( txtDisNoSt.Text ) = EmptyStr then
    begin
      aMsg := '?????J?????????X?C';
      aFocusObj := txtDisNoSt;
      Exit;
    end;
    if ( rdoDrop.Checked ) and ( cmbReason.ItemIndex < 0 ) then
    begin
      aMsg := '?????????o???]?C';
      aFocusObj := cmbReason;
      Exit;
    end;
  finally
    Result := ( aMsg = EmptyStr );
    if not Result then
    begin
      WarningMsg( aMsg );
      Self.ActiveControl := aFocusObj;
    end;
  end;
end;

procedure TfrmInvA07_5.txtDisNoStExit(Sender: TObject);
begin
  txtDisNoEd.Text := txtDisNoSt.Text;
end;

procedure TfrmInvA07_5.btnOKClick(Sender: TObject);
  var
  aDisNo: String;
  aParam: TComboBoxItemParam;
  aCommonParam: TCommonParam;
  aExecute: Boolean;
  aIndex: Integer;
begin

  if not IsDataOK then Exit;
  Screen.Cursor := crSQLWait;
  try
    ZeroMemory( @aParam, SizeOf( aParam ) );
    ZeroMemory( @aCommonParam, SizeOf( aCommonParam ) );
    GetRelationDisList( txtDisNoSt.Text, Nvl( txtDisNoEd.Text, txtDisNoSt.Text ) );
    if ( dtmMain.GetLinkToMIS ) then
    begin
      for aIndex := 0 to FList.Count -1 do
      begin
        aCommonParam.InStr1 := FList.Strings[aIndex];
      end;
    end;
    for aIndex := 0 to FList.Count - 1 do
    begin
      aCommonParam.InStr1 := FList.Strings[aIndex];
      if ( rdoDrop.Checked ) then
        aCommonParam.InStr2 := 'Y'
      else
        aCommonParam.InStr2 := EmptyStr;
      aExecute := dtmMainJ.CanDropOrRecoverDis( aCommonParam );
      if ( aExecute ) then
      begin
        if ( rdoDrop.Checked ) then
        begin
          GetCxComboBoxItemValue( cmbReason, aParam );
          aCommonParam.InStr2 := aParam.KeyValue;
          aCommonParam.InStr3 := aParam.DesValue;
        end else
        begin
          aCommonParam.InStr2 := EmptyStr;
          aCommonParam.InStr3 := EmptyStr;
        end;
        { ?o???t?????@?o }
        { aParam.InStr1 = ?o?????X
          aParam.InStr2 = ???o?N?X
          aParam.InStr3 = ???o???]
          aParam.OutStr1 = HowToCreate,
          aParam.OutStr2 = ?o?????? }
        aExecute := dtmMainJ.DropOrRecoverDis( aCommonParam );

      end;
    end;
    if ( aExecute ) then
    begin
      if ( rdoDrop.Checked ) then
        lblMsg.Caption := '???i???????@?o?????C'
      else
        lblMsg.Caption := '???i???????^?_?????C'
    end else
    begin
      lblMsg.Caption := aCommonParam.Msg;
    end;
    if ( txtDisNoSt.CanFocusEx ) then txtDisNoSt.SetFocus;
  finally
    Screen.Cursor := crDefault;
  end;
  if ( FDisNo <> EmptyStr ) and ( aExecute ) then
    Self.ModalResult := mrOk
  else if ( txtDisNoSt.CanFocusEx ) then
    txtDisNoSt.SetFocus;

end;

procedure TfrmInvA07_5.GetRelationDisList(const aDisNoSt,
  aDisNoEd: String);
  var
  aDataSet: TADOQuery;
begin
  FList.Clear;
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    ' SELECT ALLOWANCENO FROM INV014 A                       ' +
    '  WHERE A.IDENTIFYID1 = ''%s''                          ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ' +
    '    AND A.%s >= ''%s''                                  ' +
    '    AND A.%s <= ''%s''                         ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,DisnoType, aDisNoSt,DisnoType, aDisNoEd] );
  aDataSet.Open;
  aDataSet.First;
  while not aDataSet.Eof do
  begin
    FList.Add( aDataSet.FieldByName( 'ALLOWANCENO' ).AsString );
    aDataSet.Next;
  end;
  aDataSet.Close;
end;

end.
