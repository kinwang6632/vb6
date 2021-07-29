unit frmInvB02U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, dtmMainU, ADODB, DBClient, DB,
  cxControls, cxContainer,  cxEdit, cxLabel, Menus, cxLookAndFeelPainters,
  cxButtons, cxGraphics, cxDropDownEdit, cxRadioGroup, cxGroupBox, cxTextEdit,
  cxMaskEdit, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TfrmInvB02 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Label3: TLabel;
    lblMsg: TLabel;
    btnExit: TcxButton;
    btnOK: TcxButton;
    txtInvIdSt: TcxMaskEdit;
    cxGroupBox1: TcxGroupBox;
    rdoDrop: TcxRadioButton;
    rdoRecover: TcxRadioButton;
    Label2: TLabel;
    cmbReason: TcxComboBox;
    SODataSet: TClientDataSet;
    Label4: TLabel;
    txtInvIdEd: TcxMaskEdit;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure txtInvIdPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure txtInvIdStExit(Sender: TObject);
  private
    { Private declarations }
    FInvId: String;
    FList: TStringList;
    procedure InitialComboBox;
    function isDataOK : Boolean;
    function IsInvObsolete(aInvId: String): Boolean;
    procedure SetInvId(const Value: String);
    procedure PrepareMISDataSet;
    procedure UnPrepareMISDataSet;
    procedure CatchBillData(var aParam: TCommonParam);
    function MISDropOrRecover(var aParam: TCommonParam): Boolean;
    procedure GetRelationInvList(const aInvIdSt, aInvIdEd: String);
  public
    { Public declarations }
    property InvId: String read FInvId write SetInvId; 
  end;

var
  frmInvB02: TfrmInvB02;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmSOU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.FormCreate(Sender: TObject);
begin
  FList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.FormShow(Sender: TObject);
var
  sL_QueryCompetence, sL_InsertCompetence, sL_DeleteCompetence,
  sL_UpdateCompetence, sL_ExecuteCompetence,sL_VerifyCompetence: String;
begin
  InitialComboBox;
  Self.Caption := frmMain.GetFormTitleString( 'B02', '發票作廢與回復' );
  //設定權限
  dtmMain.getCompetence('B02', sL_QueryCompetence, sL_InsertCompetence,
    sL_DeleteCompetence, sL_UpdateCompetence, sL_ExecuteCompetence,sL_VerifyCompetence );
  if sL_ExecuteCompetence = 'N' then
    btnOK.Enabled := False;
  lblMsg.Caption := EmptyStr;
  if ( FInvId = EmptyStr ) then
  begin
    if ( txtInvIdSt.CanFocusEx ) then txtInvIdSt.SetFocus;
    rdoDrop.Checked := True;
  end;
  UnPrepareMISDataSet;
  PrepareMISDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.FormDestroy(Sender: TObject);
begin
  FList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  UnPrepareMISDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.InitialComboBox;
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

{ ---------------------------------------------------------------------------- }

function TfrmInvB02.isDataOK: Boolean;
var
  aMsg: String;
  aFocusObj: TWinControl;
begin
  Result := False;
  try
    if Trim( txtInvIdSt.Text ) = EmptyStr then
    begin
      aMsg := '請輸入發票號碼。';
      aFocusObj := txtInvIdSt;
      Exit;
    end;
    if ( rdoDrop.Checked ) and ( cmbReason.ItemIndex < 0 ) then
    begin
      aMsg := '請選擇做廢原因。';
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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.btnOKClick(Sender: TObject);
var
  aInvId: String;
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
    { 發票號碼 }
    GetRelationInvList( txtInvIdSt.Text, Nvl( txtInvIdEd.Text, txtInvIdSt.Text ) );
    { 多客戶的發票, 自動產生折讓單的資料, 發票系統裏沒有記錄對應關係,
      做廢後回復會沒辦法回填, 在這裏先記錄下來 }
    if ( dtmMain.GetLinkToMIS ) then
    begin
      for aIndex := 0 to FList.Count -1 do
      begin
        aCommonParam.InStr1 := FList.Strings[aIndex];
        CatchBillData( aCommonParam );
      end;
    end;
    for aIndex := 0 to FList.Count - 1 do
    begin
      aCommonParam.InStr1 := FList.Strings[aIndex];
      if ( rdoDrop.Checked ) then
        aCommonParam.InStr2 := 'Y'
      else
        aCommonParam.InStr2 := EmptyStr;
      aExecute := dtmMainJ.CanDropOrRecoverInv( aCommonParam );
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
        { 發票系統的作廢 }
        { aParam.InStr1 = 發票號碼
          aParam.InStr2 = 做廢代碼
          aParam.InStr3 = 做廢原因
          aParam.OutStr1 = HowToCreate,
          aParam.OutStr2 = 發票日期 }
        aExecute := dtmMainJ.DropOrRecoverInv( aCommonParam );
        { 客服系統資料更新 }
        if ( aExecute ) and ( dtmMain.GetLinkToMIS ) then
        begin
          { INVDATE 發票日期 }
          aCommonParam.InStr2 := aCommonParam.OutStr2;
          { HowToCreate }
          aCommonParam.InStr5 := aCommonParam.OutStr1;
          MISDropOrRecover( aCommonParam );
        end;
      end;
    end;
    if ( aExecute ) then
    begin
      if ( rdoDrop.Checked ) then
        lblMsg.Caption := '此張發票作廢完畢。'
      else
        lblMsg.Caption := '此張發票回復完成。'
    end else
    begin
      lblMsg.Caption := aCommonParam.Msg;
    end;
    if ( txtInvIdSt.CanFocusEx ) then txtInvIdSt.SetFocus;
  finally
    Screen.Cursor := crDefault;
  end;
  if ( FInvId <> EmptyStr ) and ( aExecute ) then
    Self.ModalResult := mrOk
  else if ( txtInvIdSt.CanFocusEx ) then
    txtInvIdSt.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.SetInvId(const Value: String);
begin
  if ( FInvId <> Value ) then FInvId := Value;
  txtInvIdSt.Text := FInvId;
  txtInvIdSt.Enabled := ( txtInvIdSt.Text = EmptyStr );
  {}
  txtInvIdEd.Text := txtInvIdSt.Text;
  txtInvIdEd.Enabled := txtInvIdSt.Enabled;
  {}
  if ( txtInvIdSt.Enabled ) then Exit;
  {}
  if IsInvObsolete( FInvId ) then
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

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.btnExitClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
  Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB02.IsInvObsolete(aInvId: String): Boolean;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT ISOBSOLETE            ' +
    '   FROM INV007                ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID = ''%s''       ' +
    '    AND INVID = ''%s''        ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvId] );
  dtmMain.adoComm.Open;
  Result := ( dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' );
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.txtInvIdPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    ErrorText := EmptyStr;
    WarningMsg( '發票格式有誤, 請重新輸入。' );
  end else
    ErrorText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.PrepareMISDataSet;
begin
  if ( SODataSet.FieldDefs.Count <= 0 ) then
  begin
    SODataSet.FieldDefs.Add( 'INVID', ftString, 10 );
    SODataSet.FieldDefs.Add( 'CUSTID', ftString, 10 );
    SODataSet.FieldDefs.Add( 'SEQ', ftInteger );
    SODataSet.FieldDefs.Add( 'BILLID', ftString, 15 );
    SODataSet.FieldDefs.Add( 'BILLIDITEMNO', ftInteger );
    SODataSet.FieldDefs.Add( 'INVPURPOSECODE', ftString, 3 );
    SODataSet.FieldDefs.Add( 'INVPURPOSENAME', ftString, 20 );
  end;
  SODataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.UnPrepareMISDataSet;
begin
  if not VarIsNull( SODataSet.Data ) then
    SODataSet.EmptyDataSet;
  SODataSet.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.CatchBillData(var aParam: TCommonParam);
var
  aDataSet: TADOQuery;
  aIndex: Integer;

  { --------------------------------------------------- }

  function GetSearchSql: String;
  var
    aTable: String;
  begin
    aTable := ' INV008A ';
    if aIndex = 2 then aTable := ' INV008 ';
    Result := Format(
      ' SELECT A.CUSTID, B.SEQ, B.BILLID, B.BILLIDITEMNO,       ' +
      '        A.INVUSEID, A.INVUSEDESC                         ' +
      '   FROM INV007 A, %s B                                   ' +
      '  WHERE A.INVID = B.INVID                                ' +
      '    AND A.IDENTIFYID1 = ''%s''                           ' +
      '    AND A.IDENTIFYID2 = ''%s''                           ' +
      '    AND A.COMPID = ''%s''                                ' +
      '    AND A.INVID = ''%s''                                 ',
      [aTable, IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  end;

  { ---------------------------------------------------------------------------- }

begin
  { 先取有合併的資料 INV008A, 再取沒有合併的明細 INV008 }
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  for aIndex := 1 to 2 do
  begin
    aDataSet.SQL.Text := GetSearchSql;
    aDataSet.Open;
    aDataSet.First;
    while not aDataSet.Eof do
    begin
      if ( aDataSet.FieldByName( 'BILLID' ).AsString <> EmptyStr ) and
         ( aDataSet.FieldByName( 'BILLIDITEMNO' ).AsString <> EmptyStr ) then
      begin
        if not SODataSet.Locate( 'INVID;BILLID;BILLIDITEMNO',
          VarArrayOf( [aParam.InStr1, aDataSet.FieldByName( 'BILLID' ).AsString,
          aDataSet.FieldByName( 'BILLIDITEMNO' ).AsString] ), [] ) then
        begin
          SODataSet.Last;
          SODataSet.Append;
          SODataSet.FieldByName( 'INVID' ).AsString := aParam.InStr1;
          SODataSet.FieldByName( 'CUSTID' ).AsString := aDataSet.FieldByName( 'CUSTID' ).AsString;
          SODataSet.FieldByName( 'SEQ' ).AsString := aDataSet.FieldByName( 'SEQ' ).AsString;
          SODataSet.FieldByName( 'BILLID' ).AsString := aDataSet.FieldByName( 'BILLID' ).AsString;
          SODataSet.FieldByName( 'BILLIDITEMNO' ).AsString := aDataSet.FieldByName( 'BILLIDITEMNO' ).AsString;
          SODataSet.FieldByName( 'INVPURPOSECODE' ).AsString := aDataSet.FieldByName( 'INVUSEID' ).AsString;
          SODataSet.FieldByName( 'INVPURPOSENAME' ).AsString := aDataSet.FieldByName( 'INVUSEDESC' ).AsString;
          SODataSet.Post;
        end;
      end;
      aDataSet.Next;
    end;
    aDataSet.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB02.MISDropOrRecover(var aParam: TCommonParam): Boolean;

   { ----------------------------------------- }

   function ConvertPreInvoce(const aValue: String): String;
   begin
     //若發票為mis拋檔, 預開(HowToCreate=1), 則將PREINVOICE update成1
     if ( aValue = '1' ) then
       Result := '1'
     //若發票為mis拋檔, 後開(HowToCreate=2), 則將PREINVOICE update成0
     else if ( aValue = '2' ) then
       Result := '0'
     //若發票為invoice create, 預開(HowToCreate=3), 則將PREINVOICE update成2
     else if ( aValue = '3' ) then
       Result := '2'
     //若發票為invoice create, 後開(HowToCreate=4), 則將PREINVOICE update成0
     else if ( aValue = '4' ) then
       Result := '0';
   end;

   { ----------------------------------------- }
var
  aCanExecute: Boolean;   
begin
  SODataSet.Filtered := False;
  SODataSet.Filter := Format( 'INVID=''%s''', [aParam.InStr1] );
  SODataSet.Filtered := True;
  SODataSet.First;
  while not SODataSet.Eof do
  begin
    aParam.InStr3 := SODataSet.FieldByName( 'BILLID' ).AsString;
    aParam.InStr4 := SODataSet.FieldByName( 'BILLIDITEMNO' ).AsString;
    { 該筆明細是否與客服系統介接 }
    { 參數: 1.發票號碼(INVID), 2:序號(SEQ)  }
    aCanExecute := dtmMainJ.InvDeatilLinkToMIS( aParam.InStr1,
      SODataSet.FieldByName( 'SEQ' ).AsString );
    if ( aCanExecute ) then
    begin
      { HowToCreate }
      if ( aParam.OutStr1 = '1' ) or ( aParam.OutStr1 = '2' ) or
         ( aParam.OutStr1 = '3' ) then
      begin
        if ( rdoDrop.Checked ) then
        begin
          dtmSO.DropInvData( aParam );
        end else
        begin
          aParam.InStr5 := ConvertPreInvoce( aParam.OutStr1 );
          aParam.InStr6 := SODataSet.FieldByName( 'CUSTID' ).AsString;
          aParam.InStr7 := SODataSet.FieldByName( 'INVPURPOSECODE' ).AsString;
          aParam.InStr8 := SODataSet.FieldByName( 'INVPURPOSENAME' ).AsString;
          dtmSO.RecoverInvData( aParam );
        end;
      end;
    end;  
    SODataSet.Next;
  end;
  SODataSet.Filtered := False;
  SODataSet.Filter := EmptyStr;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.GetRelationInvList(const aInvIdSt, aInvIdEd: String);
var
  aDataSet: TADOQuery;
begin
  FList.Clear;
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    ' SELECT INVID FROM INV007 A,                            ' +
    '   ( SELECT A.MAININVID FROM INV007 A                   ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                      ' +
    '        AND A.IDENTIFYID2 = ''%s''                      ' +
    '        AND A.COMPID = ''%s''                           ' +
    '        AND A.INVID BETWEEN ''%s'' AND ''%s'' ) B       ' +
    '  WHERE A.MAININVID = B.MAININVID                       ' +
    '    AND A.IDENTIFYID1 = ''%s''                          ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvIdSt, aInvIdEd,
     IDENTIFYID1, IDENTIFYID2,  dtmMain.getCompID] );
  aDataSet.Open;
  aDataSet.First;
  while not aDataSet.Eof do
  begin
    FList.Add( aDataSet.FieldByName( 'INVID' ).AsString );
    aDataSet.Next;
  end;
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB02.txtInvIdStExit(Sender: TObject);
begin
  txtInvIdEd.Text := txtInvIdSt.Text;
end;

{ ---------------------------------------------------------------------------- }

end.
