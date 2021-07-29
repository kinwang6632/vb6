unit frmInvA02_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DB, DBClient, ADODB, Grids, DBGrids,
  ComCtrls, cxControls, cxContainer, cxEdit, cxLabel, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee;

type
  TfrmInvA02_1 = class(TForm)
    Panel1: TPanel;
    btnCancel: TBitBtn;
    dsrChargeInfo33: TDataSource;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    btnOk: TBitBtn;
    Panel3: TPanel;
    Panel2: TPanel;
    dbgChargeInfo33: TDBGrid;
    dbgMdCustomer: TDBGrid;
    Panel4: TPanel;
    dbgChargeInfo34: TDBGrid;
    dsrChargeInfo34: TDataSource;
    dsrMdCustomer: TDataSource;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    lblCustID: TcxLabel;
    lblCustSName: TcxLabel;
    procedure FormShow(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FCustId: String;
    FBillId: String;
    FItemId: String;
    FTaxCode: String;
    FTaxRate: String;
    FResultDataSet: TClientDataSet;
    FTitleDataSet: TClientDataSet;
    function ValidateSelectedRecord(var aErrorMsg: String): Boolean;
    function GetCustomerName: String;
    procedure CreateFieldDefs;
    procedure UpdateSelectedRecord;
  public
    { Public declarations }
    constructor Create(const aCustId: String; const aBillId: String = '';
      const aItemId: String = ''; const aTaxCode: String = ''; const aTaxRate: String = '');
    property ResultDataSet: TClientDataSet read FResultDataSet;
    property TitleDataSet: TClientDataSet read FTitleDataSet;
  end;

var
  frmInvA02_1: TfrmInvA02_1;

implementation

uses cbUtilis, dtmSOU, dtmMainHU, dtmMainU, frmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

constructor TfrmInvA02_1.Create(const aCustId: String; const aBillId: String = '';
  const aItemId: String = ''; const aTaxCode: String = ''; const aTaxRate: String = '');
begin
  inherited Create( Application );
  FCustId := aCustId;
  FBillId := aBillId;
  FItemId := aItemId;
  FTaxCode := aTaxCode;
  FTaxRate := aTaxRate;
  FResultDataSet := TClientDataSet.Create( Self );
  FTitleDataSet := TClientDataSet.Create( Self );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.FormShow(Sender: TObject);
begin
  Caption := frmMain.GetFormTitleString( 'A02_1', '客戶抬頭與應收資料' );
  lblCustID.Caption := FCustId;
  lblCustSName.Caption := GetCustomerName;
  dtmSO.ActiveMdCustomerEx( FCustId );
  dtmSO.GetChargeInfo( FCustId, FBillId, FItemId, FTaxCode, FTaxRate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  dtmSO.adoMdCustomer.Close;
  dtmSO.adoChargeInfo33.Close;
  dtmSO.adoChargeInfo34.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.FormDestroy(Sender: TObject);
begin
  FResultDataSet.Data := Null;
  FResultDataSet.FieldDefs.Clear;
  FResultDataSet.Free;
  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  FTitleDataSet.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_1.GetCustomerName: String;
var
  aDataSet: TADOQuery;
begin
  Result := '';
  aDataSet := dtmSO.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
      ' SELECT CUSTNAME FROM %s.SO001 A   ' +
      '  WHERE A.CUSTID = ''%s''          ', [dtmMain.getMisDbOwner, FCustId] );
  aDataSet.Open;
  Result := aDataSet.FieldByName( 'CUSTNAME' ).AsString;
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.btnOkClick(Sender: TObject);
var
  aErrorMsg: String;
begin
  if not ValidateSelectedRecord( aErrorMsg ) then
  begin
    WarningMsg( aErrorMsg );
    Exit;
  end;
  CreateFieldDefs;
  UpdateSelectedRecord;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.btnCancelClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_1.ValidateSelectedRecord(var aErrorMsg: String): Boolean;
var
  aTaxCode, aRate1, aTargetTaxCode, aTargetRate1 : String;
  aIndex, aIndex2: Integer;
  aAccountNo, aInvSeqNo: String;
  aDBGrid: TDBGrid;
  aDataSet: TADOQuery;
begin
  Result := False;
  aErrorMsg := '';

  if ( dbgChargeInfo33.SelectedRows.Count +
       dbgChargeInfo34.SelectedRows.Count = 0 ) then
  begin
    aErrorMsg := '您尚未選取任何一筆應收資料!';
    Exit;
  end;

  aAccountNo := dbgMdCustomer.DataSource.DataSet.FieldByName('ACCOUNTNO').AsString;
  aInvSeqNo := dbgMdCustomer.DataSource.DataSet.FieldByName('INVSEQNO').AsString;

  //down, 檢查 SO033, SO034 的 account 是否與 title data 一樣...

  for aIndex := 1 to 2 do
  begin
    if ( aIndex = 1 ) then
      aDBGrid := dbgChargeInfo33
    else
      aDBGrid := dbgChargeInfo34;
    {}  
    aDataSet := TADOQuery( aDBGrid.DataSource.DataSet );
    for aIndex2 := 0 to aDBGrid.SelectedRows.Count - 1 do
    begin
      aDataSet.GotoBookmark( Pointer( aDBGrid.SelectedRows.Items[aIndex2] ) );
      if aAccountNo <> aDataSet.FieldByName( 'ACCOUNTNO' ).AsString then
      begin
        if ( aIndex = 1 ) then
        begin
          aErrorMsg := ( aErrorMsg +
            '您所選取之未日結資料的 Account No 與抬頭資料的 Account No 不一樣!'#13#10 );
        end else
        begin
          aErrorMsg := ( aErrorMsg +
            '您所選取之已日結資料的 Account No 與抬頭資料的 Account No 不一樣!'#13#10 );
        end;
        Break;
      end;
      {}
      if ( aInvSeqNo <> aDataSet.FieldByName( 'INVSEQNO' ).AsString ) then
      begin
        if ( aIndex = 1 ) then
        begin
          aErrorMsg := ( aErrorMsg +
            '您所選取之未日結資料的 發票抬頭流水號 與抬頭資料的 發票抬頭流水號 不一樣!'#13#10 );
        end else
        begin
          aErrorMsg := ( aErrorMsg +
            '您所選取之已日結資料的 發票抬頭流水號 與抬頭資料的 發票抬頭流水號 不一樣!'#13#10 );
        end;
        Break;
      end;
    end;
  end;

  aTargetTaxCode := '';
  aTargetRate1 := '';

  //down, 檢查 SO033與So034 的 稅率 是否一樣..
  for aIndex := 1 to 2 do
  begin
    case aIndex of
      1: aDBGrid := dbgChargeInfo33;
      2: aDBGrid := dbgChargeInfo34;
    end;
    aDataSet := TADOQuery( aDBGrid.DataSource.DataSet );
    for aIndex2 := 0 to aDBGrid.SelectedRows.Count - 1 do
    begin
      aDataSet.GotoBookmark( Pointer( aDBGrid.SelectedRows.Items[aIndex2] ) );
      aTaxCode := aDataSet.FieldByName( 'TAXCODE' ).AsString;
      aRate1 := aDataSet.FieldByName( 'RATE1' ).AsString;
      if aTargetTaxCode = '' then
      begin
        aTargetTaxCode := aTaxCode;
        aTargetRate1 := aRate1;
      end else
      begin
        if ( aTargetTaxCode <> aTaxCode) or ( aTargetRate1 <> aRate1 )then
        begin
          aErrorMsg := ( aErrorMsg + '您所選取之收費資料的稅別/稅率不一樣!'#13#10 );
          Break;
        end;
      end;
    end;
  end;

  Result := ( aErrorMsg = '' );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.CreateFieldDefs;
var
  aDataSet: TADOQuery;
  aDBGrid: TDBGrid;
  aIndex, aIndex2 : Integer;
begin
  FResultDataSet.Data := Null;
  FResultDataSet.FieldDefs.Clear;
  for aIndex := 1 to 2 do
  begin
    if ( aIndex = 1 ) then
      aDBGrid := dbgChargeInfo33
    else
      aDBGrid := dbgChargeInfo34;
    {}
    aDataSet := TADOQuery( aDBGrid.DataSource.DataSet );
    for aIndex2 := 0 to aDataSet.FieldCount - 1 do
    begin
      if FResultDataSet.FieldDefs.IndexOf( aDataSet.Fields[aIndex2].FieldName ) <= -1 then
      begin
        FResultDataSet.FieldDefs.Add( aDataSet.Fields[aIndex2].FieldName,
          aDataSet.Fields[aIndex2].DataType, aDataSet.Fields[aIndex2].Size );
      end;
    end;
  end;
  FResultDataSet.CreateDataSet;

  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  aDataSet := TADOQuery( dbgMdCustomer.DataSource.DataSet );
  for aIndex := 0 to aDataSet.FieldCount - 1 do
  begin
    if FTitleDataSet.FieldDefs.IndexOf( aDataSet.Fields[aIndex].FieldName ) <= -1 then
    begin
      FTitleDataSet.FieldDefs.Add( aDataSet.Fields[aIndex].FieldName,
        aDataSet.Fields[aIndex].DataType, aDataSet.Fields[aIndex].Size );
    end;
  end;
  FTitleDataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_1.UpdateSelectedRecord;
var
  aDataSet: TADOQuery;
  aDBGrid: TDBGrid;
  aIndex, aIndex2, aIndex3: Integer;
  aField: TField;
begin
  { 收費項目 }
  for aIndex := 1 to 2 do
  begin
    if ( aIndex = 1 ) then
      aDBGrid := dbgChargeInfo33
    else
      aDBGrid := dbgChargeInfo34;
    {}
    aDataSet := TADOQuery( aDBGrid.DataSource.DataSet );
    for aIndex2 := 0 to aDBGrid.SelectedRows.Count - 1 do
    begin
      aDataSet.GotoBookmark( Pointer( aDBGrid.SelectedRows.Items[aIndex2] ) );
      FResultDataSet.Append;
      for aIndex3 := 0 to aDataSet.FieldCount - 1 do
      begin
        aField := FResultDataSet.FindField( aDataSet.Fields[aIndex3].FieldName );
        if Assigned( aField ) then
          aField.Value := aDataSet.Fields[aIndex3].Value;
      end;
      FResultDataSet.Post;
    end;
  end;
  { 發票抬頭 }
  aDataSet := TADOQuery( dbgMdCustomer.DataSource.DataSet );
  FTitleDataSet.Append;
  for aIndex := 0 to aDataSet.FieldCount - 1 do
  begin
    aField := FTitleDataSet.FindField( aDataSet.Fields[aIndex].FieldName );
    if Assigned( aField ) then
      aField.Value := aDataSet.Fields[aIndex].Value;
  end;
  FTitleDataSet.Post;
end;

{ ---------------------------------------------------------------------------- }

end.
