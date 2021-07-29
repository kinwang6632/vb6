unit frmSearchCustU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, DBClient, StdCtrls, ComCtrls, ExtCtrls, ActnList, dtmMainU,
  cxLookAndFeelPainters, cxButtons, cxCheckBox, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxRadioGroup, Menus, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxDBData, cxGridLevel,
  cxClasses, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxGridBandedTableView, cxGridDBBandedTableView,
  Buttons, dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmSearchCust = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    rdoCustId: TcxRadioButton;
    rdoCustSName: TcxRadioButton;
    txtValue: TcxTextEdit;
    btnSearch: TcxButton;
    rdoTel: TcxRadioButton;
    rdoContactee: TcxRadioButton;
    rdoBusinessId: TcxRadioButton;
    rdoInvAddr: TcxRadioButton;
    LvResult: TcxGridLevel;
    ResultGrid: TcxGrid;
    Bevel1: TBevel;
    DataSource1: TDataSource;
    TvResult: TcxGridDBBandedTableView;
    TvResultCUSTID: TcxGridDBBandedColumn;
    TvResultCUSTSNAME: TcxGridDBBandedColumn;
    TvResultCUSTNAME: TcxGridDBBandedColumn;
    TvResultMAILADDR: TcxGridDBBandedColumn;
    TvResultTEL1: TcxGridDBBandedColumn;
    TvResultTEL2: TcxGridDBBandedColumn;
    TvResultTEL3: TcxGridDBBandedColumn;
    TvResultAPPCONTACTEE1: TcxGridDBBandedColumn;
    TvResultAPPCONTACTEE2: TcxGridDBBandedColumn;
    TvResultFINACONTACTEE1: TcxGridDBBandedColumn;
    TvResultFINACONTACTEE2: TcxGridDBBandedColumn;
    TvResultTITLESNAME: TcxGridDBBandedColumn;
    TvResultTITLENAME: TcxGridDBBandedColumn;
    TvResultBUSINESSID: TcxGridDBBandedColumn;
    TvResultINVADDR: TcxGridDBBandedColumn;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    Bevel2: TBevel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure TvResultDblClick(Sender: TObject);
  private
    { Private declarations }
    FCustId: String;
    FTitleDataSet: TClientDataSet;
    function GetSearchSql: String;
    function Search: Boolean;
    function CheckCanSearch: Boolean;
    procedure CreateFieldDefs;
    procedure UpdateSelectedRecord;
  public
    { Public declarations }
    constructor Create(const aCustId: String); overload;
    property TitleDataSet: TClientDataSet read FTitleDataSet;
  end;

var
  frmSearchCust: TfrmSearchCust;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

constructor TfrmSearchCust.Create(const aCustId: String);
begin
  inherited Create( Application );
  FCustId := aCustId;
  FTitleDataSet := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.FormCreate(Sender: TObject);
begin
  rdoCustId.Checked := True;
  txtValue.Text := EmptyStr;
  dtmMain.adoComm.Close;
  {Cancel 1000 record limit By Kin 2016/02/16}
//  dtmMain.adoComm.MaxRecords := 1000;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.FormShow(Sender: TObject);
begin
  txtValue.Text := FCustId;
  btnOk.Enabled := False;
  if ( txtValue.Text <> EmptyStr ) then btnOK.Enabled := Search;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.FormDestroy(Sender: TObject);
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.MaxRecords := 0;
  if not VarIsNull( FTitleDataSet.Data ) then FTitleDataSet.EmptyDataSet;
  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  FTitleDataSet.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfrmSearchCust.GetSearchSql: String;
var
  aBody, aWhere, aOrder, aValue: String;
begin
  aValue := '%' + txtValue.Text + '%'; 
  aBody := Format(
    ' SELECT A.CUSTID,                 ' +
    '        A.CUSTSNAME,              ' +
    '        A.CUSTNAME,               ' +
    '        B.MAILADDR,               ' +
    '        A.TEL1,                   ' +
    '        A.TEL2,                   ' +
    '        A.TEL3,                   ' +
    '        A.APPCONTACTEE1,          ' +
    '        A.APPCONTACTEE2,          ' +
    '        A.FINACONTACTEE1,         ' +
    '        A.FINACONTACTEE2,         ' +
    '        B.TITLESNAME,             ' +
    '        B.TITLENAME,              ' +
    '        B.BUSINESSID,             ' +
    '        B.MAILADDR AS MAILADDR2,  ' +
    '        B.INVADDR,                ' +
    '        B.MEMO,                   ' +
    '        B.MZIPCODE                ' +
    '   FROM INV002 A, INV019 B        ' +
    '  WHERE A.IDENTIFYID1 = B.IDENTIFYID1   ' +
    '    AND A.IDENTIFYID2 = B.IDENTIFYID2   ' +
    '    AND A.COMPID = B.COMPID             ' +
    '    AND A.CUSTID = B.CUSTID             ' +
    '    AND A.IDENTIFYID1 = ''%s''    ' +
    '    AND A.IDENTIFYID2 = ''%s''    ' +
    '    AND A.COMPID = ''%s''         ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );


  { 客編 }
  if ( rdoCustId.Checked ) then
  begin
    aWhere := Format( ' AND A.CUSTID LIKE ''%s''  ', [aValue] );
    aOrder := ' ORDER BY A.CUSTID '; 
  end;  
  { 簡稱/全名 }
  if ( rdoCustSName.Checked ) then
  begin
    aWhere := Format(
      ' AND ( A.CUSTSNAME LIKE ''%s'' OR   ' +
      '       A.CUSTNAME LIKE ''%s''  OR   ' +
      '       B.TITLESNAME LIKE ''%s'' OR  ' +
      '       B.TITLENAME LIKE ''%s''  )   ', [aValue, aValue, aValue, aValue] );
    aOrder := ' ORDER BY A.CUSTSNAME ';
  end;    
  { 連絡人 }
  if ( rdoContactee.Checked ) then
  begin
    aWhere := Format(
      ' AND ( A.APPCONTACTEE1 LIKE ''%s''  OR   ' +
      '       A.APPCONTACTEE2 LIKE ''%s''  OR   ' +
      '       A.FINACONTACTEE1 LIKE ''%s''  OR  ' +
      '       A.FINACONTACTEE2 LIKE ''%s'' )    ',  [aValue, aValue, aValue, aValue] );
    aOrder := ' ORDER BY A.APPCONTACTEE1 ';
  end;
  { 統編 }
  if ( rdoBusinessId.Checked ) then
  begin
    aWhere := Format( ' AND B.BUSINESSID LIKE ''%s''  ', [aValue] );
    aOrder := ' ORDER BY B.BUSINESSID ';
  end;
  { 電話 }
  if ( rdoTel.Checked ) then
  begin
    aWhere := Format(
      ' AND ( A.TEL1 LIKE ''%s'' OR    ' +
      '       A.TEL2 LIKE ''%s'' OR    ' +
      '       A.TEL3 LIKE ''%s''  )    ', [aValue, aValue, aValue] );
    aOrder := ' ORDER BY A.TEL1 ';
  end;
  { 發票/郵寄 }
  if ( rdoInvAddr.Checked ) then
  begin
    aWhere := Format(
      ' AND ( B.MAILADDR LIKE ''%s'' OR  ' +
      '       B.INVADDR LIKE ''%s'' )    ', [aValue, aValue] );
    aOrder := ' ORDER BY A.MAILADDR ';
  end;
  Result := ( aBody + aWhere + aOrder );
end;

{ ---------------------------------------------------------------------------- }

function TfrmSearchCust.Search: Boolean;
begin
  Result := CheckCanSearch;
  if Result then
  begin
    Screen.Cursor := crSQLWait;
    try
      dtmMain.adoComm.Close;
      dtmMain.adoComm.SQL.Text := GetSearchSql;
      dtmMain.adoComm.Open;
      Result := not dtmMain.adoComm.IsEmpty;
    finally
      Screen.Cursor := crDefault;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TfrmSearchCust.CheckCanSearch: Boolean;
begin
  Result := False;
  if Trim( txtValue.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入搜尋值。' );
    if ( txtValue.CanFocusEx ) then txtValue.SetFocus;
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.btnSearchClick(Sender: TObject);
begin
  btnOK.Enabled := False;
  btnOK.Enabled := Search;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.btnOkClick(Sender: TObject);
begin
  CreateFieldDefs;
  UpdateSelectedRecord;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.btnCancelClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSearchCust.TvResultDblClick(Sender: TObject);
begin
  btnOkClick( btnOk );
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmSearchCust.CreateFieldDefs;
var
  aIndex: Integer;
  aDataSet: TADOQuery;
begin
  FTitleDataSet.Data := Null;
  FTitleDataSet.FieldDefs.Clear;
  aDataSet := TADOQuery( DataSource1.DataSet );
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

procedure TfrmSearchCust.UpdateSelectedRecord;
var
  aIndex: Integer;
  aDataSet: TADOQuery;
  aField: TField;
begin
  aDataSet := TADOQuery( DataSource1.DataSet );
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
