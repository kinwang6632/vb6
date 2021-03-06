unit frmCD078B3_3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, DBClient, ADODB,
  cbStyleController, cbDBController, cbDataLookup, frmCD078B1U,
  dxSkinsDefaultPainters, dxSkinscxPCPainter, cxGraphics, 
  cxDataStorage, cxEdit, cxDBData, cxCurrencyEdit, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cxContainer, cxTextEdit, cxCheckBox, dxSkinsCore,
  cxStyles, cxCustomData, cxFilter, cxData, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TfrmCD078B3_3 = class(TForm)
    Bevel1: TBevel;
    btnSave: TButton;
    PTCode: TDataLookup;
    Label1: TLabel;
    CD078A2Grid: TcxGrid;
    gvPT: TcxGridDBTableView;
    gvPTCol1: TcxGridDBColumn;
    gvPTCol2: TcxGridDBColumn;
    gvPTCol3: TcxGridDBColumn;
    glPT: TcxGridLevel;
    dsCD078A2: TDataSource;
    btnUpdatePTCode: TBitBtn;
    btnDelPTCode: TBitBtn;
    Label2: TLabel;
    edtDepositAmt: TcxCurrencyEdit;
    Bevel2: TBevel;
    edtDefault: TcxCheckBox;
    gvPTCol4: TcxGridDBColumn;
    Deposit: TDataLookup;
    Label3: TLabel;
    gvPTCol5: TcxGridDBColumn;
    gvPTCol6: TcxGridDBColumn;
    gvPTCol7: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnUpdatePTCodeClick(Sender: TObject);
    procedure PTCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure PTCodeCodeNamePropertiesChange(Sender: TObject);
    procedure PTCodeCodeNoPropertiesChange(Sender: TObject);
    procedure btnDelPTCodeClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure DepositCodeNoPropertiesChange(Sender: TObject);
    procedure DepositCodeNamePropertiesChange(Sender: TObject);
    procedure DepositCodeNamePropertiesInitPopup(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyArray: array [1..4] of String;
    FReader: TADOQuery;
    FSourceDataSet: TClientDataSet;
    FCD078A2DataSet: TClientDataSet;
    FDepositAttr: Integer;
    function GetDefualtDepositAmount: Double;
    function GetAlreadyUsePTCode: String;
    function VdDepositAmt: Boolean;
    function VdDepositDefault(const AIsSave: Boolean): Boolean;
    procedure ClearInput;
    procedure RecordUpdate(ADeposit, ADepositName, APTCode, APTName: String; ADepositAmt: Double);
    procedure RecordDelete(const APTCode: String);
    procedure PrepareDataSet;
    procedure UnPrepareDataSet;
    procedure CopyFromSource;
    procedure SaveData;
  public
    { Public declarations }
    constructor Create(AMode: TDML; ADataSet: TClientDataSet; AKeys: array of String); reintroduce;
    property DepositAttr: Integer read FDepositAttr write FDepositAttr;
  end;

var
  frmCD078B3_3: TfrmCD078B3_3;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B3_3.Create(AMode: TDML; ADataSet: TClientDataSet; AKeys: array of String);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyArray[1] := aKeys[0]; { BPCode }
  FKeyArray[2] := aKeys[1]; { CItemCode }
  FKeyArray[3] := aKeys[2]; { ServiceType }
  FSourceDataSet := ADataSet;
  {}
  FReader := TADOQuery.Create( nil );
  FReader.Connection := DBController.DataConnection;
  FCD078A2DataSet := TClientDataSet.Create( nil );
  PrepareDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.FormCreate(Sender: TObject);
begin
  CopyFromSource;
  dsCD078A2.DataSet := FCD078A2DataSet;
  Deposit.Initializa;
  PTCode.Initializa;
  gvPT.ViewData.Expand( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.FormShow(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.FormDestroy(Sender: TObject);
begin
  FSourceDataSet := nil;
  UnPrepareDataSet;
  FCD078A2DataSet.Free;
  FReader.Free;
  Deposit.Finaliza;
  PTCode.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.btnSaveClick(Sender: TObject);
begin
  if not VdDepositDefault( True ) then Exit;
  SaveData;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.btnUpdatePTCodeClick(Sender: TObject);
begin
  if ( Deposit.Value = EmptyStr ) or
     ( Deposit.ValueName = EmptyStr ) or
     ( PTCode.Value = EmptyStr ) or
     ( PTCode.ValueName = EmptyStr ) then
    Exit;
  if not VdDepositAmt then Exit;
  if edtDefault.Checked then
  begin
    if not VdDepositDefault( False ) then Exit;
  end;
  RecordUpdate( Deposit.Value, Deposit.ValueName, PTCode.Value,
    PTCode.ValueName, edtDepositAmt.Value );
  ClearInput;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.btnDelPTCodeClick(Sender: TObject);
begin
  RecordDelete( FCD078A2DataSet.FieldByName( 'PTCode' ).AsString );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.DepositCodeNoPropertiesChange(Sender: TObject);
begin
  Deposit.CodeNoPropertiesChange( Sender );
  edtDepositAmt.Value := GetDefualtDepositAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.DepositCodeNamePropertiesChange(Sender: TObject);
begin
  Deposit.CodeNamePropertiesChange( Sender );
  edtDepositAmt.Value := GetDefualtDepositAmount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.DepositCodeNamePropertiesInitPopup(Sender: TObject);
var
  aRefCode: String;
begin
  case FDepositAttr of
    1: aRefCode := ' ''15'', ''16''  ';
    2: aRefCode := ' ''14'' ';
  else
    Exit;
  end;
  {}
  Deposit.SQL.Text := Format(
    '  SELECT A.CODENO, A.DESCRIPTION         ' +
    '    FROM %s.CD019 A                      ' +
    '   WHERE A.PERIODFLAG = ''0''            ' +
    '     AND A.SIGN = ''+''                  ' +
    '     AND A.REFNO IN ( %s )               ' +
    '     AND A.SERVICETYPE = ''%s''          ' +
    '     AND A.STOPFLAG = ''0''              ' +
    '   ORDER BY A.CODENO                     ',
    [DBController.LoginInfo.DbAccount, aRefCode, FKeyArray[3]] );
  Deposit.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.PTCodeCodeNoPropertiesChange(Sender: TObject);
begin
  PTCode.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.PTCodeCodeNamePropertiesChange(Sender: TObject);
begin
  PTCode.CodeNamePropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.PTCodeCodeNamePropertiesInitPopup(Sender: TObject);
var
  aUsePTCode: String;
begin
  aUsePTCode := QuotedValue( GetAlreadyUsePTCode );
  PTCode.SQL.Text := Format(
    '  SELECT A.CODENO, A.DESCRIPTION         ' +
    '    FROM %s.CD032 A                      ' +
    '   WHERE A.STOPFLAG = ''0''              ' +
    '     AND A.CODENO NOT IN ( %s )          ' +
    '   ORDER BY A.CODENO                     ',
    [DBController.LoginInfo.DbAccount, aUsePTCode] );
  PTCode.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.ClearInput;
begin
  PTCode.Clear;
  Deposit.Clear;
  edtDepositAmt.Value := 0;
  edtDefault.Checked := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.RecordUpdate(ADeposit, ADepositName, APTCode, APTName: String;
  ADepositAmt: Double);
var
  aFound: Boolean;
begin
  if ( ADeposit = EmptyStr ) or
     ( APTCode = EmptyStr ) or
     ( APTName = EmptyStr ) or
     ( ADepositAmt <= 0 ) then
   Exit;
 { Detail }  
  aFound := FCD078A2DataSet.Locate( 'BPCode;CItemCode;ServiceType;DepositCode;PTCode',
    VarArrayOf( [FKeyArray[1], FKeyArray[2], FKeyArray[3], ADeposit, APTCode] ), [] );
  if ( aFound ) then
    FCD078A2DataSet.Edit
  else
    FCD078A2DataSet.Append;
  FCD078A2DataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
  FCD078A2DataSet.FieldByName( 'CItemCode' ).AsString := FKeyArray[2];
  FCD078A2DataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
  FCD078A2DataSet.FieldByName( 'DepositCode' ).AsString := ADeposit;
  FCD078A2DataSet.FieldByName( 'DepositName' ).AsString := ADepositName;
  FCD078A2DataSet.FieldByName( 'PTCode' ).AsString := APTCode;
  FCD078A2DataSet.FieldByName( 'PTName' ).AsString := APTName;
  FCD078A2DataSet.FieldByName( 'PTGroup' ).AsString := Format( '%s - %s', [APTCode, APTName] );
  FCD078A2DataSet.FieldByName( 'DepositAmt' ).AsFloat := ADepositAmt;
  FCD078A2DataSet.FieldByName( 'DepositDefault' ).AsString := IIF( edtDefault.Checked, '1', '0' );
  FCD078A2DataSet.Post;
  gvPT.ViewData.Expand( True );
  ActiveControl := CD078A2Grid;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.RecordDelete(const APTCode: String);
var
  aFound: Boolean;
begin
  if ( APTCode <> EmptyStr ) then
  begin
    aFound := FCD078A2DataSet.Locate( 'BPCode;CItemCode;ServiceType;PTCode',
      VarArrayOf( [FKeyArray[1], FKeyArray[2], FKeyArray[3], APTCode] ), [] );
    if ( aFound ) then FCD078A2DataSet.Delete;
  end;
  ActiveControl := CD078A2Grid;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_3.GetDefualtDepositAmount: Double;
begin
  FReader.Close;
  FReader.SQL.Text := Format(
    '  SELECT NVL( A.AMOUNT, 0 ) FROM %S.CD019 A ' +
    '   WHERE A.CODENO = ''%S''                  ',
    [DBController.LoginInfo.DbAccount, Deposit.Value] );
  FReader.Open;
  Result := FReader.Fields[0].AsFloat;
  FReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_3.GetAlreadyUsePTCode: String;
var
  aBMark: TBookmark;
begin
  Result := EmptyStr;
  aBMark := FCD078A2DataSet.GetBookmark;
  try
    FCD078A2DataSet.DisableConstraints;
    try
      FCD078A2DataSet.First;
      while not ( FCD078A2DataSet.Eof ) do
      begin
        Result := AddValue( FCD078A2DataSet.FieldByName( 'PTCode' ).AsString, Result );
        FCD078A2DataSet.Next;
      end;
      FCD078A2DataSet.GotoBookmark( aBMark );
    finally
      FCD078A2DataSet.EnableControls;
    end;
  finally
    FCD078A2DataSet.FreeBookmark( aBMark );
  end;
  if ( Result = EmptyStr ) then Result := '-1';
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_3.VdDepositAmt: Boolean;
begin
  Result := False;
  if ( edtDepositAmt.Value <= 0 ) then
  begin
    WarningMsg( '?i?O?????j???B???????J,?B???i???t???C' );
    if ( edtDepositAmt.CanFocusEx ) then edtDepositAmt.SetFocus;
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_3.VdDepositDefault(const AIsSave: Boolean): Boolean;
var
  aDefaultCount: Integer;
begin
  Result := False;
  aDefaultCount := 0;
  FCD078A2DataSet.DisableControls;
  try
    FCD078A2DataSet.First;
    while not FCD078A2DataSet.Eof do
    begin
      FCD078A2DataSet.First;
      while not FCD078A2DataSet.Eof do
      begin
        if ( FCD078A2DataSet.FieldByName( 'DepositDefault' ).AsString = '1' ) then
          Inc( aDefaultCount );
        FCD078A2DataSet.Next;
      end;
    end;
  finally
    FCD078A2DataSet.EnableControls;
  end;
  if ( AIsSave ) then
  begin
    if ( aDefaultCount <= 0 ) then
    begin
      WarningMsg( '???????w?@???i?w?]?I???O?????j?C' );
      Exit;
    end else
    if ( aDefaultCount >= 2 ) then
    begin
      WarningMsg( '?i?w?]?I???O?????j???i????????,?i?w?]?I???O?????j?u?i?]?w?@???C' );
      Exit;
    end;
  end else
  if ( aDefaultCount >= 1 ) then
  begin
    WarningMsg(  '?i?w?]?I???O?????j???i????????,?i?w?]?I???O?????j?u?i?]?w?@???C' );
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.PrepareDataSet;
begin
  { Detail DataSet }
  if ( FCD078A2DataSet.FieldDefs.Count <= 0 ) then
  begin
    FCD078A2DataSet.FieldDefs.Add( 'BPCode', ftString, 10 );
    FCD078A2DataSet.FieldDefs.Add( 'CitemCode', ftString, 5 );
    FCD078A2DataSet.FieldDefs.Add( 'ServiceType', ftString, 1 );
    FCD078A2DataSet.FieldDefs.Add( 'DepositCode', ftString, 5 );
    FCD078A2DataSet.FieldDefs.Add( 'DepositName', ftString, 40 );
    FCD078A2DataSet.FieldDefs.Add( 'PTCode', ftString, 3 );
    FCD078A2DataSet.FieldDefs.Add( 'PTName', ftString, 20 );
    FCD078A2DataSet.FieldDefs.Add( 'PTGroup', ftString, 100 );
    FCD078A2DataSet.FieldDefs.Add( 'DepositAmt', ftFloat );
    FCD078A2DataSet.FieldDefs.Add( 'DepositDefault', ftString, 1 );
    FCD078A2DataSet.CreateDataSet;
    FCD078A2DataSet.AddIndex( 'IDX1', 'BPCode;CitemCode;ServiceType;DepositCode;PTCode', [ixUnique] );
    FCD078A2DataSet.IndexName := 'IDX1';
  end;
  FCD078A2DataSet.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.UnPrepareDataSet;
begin
  FCD078A2DataSet.EmptyDataSet;
  if ( FCD078A2DataSet.IndexName <> EmptyStr ) then
    FCD078A2DataSet.DeleteIndex( FCD078A2DataSet.IndexName );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.CopyFromSource;
var
  aPTCode, aPTName: String;
begin
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
  begin
    { Detail }
    aPTCode := FSourceDataSet.FieldByName( 'PTCode' ).AsString;
    aPTName := FSourceDataSet.FieldByName( 'PTName' ).AsString;
    FCD078A2DataSet.Append;
    FCD078A2DataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
    FCD078A2DataSet.FieldByName( 'CItemCode' ).AsString := FKeyArray[2];
    FCD078A2DataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
    FCD078A2DataSet.FieldByName( 'DepositCode' ).AsString := FSourceDataSet.FieldByName( 'DepositCode' ).AsString;
    FCD078A2DataSet.FieldByName( 'DepositName' ).AsString := FSourceDataSet.FieldByName( 'DepositName' ).AsString;;
    FCD078A2DataSet.FieldByName( 'PTCode' ).AsString := aPTCode;
    FCD078A2DataSet.FieldByName( 'PTName' ).AsString := aPTName;
    FCD078A2DataSet.FieldByName( 'PTGroup' ).AsString := Format( '%s - %s', [aPTCode, aPTName] );
    FCD078A2DataSet.FieldByName( 'DepositAmt' ).AsFloat := FSourceDataSet.FieldByName( 'DepositAmt' ).AsFloat;
    FCD078A2DataSet.FieldByName( 'DepositDefault' ).AsFloat := FSourceDataSet.FieldByName( 'DepositDefault' ).AsFloat;
    FCD078A2DataSet.Post;
    FSourceDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_3.SaveData;
begin
  { ???????R??, ?o???????O???]?w Filter, ???H???|?h?R??????????????  }
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
    FSourceDataSet.Delete;
  { ?N?????]?w?????@???? Source }  
  FCD078A2DataSet.First;
  while not FCD078A2DataSet.Eof do
  begin
    FSourceDataSet.Append;
    FSourceDataSet.FieldByName( 'BPCode' ).AsString := FCD078A2DataSet.FieldByName( 'BPCode' ).AsString;
    FSourceDataSet.FieldByName( 'CItemCode' ).AsString := FCD078A2DataSet.FieldByName( 'CItemCode' ).AsString;
    FSourceDataSet.FieldByName( 'ServiceType' ).AsString := FCD078A2DataSet.FieldByName( 'ServiceType' ).AsString;
    FSourceDataSet.FieldByName( 'DepositCode' ).AsString := FCD078A2DataSet.FieldByName( 'DepositCode' ).AsString;
    FSourceDataSet.FieldByName( 'DepositName' ).AsString := FCD078A2DataSet.FieldByName( 'DepositName' ).AsString;
    FSourceDataSet.FieldByName( 'PTCode' ).AsString := FCD078A2DataSet.FieldByName( 'PTCode' ).AsString;
    FSourceDataSet.FieldByName( 'PTName' ).AsString := FCD078A2DataSet.FieldByName( 'PTName' ).AsString;
    FSourceDataSet.FieldByName( 'DepositAmt' ).AsFloat := FCD078A2DataSet.FieldByName( 'DepositAmt' ).AsFloat;
    FSourceDataSet.FieldByName( 'DepositDefault' ).AsString := FCD078A2DataSet.FieldByName( 'DepositDefault' ).AsString;
    FSourceDataSet.Post;
    FCD078A2DataSet.Next;
  end; 
end;

{ ---------------------------------------------------------------------------- }

end.
