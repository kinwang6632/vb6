unit frmCD078B4_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBClient, ADODB, StdCtrls, ExtCtrls, ImgList,
  cbStyleController, cbDBController, cbDataLookup, frmCD078B1U, 
  dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, 
  cxGraphics, cxDataStorage, cxEdit, cxDBData, cxTextEdit,
  Buttons, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxControls, cxGridCustomView, cxGrid, cxGridCommon,
  cxStyles, cxCustomData, cxFilter, cxData, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;


type
  TfrmCD078B4_1 = class(TForm)
    Bevel1: TBevel;
    Label1: TLabel;
    Bevel2: TBevel;
    btnSave: TButton;
    InstCode: TDataLookup;
    InstGrid: TcxGrid;
    gvInst: TcxGridDBTableView;
    gvInstCol1: TcxGridDBColumn;
    gvInstCol2: TcxGridDBColumn;
    glInst: TcxGridLevel;
    btnUpdatePTCode: TBitBtn;
    btnDelPTCode: TBitBtn;
    dsCD078B1: TDataSource;
    Label2: TLabel;
    PRCode2: TDataLookup;
    gvInstCol3: TcxGridDBColumn;
    gvInstCol4: TcxGridDBColumn;
    ImageList1: TImageList;
    procedure FormCreate(Sender: TObject);
    procedure PTCodeCodeNoPropertiesChange(Sender: TObject);
    procedure InstCodeCodeNamePropertiesChange(Sender: TObject);
    procedure InstCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnUpdatePTCodeClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnDelPTCodeClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure PRCode2CodeNamePropertiesInitPopup(Sender: TObject);
    procedure PRCode2CodeNamePropertiesChange(Sender: TObject);
    procedure PRCode2CodeNoPropertiesChange(Sender: TObject);
    procedure gvInstCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
  private
    { Private declarations }
    FMode: TDML;

    FKeyArray: array [1..3] of String;
    FSourceDataSet: TClientDataSet;
    FCD078B1DataSet: TClientDataSet;
    FReader: TADOQuery;
    procedure ClearInput;
    procedure RecordUpdate(const AInstCode2, AInstName2, APrCode2, APrName2: String);
    procedure RecordDelete;
    procedure PrepareDataSet;
    procedure UnPrepareDataSet;
    procedure CopyFromSource;
    procedure SaveData;
    function SameWithMasterInstCode: Boolean; overload;
    function SameWithMasterInstCode(const AInstCode: String): Boolean; overload;
    function GetAlreadyUseInstCode: String;
    function GetAlreadyUsePrCode: String;
  public
    { Public declarations }
    constructor Create(AMode: TDML; ADataSet: TClientDataSet; AKeys: array of String); reintroduce;
  end;

var
  frmCD078B4_1: TfrmCD078B4_1;

implementation

uses cbUtilis;

{$R *.dfm}

{ TForm1 }

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B4_1.Create(AMode: TDML; ADataSet: TClientDataSet;
  AKeys: array of String);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyArray[1] := aKeys[0]; { BPCode }
  FKeyArray[2] := aKeys[1]; { InstCode }
  FKeyArray[3] := aKeys[2]; { ServiceType }
  FSourceDataSet := ADataSet;
  FReader := TADOQuery.Create( nil );
  FReader.Connection := DBController.DataConnection;
  FCD078B1DataSet := TClientDataSet.Create( nil );
  PrepareDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.FormCreate(Sender: TObject);
begin
  CopyFromSource;
  dsCD078B1.DataSet := FCD078B1DataSet;
  InstCode.Initializa;
  PRCode2.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.FormDestroy(Sender: TObject);
begin
  FSourceDataSet := nil;
  UnPrepareDataSet;
  FCD078B1DataSet.Free;
  FReader.Free;
  InstCode.Finaliza;
  PRCode2.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.UnPrepareDataSet;
begin
  FCD078B1DataSet.EmptyDataSet;
  if ( FCD078B1DataSet.IndexName <> EmptyStr ) then
    FCD078B1DataSet.DeleteIndex( FCD078B1DataSet.IndexName );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.PRCode2CodeNamePropertiesChange(Sender: TObject);
begin
  PrCode2.CodeNamePropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.PRCode2CodeNamePropertiesInitPopup(Sender: TObject);
var
  aUsePrCode: String;
begin
  Screen.Cursor := crSQLWait;
  try
    aUsePrCode := QuotedValue( GetAlreadyUsePrCode );
    if ( InstCode.Value = EmptyStr ) then
    begin
      PRCode2.SQL.Text := Format(
        ' SELECT CODENO, DESCRIPTION ' +
        '   FROM %s.CD007            ' +
        '  WHERE 1 = 2               ',
        [DBController.LoginInfo.DbAccount] );
    end else
    begin
      PRCode2.SQL.Text := Format(
        ' SELECT A.CODENO, A.DESCRIPTION   ' +
        '   FROM %s.CD007 A                ' +
        '  WHERE A.SERVICETYPE =''%s''     ' +
        '    AND A.STOPFLAG = 0            ' +
        '    AND A.REFNO IN ( ''10'' )     ' +
        '    AND A.CODENO NOT IN ( %s )    ' +
        '  ORDER BY A.CODENO               ',
        [DBController.LoginInfo.DbAccount, FKeyArray[3], aUsePrCode] );
    end;
    PRCode2.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.PrepareDataSet;
begin
  if ( FCD078B1DataSet.FieldDefs.Count <= 0 ) then
  begin
    FCD078B1DataSet.FieldDefs.Add( 'BPCode', ftString, 10 );
    FCD078B1DataSet.FieldDefs.Add( 'InstCode', ftString, 3 );
    FCD078B1DataSet.FieldDefs.Add( 'ServiceType', ftString, 1 );
    FCD078B1DataSet.FieldDefs.Add( 'InstCode2', ftString, 3 );
    FCD078B1DataSet.FieldDefs.Add( 'InstName2', ftString, 40 );
    FCD078B1DataSet.FieldDefs.Add( 'PrCode2', ftString, 3 );
    FCD078B1DataSet.FieldDefs.Add( 'PrName2', ftString, 40 );
    FCD078B1DataSet.CreateDataSet;
    FCD078B1DataSet.AddIndex( 'IDX1', 'BPCode;InstCode;ServiceType;InstCode2', [ixUnique] );
    FCD078B1DataSet.IndexName := 'IDX1';
  end;
  FCD078B1DataSet.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.RecordDelete;
begin
  if ( not FCD078B1DataSet.IsEmpty ) then
  begin
    if not SameWithMasterInstCode then
      FCD078B1DataSet.Delete;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.ClearInput;
begin
  InstCode.Clear;
  PRCode2.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.CopyFromSource;
begin
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
  begin
    FCD078B1DataSet.Append;
    FCD078B1DataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
    FCD078B1DataSet.FieldByName( 'InstCode' ).AsString := FKeyArray[2];
    FCD078B1DataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
    FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString := FSourceDataSet.FieldByName( 'InstCode2' ).AsString;
    FCD078B1DataSet.FieldByName( 'InstName2' ).AsString := FSourceDataSet.FieldByName( 'InstName2' ).AsString;
    FCD078B1DataSet.FieldByName( 'PrCode2' ).AsString := FSourceDataSet.FieldByName( 'PrCode2' ).AsString;
    FCD078B1DataSet.FieldByName( 'PrName2' ).AsString := FSourceDataSet.FieldByName( 'PrName2' ).AsString;
    FCD078B1DataSet.Post;
    FSourceDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.RecordUpdate(const AInstCode2, AInstName2, APrCode2, APrName2: String);
var
  aFound: Boolean;
begin
  if ( AInstCode2 = EmptyStr ) or
     ( AInstName2 = EmptyStr ) then
   Exit;
  aFound := FCD078B1DataSet.Locate( 'BPCode;InstCode;ServiceType;InstCode2',
    VarArrayOf( [FKeyArray[1], FKeyArray[2], FKeyArray[3], AInstCode2] ), [] );
  if ( aFound ) then
    FCD078B1DataSet.Edit
  else
    FCD078B1DataSet.Append;
  FCD078B1DataSet.FieldByName( 'BPCode' ).AsString := FKeyArray[1];
  FCD078B1DataSet.FieldByName( 'InstCode' ).AsString := FKeyArray[2];
  FCD078B1DataSet.FieldByName( 'ServiceType' ).AsString := FKeyArray[3];
  FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString := AInstCode2;
  FCD078B1DataSet.FieldByName( 'InstName2' ).AsString := AInstName2;
  FCD078B1DataSet.FieldByName( 'PrCode2' ).AsString := APrCode2;
  FCD078B1DataSet.FieldByName( 'PrName2' ).AsString := APrName2;
  FCD078B1DataSet.Post;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.SaveData;
begin
  { ???????R??, ?o???????O???]?w Filter, ???H???|?h?R??????????????  }
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
    FSourceDataSet.Delete;
  { ?N?????]?w?????@???? Source }
  FCD078B1DataSet.First;
  while not FCD078B1DataSet.Eof do
  begin
    FSourceDataSet.Append;
    FSourceDataSet.FieldByName( 'BPCode' ).AsString := FCD078B1DataSet.FieldByName( 'BPCode' ).AsString;
    FSourceDataSet.FieldByName( 'InstCode' ).AsString := FCD078B1DataSet.FieldByName( 'InstCode' ).AsString;
    FSourceDataSet.FieldByName( 'ServiceType' ).AsString := FCD078B1DataSet.FieldByName( 'ServiceType' ).AsString;
    FSourceDataSet.FieldByName( 'InstCode2' ).AsString := FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString;
    FSourceDataSet.FieldByName( 'InstName2' ).AsString := FCD078B1DataSet.FieldByName( 'InstName2' ).AsString;
    FSourceDataSet.FieldByName( 'PrCode2' ).AsString := FCD078B1DataSet.FieldByName( 'PrCode2' ).AsString;
    FSourceDataSet.FieldByName( 'PrName2' ).AsString := FCD078B1DataSet.FieldByName( 'PrName2' ).AsString;
    FSourceDataSet.Post;
    FCD078B1DataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4_1.SameWithMasterInstCode: Boolean;
begin
  Result := False;
  if ( not FCD078B1DataSet.IsEmpty ) then
    Result := SameWithMasterInstCode( FCD078B1DataSet.FieldByName( 'INSTCODE2' ).AsString );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4_1.SameWithMasterInstCode(const AInstCode: String): Boolean;
begin
  Result := ( AInstCode = FKeyArray[2] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4_1.GetAlreadyUseInstCode: String;
var
  aBMark: TBookmark;
begin
  { ???????v?? InstCode ???i?H?Q?? }
  Result := FKeyArray[2] + ',';
  aBMark := FCD078B1DataSet.GetBookmark;
  try
    FCD078B1DataSet.DisableConstraints;
    try
      FCD078B1DataSet.First;
      while not ( FCD078B1DataSet.Eof ) do
      begin
        Result := ( Result + FCD078B1DataSet.FieldByName( 'InstCode2' ).AsString + ',' );
        FCD078B1DataSet.Next;
      end;
      FCD078B1DataSet.GotoBookmark( aBMark );
    finally
      FCD078B1DataSet.EnableControls;
    end;
  finally
    FCD078B1DataSet.FreeBookmark( aBMark );
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B4_1.GetAlreadyUsePrCode: String;
var
  aBMark: TBookmark;
begin
  aBMark := FCD078B1DataSet.GetBookmark;
  try
    FCD078B1DataSet.DisableConstraints;
    try
      FCD078B1DataSet.First;
      while not ( FCD078B1DataSet.Eof ) do
      begin
        if ( FCD078B1DataSet.FieldByName( 'PrCode2' ).AsString <> EmptyStr ) then
          Result := ( Result + FCD078B1DataSet.FieldByName( 'PrCode2' ).AsString + ',' );
        FCD078B1DataSet.Next;
      end;
      FCD078B1DataSet.GotoBookmark( aBMark );
    finally
      FCD078B1DataSet.EnableControls;
    end;
  finally
    FCD078B1DataSet.FreeBookmark( aBMark );
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
  if ( Result = EmptyStr ) then
    Result := '-1';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.PRCode2CodeNoPropertiesChange(Sender: TObject);
begin
  PrCode2.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.PTCodeCodeNoPropertiesChange(Sender: TObject);
begin
  InstCode.CodeNoPropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.InstCodeCodeNamePropertiesChange(Sender: TObject);
begin
  InstCode.CodeNamePropertiesChange( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.InstCodeCodeNamePropertiesInitPopup(Sender: TObject);
var
  aUseInstCode: String;
begin
  Screen.Cursor := crSQLWait;
  try
    aUseInstCode := QuotedValue( GetAlreadyUseInstCode );
    InstCode.SQL.Text := Format(
      '  SELECT A.CODENO, A.DESCRIPTION         ' +
      '    FROM %s.CD005 A                      ' +
      '   WHERE A.STOPFLAG = ''0''              ' +
      '     AND A.SERVICETYPE = ''%s''          ' +
      '     AND A.CODENO NOT IN ( %s )          ' +
      '     AND A.REFNO = ''16''                ' +
      '   ORDER BY A.CODENO                     ',
      [DBController.LoginInfo.DbAccount, FKeyArray[3], aUseInstCode] );
    InstCode.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.btnUpdatePTCodeClick(Sender: TObject);
begin
  if ( InstCode.Value = EmptyStr ) or
     ( InstCode.ValueName = EmptyStr ) then
    Exit;
  RecordUpdate( InstCode.Value, InstCode.ValueName, PRCode2.Value, PRCode2.ValueName );
  ClearInput;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.btnDelPTCodeClick(Sender: TObject);
begin
  RecordDelete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.btnSaveClick(Sender: TObject);
begin
  SaveData;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B4_1.gvInstCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aRect: TRect;
  aInstCode2, aText: String;
  aDrawFlag: Integer;
begin
  aInstCode2 := AViewInfo.GridRecord.DisplayTexts[gvInstCol1.Index];
  if SameWithMasterInstCode( aInstCode2 ) then
  begin
    ACanvas.FillRect( AViewInfo.ContentBounds );
    if ( AViewInfo.Item = gvInstCol1 ) then
    begin
      aRect := AViewInfo.ContentBounds;
      InflateRect( aRect, -2, 0 );
      ACanvas.DrawImage( ImageList1, aRect.Left, aRect.Top, 0 );
      InflateRect( aRect, ImageList1.Width + 4, 0 );
      aText := AViewInfo.GridRecord.DisplayTexts[AViewInfo.Item.Index];
      aDrawFlag := ( cxSingleLine	or cxShowEndEllipsis or cxAlignVCenter or cxAlignHCenter );
      ACanvas.DrawTexT( aText, aRect, aDrawFlag );
      ADone := True;
    end else
    begin
      if ( AViewInfo.State = gcsNone ) and not AViewInfo.GridRecord.Focused then
        ACanvas.Font.Color := clBlue;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.

