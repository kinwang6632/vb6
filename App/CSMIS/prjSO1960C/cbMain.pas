unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ActnList, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid, ADODB,
  DBClient, Provider, cxGridStrs;

type
  TfmSO1960C = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    cmbSearchType: TComboBox;
    txtMac: TEdit;
    ActionList1: TActionList;
    actSearch: TAction;
    actClear: TAction;
    btnSearch: TButton;
    btnClear: TButton;
    Button7: TButton;
    gvMac: TcxGridDBTableView;
    glMac: TcxGridLevel;
    cxGridMac: TcxGrid;
    dsDataTable: TDataSource;
    ADOConnection1: TADOConnection;
    DataReader: TADOQuery;
    DataTable: TClientDataSet;
    DataGeter: TADOQuery;
    gvMacHFC: TcxGridDBColumn;
    gvMacCOMPCODE_1: TcxGridDBColumn;
    gvMacMODELNO: TcxGridDBColumn;
    gvMacMTAMAC: TcxGridDBColumn;
    gvMacCMISSUE: TcxGridDBColumn;
    gvMacMODEID: TcxGridDBColumn;
    gvMacMODELCODE: TcxGridDBColumn;
    gvMacUSEFLAG: TcxGridDBColumn;
    gvMacVENDORCODE: TcxGridDBColumn;
    gvMacVENDORNAME: TcxGridDBColumn;
    gvMacSPEC: TcxGridDBColumn;
    gvMacUPDEN: TcxGridDBColumn;
    gvMacUPDTIME: TcxGridDBColumn;
    gvMacBATCHNO: TcxGridDBColumn;
    gvMacMATERIALNO: TcxGridDBColumn;
    gvMacDESCRIPTION: TcxGridDBColumn;
    DataSO507: TClientDataSet;
    glSo: TcxGridLevel;
    gvSo: TcxGridDBTableView;
    SoData: TClientDataSet;
    dsSoData: TDataSource;
    gvSoKeyId: TcxGridDBColumn;
    gvSoCompCode: TcxGridDBColumn;
    gvSoCompName: TcxGridDBColumn;
    gvSoCustId: TcxGridDBColumn;
    gvSoInstDate: TcxGridDBColumn;
    gvSoFaciStatus: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure actSearchExecute(Sender: TObject);
    procedure actClearExecute(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Button7Click(Sender: TObject);
    procedure gvSoCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
  private
    { Private declarations }
    FComOwner: String;
    FComLink: String;
    function ConnectToCATVN: Boolean;
    function BuildSo004Sql(const aOwner, aLink, aMac: String): String;  
    procedure GetComLinkData;
    function GetCompName(aCompCode: String): String;
    procedure DoGetComMacData(aMac: String);
    procedure PrepareDataTable;
    procedure UnPrepareDataTable;
    {}
    procedure PrepareSoData;
    procedure UnPrepareSoData;
    {}
    procedure GetSoMacData;
    procedure PrepareSO507;
    procedure UnPrepareSO507;
    procedure GetSoOwnerAndLink(aCompCode: String; var aOwner, aLink: String);
    procedure ChangeDevexComponentLang;
  public
    { Public declarations }
    procedure DoSearch;
  end;

var
  fmSO1960C: TfmSO1960C;

implementation

{$R *.dfm}

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

function PadDblinkSymbol(aDblink: String): String;
begin
  Result := aDblink;
  if Trim( aDblink ) <> EmptyStr then
    if ( aDblink[1] <> '@' ) then
     Result := '@' + aDblink;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.FormCreate(Sender: TObject);
begin
  if not ( ParamCount > 0 ) then
  begin
    WarningMsg( '傳入之程式啟動參數有誤。' );
    Application.Terminate;
    Exit;
  end;
  if not ConnectToCATVN then
  begin
    Application.Terminate;
    Exit;
  end;
  UnPrepareSO507;
  PrepareSO507;
  GetComLinkData;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.FormShow(Sender: TObject);
begin
  ChangeDevexComponentLang;
end;

{ ---------------------------------------------------------------------------- }

function TfmSO1960C.ConnectToCATVN: Boolean;
begin
  Result := False;
  ADOConnection1.Close;
  ADOConnection1.ConnectionString :=
   'Provider=MSDAORA.1;Password=COM;User ID=COM;Data Source=CATVN;' +
   'Persist Security Info=True';
  try
    ADOConnection1.Open;
    Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '資料庫連接失敗, 訊息:%s。', [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.actSearchExecute(Sender: TObject);
begin
  if Trim( txtMac.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入MAC序號。' );
    if txtMac.CanFocus then txtMac.SetFocus;
    Exit;
  end;
  if ( cmbSearchType.ItemIndex = 0 ) and
     ( Length( txtMac.Text ) in [1..9] ) then
  begin
    WarningMsg( '使用模糊查詢,最少必須輸入10個字元。' );
    if txtMac.CanFocus then txtMac.SetFocus;
    Exit;
  end;
//  if ( cmbSearchType.ItemIndex = 1 ) and
//     ( Length( txtMac.Text ) < 12 ) then
//  begin
//    WarningMsg( '使用完全查詢,請輸入至少12個字元。' );
//    if txtMac.CanFocus then txtMac.SetFocus;
//    Exit;
//  end;
  DoSearch;   
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.DoSearch;
begin
  Screen.Cursor := crSQLWait;
  try
    UnPrepareSoData;
    UnPrepareDataTable;
    {}
    PrepareDataTable;
    PrepareSoData;
    {}
    DataTable.DisableControls;
    SoData.DisableControls;
    try
      DoGetComMacData( txtMac.Text );
      if DataTable.IsEmpty then
      begin
        WarningMsg( '查無此MAC序號,請確認您輸入的序號是否正確。' );
        if txtMac.CanFocus then txtMac.SetFocus;
        Exit;
      end;
      GetSoMacData;
      if cxGridMac.CanFocus then cxGridMac.SetFocus;
    finally
      SoData.EnableControls;
      DataTable.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.GetComLinkData;
begin
  DataGeter.Close;
  DataGeter.SQL.Text :=
    ' select * from so507 ';
  DataGeter.Open;
  DataGeter.First;
  while not DataGeter.Eof do
  begin
    DataSO507.Append;
    DataSO507.FieldByName( 'codeno' ).AsString := DataGeter.FieldByName( 'codeno' ).AsString;
    DataSO507.FieldByName( 'description' ).AsString := DataGeter.FieldByName( 'description' ).AsString;
    DataSO507.FieldByName( 'compname' ).AsString := DataGeter.FieldByName( 'compname' ).AsString;
    DataSO507.FieldByName( 'link' ).AsString := DataGeter.FieldByName( 'link' ).AsString;
    DataSO507.Post;
    DataGeter.Next;
  end;
  DataGeter.Close;
  if not DataSO507.Locate( 'codeno', '0', [] ) then
  begin
    ErrorMsg( '對應不到客服物料總倉資料區,無法執行本功能。' );
    Application.Terminate;
    Exit;
  end;
  FComLink := DataSO507.FieldByName( 'link' ).AsString;
  FComOwner := DataSO507.FieldByName( 'description' ).AsString;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.DoGetComMacData(aMac: String);

   { --------------------------------------- }

   function GetUseFlagText: String;
   begin
     Result := EmptyStr;
     if DataReader.FieldByName( 'useflag' ).AsString = '0' then
       Result := '未領料'
     else if DataReader.FieldByName( 'useflag' ).AsString = '1' then
       Result := '已領料/使用中'
     else if DataReader.FieldByName( 'useflag' ).AsString = '2' then
       Result := '註銷';
   end;

   { --------------------------------------- }

var
  aSql, aCMCompName: String;
begin
  if ( cmbSearchType.ItemIndex = 0 ) then
  begin
    aSql := Format(
      ' select * from %0:s.so306%1:s where hfcmac like ''%2:s'' ',
      [FComOwner, PadDblinkSymbol( FComLink ), aMac + '%'] );
  end else
  begin
    aSql := Format(
      ' select * from %0:s.so306%1:s where hfcmac = ''%2:s'' ',
      [FComOwner, PadDblinkSymbol( FComLink ), aMac] );
  end;    
  DataReader.Close;
  DataReader.SQL.Text := aSql;
  DataReader.Open;
  DataReader.First;
  while not DataReader.Eof do
  begin
    aCMCompName := GetCompName( DataReader.FieldByName( 'compcode' ).AsString );
    DataTable.Append;
    DataTable.FieldByName( 'cmcompcode' ).AsString := DataReader.FieldByName( 'compcode' ).AsString;
    DataTable.FieldByName( 'cmcompname' ).AsString := aCMCompName;
    DataTable.FieldByName( 'hfcmac' ).AsString := DataReader.FieldByName( 'hfcmac' ).AsString;
    DataTable.FieldByName( 'modelno' ).AsString := DataReader.FieldByName( 'modelno' ).AsString;
    DataTable.FieldByName( 'mtamac' ).AsString := DataReader.FieldByName( 'mtamac' ).AsString;
    DataTable.FieldByName( 'cmissue' ).AsString := DataReader.FieldByName( 'cmissue' ).AsString;
    DataTable.FieldByName( 'modeid' ).AsString :=  DataReader.FieldByName( 'modeid' ).AsString;
    DataTable.FieldByName( 'modelcode' ).AsString := DataReader.FieldByName( 'modelcode' ).AsString;
    DataTable.FieldByName( 'useflag' ).AsString := GetUseFlagText;
    DataTable.FieldByName( 'vendorcode' ).AsString := DataReader.FieldByName( 'vendorcode' ).AsString;
    DataTable.FieldByName( 'vendorname' ).AsString := DataReader.FieldByName( 'vendorname' ).AsString;
    DataTable.FieldByName( 'spec' ).AsString := DataReader.FieldByName( 'spec' ).AsString;
    DataTable.FieldByName( 'upden' ).AsString := DataReader.FieldByName( 'upden' ).AsString;
    DataTable.FieldByName( 'updtime' ).AsString := DataReader.FieldByName( 'updtime' ).AsString;
    DataTable.FieldByName( 'batchno' ).AsString := DataReader.FieldByName( 'batchno' ).AsString;
    DataTable.FieldByName( 'materialno' ).AsString := DataReader.FieldByName( 'materialno' ).AsString;
    DataTable.FieldByName( 'description' ).AsString :=  DataReader.FieldByName( 'description' ).AsString;
    DataTable.Post;
    DataReader.Next;
  end;
  DataReader.Close;
  DataTable.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.PrepareDataTable;
begin
  if ( DataTable.FieldDefs.Count <= 0 ) then
  begin
    DataTable.FieldDefs.Add( 'cmcompcode', ftString, 20 );
    DataTable.FieldDefs.Add( 'cmcompname', ftString, 20 );
    DataTable.FieldDefs.Add( 'hfcmac', ftString, 20 );
    DataTable.FieldDefs.Add( 'modelno', ftString, 30 );
    DataTable.FieldDefs.Add( 'mtamac', ftString, 20 );
    DataTable.FieldDefs.Add( 'cmissue', ftDateTime );
    DataTable.FieldDefs.Add( 'modeid', ftString, 1 );
    DataTable.FieldDefs.Add( 'modelcode', ftString, 3 );
    DataTable.FieldDefs.Add( 'useflag', ftString, 20 );
    DataTable.FieldDefs.Add( 'vendorcode', ftString, 5 );
    DataTable.FieldDefs.Add( 'vendorname', ftString, 20 );
    DataTable.FieldDefs.Add( 'spec', ftString, 50 );
    DataTable.FieldDefs.Add( 'upden', ftString, 20 );
    DataTable.FieldDefs.Add( 'updtime', ftString, 19 );
    DataTable.FieldDefs.Add( 'batchno', ftString, 15 );
    DataTable.FieldDefs.Add( 'materialno', ftString, 15 );
    DataTable.FieldDefs.Add( 'description', ftString, 30 );
  end;
  DataTable.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.PrepareSoData;
begin
  if ( SoData.FieldDefs.Count <= 0 ) then
  begin
    SoData.FieldDefs.Add( 'hfcmac', ftString, 20 );
    SoData.FieldDefs.Add( 'compcode', ftString, 20 );
    SoData.FieldDefs.Add( 'compname', ftString, 50 );
    SoData.FieldDefs.Add( 'custid', ftString, 20 );
    SoData.FieldDefs.Add( 'instdate', ftDateTime );
    SoData.FieldDefs.Add( 'facistatus', ftString, 20 );
  end;
  SoData.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.UnPrepareDataTable;
begin
  if not VarIsNull( DataTable.Data ) then
    DataTable.EmptyDataSet;
  DataTable.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.UnPrepareSoData;
begin
  if not VarIsNull( SoData.Data ) then
    SoData.EmptyDataSet;
  SoData.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.GetSoMacData;

  { -------------------------------------------- }

  function GetFaciStatus: String;
  begin
    Result := EmptyStr;
    if ( ( DataReader.FieldByName( 'sno' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'instdate' ).AsString = EmptyStr ) and
         ( DataReader.FieldByName( 'instdate' ).AsString = EmptyStr ) ) then
    begin
      Result := '安裝中';
    end else
    if ( ( DataReader.FieldByName( 'instdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'prdate' ).AsString = EmptyStr ) and
         ( DataReader.FieldByName( 'prsno' ).AsString = EmptyStr ) ) then
    begin
      Result := '使用中';
    end else
    if ( ( DataReader.FieldByName( 'instdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'prdate' ).AsString = EmptyStr ) and
         ( DataReader.FieldByName( 'prsno' ).AsString <> EmptyStr ) ) then
    begin
      Result := '拆除中';
    end else
    if ( ( DataReader.FieldByName( 'instdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'prdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'getdate' ).AsString = EmptyStr ) ) then
    begin
      Result := '已拆除待取回';
    end;
    if ( ( DataReader.FieldByName( 'instdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'prdate' ).AsString <> EmptyStr ) and
         ( DataReader.FieldByName( 'getdate' ).AsString <> EmptyStr ) ) then
    begin
      Result := '已拆除並取回';
    end;
  end;

  { -------------------------------------------- }

var
  aMac, aSql: String;
begin
  DataTable.First;
  while not DataTable.Eof do
  begin
    aMac := DataTable.FieldByName( 'hfcmac' ).AsString;
    DataSO507.First;
    // 該筆 MAC 開始掃所有 SO 裏的 SO004 資料
    while not DataSO507.Eof do
    begin
      if ( DataSO507.FieldByName( 'codeno' ).AsInteger in [0,13,14] ) then
      begin
        DataSO507.Next;
        Continue;
      end;
      aSql := BuildSo004Sql(
        DataSO507.FieldByName( 'description' ).AsString,
        DataSO507.FieldByName( 'link' ).AsString, aMac );
      DataReader.Close;
      DataReader.SQL.Text := aSql;
      DataReader.Open;
      DataReader.First;
      while not DataReader.Eof do
      begin
        SoData.Append;
        SoData.FieldByName( 'hfcmac' ).AsString := aMac;
        SoData.FieldByName( 'compcode' ).AsString := DataReader.FieldByName( 'compcode' ).AsString;
        SoData.FieldByName( 'compname' ).AsString := GetCompName( DataReader.FieldByName( 'compcode' ).AsString );
        SoData.FieldByName( 'custid' ).AsString := DataReader.FieldByName( 'custid' ).AsString;
        SoData.FieldByName( 'instdate' ).AsString := DataReader.FieldByName( 'instdate' ).AsString;
        SoData.FieldByName( 'facistatus' ).AsString := GetFaciStatus;
        SoData.Post;
        DataReader.Next;
      end;
      DataReader.Close;
      DataSO507.Next;
    end;
    DataTable.Next;
  end;
  DataTable.First;
end;

{ ---------------------------------------------------------------------------- }

function TfmSO1960C.BuildSo004Sql(const aOwner, aLink, aMac: String): String;
begin
  Result := Format(
    ' select so004.compcode, so004.custid, so004.facisno, so004.instdate,  ' +
    '        so004.sno, so004.prdate, so004.prsno, so004.getdate           ' +
    '   from %0:s.so004%1:s,                     ' +
    '        %0:s.so001%1:s,                     ' +
    '        %0:s.cd022%1:s                      ' +
    '  where so001.custid = so004.custid         ' +
    '    and so004.facicode = cd022.codeno       ' +
    '    and cd022.refno in ( 2, 5 )             ' +
    '    and so004.facisno like ''%2:s''         ',
    [aOwner, PadDblinkSymbol( aLink ), aMac + '%'] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.actClearExecute(Sender: TObject);
begin
  txtMac.Clear;
  cmbSearchType.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.PrepareSO507;
begin
  if ( DataSO507.FieldDefs.Count <= 0 ) then
  begin
    DataSO507.FieldDefs.Add( 'codeno', ftString, 3 );
    DataSO507.FieldDefs.Add( 'description', ftString, 30 );
    DataSO507.FieldDefs.Add( 'compname', ftString, 30 );
    DataSO507.FieldDefs.Add( 'link', ftString, 30 );
  end;
  DataSO507.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.UnPrepareSO507;
begin
  if not VarIsNull( DataSO507.Data ) then
    DataSO507.EmptyDataSet;
  DataSO507.Data := Null;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  UnPrepareSO507;
  UnPrepareSoData;
  UnPrepareDataTable;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.GetSoOwnerAndLink(aCompCode: String; var aOwner,
  aLink: String);
begin
  aOwner := EmptyStr;
  aLink := EmptyStr;
  if DataSO507.Locate( 'codeno', aCompCode, [] ) then
  begin
    aOwner := DataSO507.FieldByName( 'description' ).AsString;
    aLink :=  DataSO507.FieldByName( 'Link' ).AsString;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.Button7Click(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.ChangeDevexComponentLang;
begin
  cxSetResourceString( @scxGridNoDataInfoText, '' );
end;

{ ---------------------------------------------------------------------------- }

function TfmSO1960C.GetCompName(aCompCode: String): String;
begin
  Result := EmptyStr;
  if DataSO507.Locate( 'codeno', aCompCode, [] ) then
    Result := DataSO507.FieldByName( 'compname' ).AsString;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSO1960C.gvSoCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aColumn: TcxGridDBColumn;
  aText: String;
begin
  if AViewInfo.Selected then
  begin
    ACanvas.Font.Color := clHighlightText;
  end else
  begin
    aColumn := gvSo.GetColumnByFieldName( 'facistatus' );
    aText := VarToStrDef( AViewInfo.GridRecord.Values[aColumn.Index], EmptyStr );
    if ( aText = '使用中' ) then
      ACanvas.Font.Color := clBlue
    else
      ACanvas.Font.Color := clRed;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
