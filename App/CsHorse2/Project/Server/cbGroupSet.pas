unit cbGroupSet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, StdCtrls, DB,
{ App: }
  cbStyleModule, cbLanguage, cbSrvClass, cbHrHelper, cbOraModule,
{ ODAC: }
  Ora, MemDS, VirtualTable,
{ Developer Express: }
  dxSkinsCore, dxSkinsDefaultPainters, cxLookAndFeelPainters,
  cxContainer, cxCustomData, cxStyles, cxGraphics,
  cxControls, cxInplaceContainer,
  cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxImageComboBox,
  cxButtons, cxSplitter, cxTL, cxTLData, cxDBTL, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinSilver, dxSkinStardust,
  dxSkinValentine, dxSkinXmas2008Blue, dxSkinSummer2008;

type
  TfmGroupSet = class(TForm)
    Panel1: TPanel;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    Bevel1: TBevel;
    HeaderPanel: TPanel;
    HeaderImage: TImage;
    txtCompInfo: TcxImageComboBox;
    CH009Tree: TcxDBTreeList;
    CH009TreeCol1: TcxDBTreeListColumn;
    CH009TreeCol2: TcxDBTreeListColumn;
    CH009TreeCol3: TcxDBTreeListColumn;
    CH009TreeCol4: TcxDBTreeListColumn;
    CH009TreeCol5: TcxDBTreeListColumn;
    cxSplitter1: TcxSplitter;
    CD071Tree: TcxDBTreeList;
    CD071TreeCol1: TcxDBTreeListColumn;
    CD071TreeCol2: TcxDBTreeListColumn;
    CD071TreeCol3: TcxDBTreeListColumn;
    CD071TreeCol4: TcxDBTreeListColumn;
    CD071TreeCol5: TcxDBTreeListColumn;
    dsCH009: TDataSource;
    dsCD071: TDataSource;
    CH009Buffer: TVirtualTable;
    CD071Buffer: TVirtualTable;
    LoadTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LoadTimerTimer(Sender: TObject);
    procedure txtCompInfoPropertiesChange(Sender: TObject);
    procedure CH009TreeDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure CH009TreeDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure CD071TreeDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure CD071TreeDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure CH009TreeCustomDrawCell(Sender: TObject; ACanvas: TcxCanvas;
      AViewInfo: TcxTreeListEditCellViewInfo; var ADone: Boolean);
    procedure CD071TreeCustomDrawCell(Sender: TObject; ACanvas: TcxCanvas;
      AViewInfo: TcxTreeListEditCellViewInfo; var ADone: Boolean);
    procedure btnConfirmClick(Sender: TObject);
  private
    { Private declarations }
    FOraModule: TOraModule;
    procedure InitLanguage;
    procedure InitControl;
    procedure BuildSoCompInfo;
    function BeginSession(AppSo: TAppSo): Boolean;
    procedure EndSession;
    procedure FilterBuffer(ABuffer: TVirtualTable; const ACompCode: String);
    procedure LoadBuffer;
    procedure LoadCH009(const AppSo: TAppSo);
    procedure LoadCD071(const AppSo: TAppSo);
    procedure CompareCH009WithCD071;
    procedure CD071ToCH009(const ASndTree, ASrcTree:TcxDBTreeList;
      const ASndNode: TcxTreeListDataNode);
    procedure CH009ToCD071(const ASndTree, ASrcTree:TcxDBTreeList);
    procedure ProcessChildNode(const ABuffer: TDataSet; ANode: TcxTreeListNode;
      const AIsChild: Boolean = False);
    function SaveAll: Boolean;
    function SaveToCH009(const AppSo: TAppSo): Boolean;
  public
    { Public declarations }
  end;

var
  fmGroupSet: TfmGroupSet;

implementation

{$R *.dfm}

uses cbUtilis, cbConfigModule;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.FormCreate(Sender: TObject);
begin
  FOraModule := TOraModule.Create( nil );
  InitLanguage;
  InitControl;
  BuildSoCompInfo;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.FormShow(Sender: TObject);
begin
  LoadTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.FormDestroy(Sender: TObject);
begin
  FOraModule.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.LoadTimerTimer(Sender: TObject);
begin
  LoadTimer.Enabled := False;
  LoadBuffer;
  if txtCompInfo.Properties.Items.Count > 0 then
    txtCompInfo.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.InitLanguage;
var
  aIndex: Integer;
begin
  for aIndex := 0 to CH009Tree.ColumnCount - 1 do
  begin
    CH009Tree.Columns[aIndex].Caption.Text :=
      LanguageManager.Get( CH009Tree.Columns[aIndex].Name );
  end;
  {}
  for aIndex := 0 to CD071Tree.ColumnCount - 1 do
  begin
    CD071Tree.Columns[aIndex].Caption.Text :=
      LanguageManager.Get( CD071Tree.Columns[aIndex].Name );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.InitControl;
begin
  StyleModule.ImgList32.GetIcon( StyleModule.GroupSetHeaderImage,
    HeaderImage.Picture.Icon );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.BuildSoCompInfo;
var
  aIndex: Integer;
  aItem: TcxImageComboBoxItem;
begin
  txtCompInfo.Properties.OnChange := nil;
  try
    txtCompInfo.Properties.Items.Clear;
    for aIndex := 0 to ConfigModule.SoList.Count - 1 do
    begin
      if ( ConfigModule.SoList[aIndex].Selected ) then
      begin
        aItem := txtCompInfo.Properties.Items.Add;
        aItem.Value := ConfigModule.SoList[aIndex].CompCode;
        aItem.Description := ConfigModule.SoList[aIndex].CompName;
        aItem.ImageIndex := 2;
      end;
    end;
  finally
    txtCompInfo.Properties.OnChange := txtCompInfoPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmGroupSet.BeginSession(AppSo: TAppSo): Boolean;
begin
  if ( FOraModule.OraSession.Connected ) then FOraModule.OraSession.Close;
  FOraModule.OraSession.ConnectString := Format( '%s/%s@%s',
    [AppSo.DbUserId, AppSo.DbUserPass, AppSo.DbAliase] );
  try
    FOraModule.OraSession.Open;
  except
    on E: Exception do
      ErrorMsg( LanguageManager.GetFmt( 'SOraSessionOpenError', [E.Message] ) );
  end;
  Result := FOraModule.OraSession.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.EndSession;
begin
  if ( FOraModule.OraSession.Connected ) then FOraModule.OraSession.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.FilterBuffer(ABuffer: TVirtualTable; const ACompCode: String);
begin
  if Assigned( ABuffer ) then
  begin
    ABuffer.Filtered := False;
    ABuffer.Filter := Format( 'CompCode=''%s''', [ACompCode] );
    ABuffer.Filtered := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.LoadBuffer;
var
  aIndex: Integer;
  aAppSo: TAppSo;
begin
  CH009Buffer.Clear;
  CD071Buffer.Clear;
  TBufferHelper.CreateFieldDefs( biCH009, CH009Buffer );
  TBufferHelper.CreateFieldDefs( biCD071, CD071Buffer );
  Screen.Cursor := crHourGlass;
  try
    for aIndex := 0 to txtCompInfo.Properties.Items.Count - 1 do
    begin
      aAppSo := ConfigModule.SoList[ConfigModule.SoList.IndexOf(
        txtCompInfo.Properties.Items[aIndex].Value )];
      if Assigned( aAppSo ) then
      begin
        if BeginSession( aAppSo ) then
        begin
          LoadCH009( aAppSo );
          LoadCD071( aAppSo );
        end;
      end;
      EndSession;
    end;
    CompareCH009WithCD071;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.LoadCH009(const AppSo: TAppSo);
var
  aSql: String;
  aReader: TOraQuery;
begin
  TBufferHelper.CreateSqlText( biCH009, aSql );
  CH009Buffer.DisableControls;
  try
    aReader := FOraModule.OraQuery;
    aReader.Close;
    aReader.SQL.Text := Format( aSql, [AppSo.CompCode, AppSo.CompName] );
    aReader.Open;
    aReader.First;
    while not aReader.Eof do
    begin
      CH009Buffer.Append;
      CH009Buffer.FieldByName( 'CompCode' ).AsString := aReader.FieldByName( 'CompCode' ).AsString;
      CH009Buffer.FieldByName( 'CompName' ).AsString := aReader.FieldByName( 'CompName' ).AsString;
      CH009Buffer.FieldByName( 'GroupId' ).AsString := aReader.FieldByName( 'GroupId' ).AsString;
      CH009Buffer.FieldByName( 'GroupName' ).AsString := aReader.FieldByName( 'GroupName' ).AsString;
      CH009Buffer.FieldByName( 'RGroupId' ).AsString := aReader.FieldByName( 'RGroupId' ).AsString;
      CH009Buffer.FieldByName( 'StopFlag' ).AsString := aReader.FieldByName( 'StopFlag' ).AsString;
      CH009Buffer.FieldByName( 'ImageIndex' ).AsInteger := StyleModule.GroupImageIndex;
      CH009Buffer.Post;
      aReader.Next;
    end;
    CH009Buffer.First;
    CH009Buffer.IndexFieldNames := 'CompCode;GroupId';
  finally
    CH009Buffer.EnableControls;
  end;
  aReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.LoadCD071(const AppSo: TAppSo);
var
  aSql: String;
  aReader: TOraQuery;
begin
  TBufferHelper.CreateSqlText( biCD071, aSql );
  CD071Buffer.DisableControls;
  try
    aReader := FOraModule.OraQuery;
    aReader.Close;
    aReader.SQL.Text := Format( aSql, [AppSo.CompCode, AppSo.CompName] );
    aReader.Open;
    aReader.First;
    while not aReader.Eof do
    begin
      CD071Buffer.Append;
      CD071Buffer.FieldByName( 'CompCode' ).AsString := aReader.FieldByName( 'CompCode' ).AsString;
      CD071Buffer.FieldByName( 'CompName' ).AsString := aReader.FieldByName( 'CompName' ).AsString;
      CD071Buffer.FieldByName( 'CodeNo' ).AsString := aReader.FieldByName( 'CodeNo' ).AsString;
      CD071Buffer.FieldByName( 'Description' ).AsString := aReader.FieldByName( 'Description' ).AsString;
      CD071Buffer.FieldByName( 'StopFlag' ).AsString := aReader.FieldByName( 'StopFlag' ).AsString;
      CD071Buffer.FieldByName( 'ImageIndex' ).AsInteger := StyleModule.GroupImageIndex;
      CD071Buffer.Post;
      aReader.Next;
    end;
    CD071Buffer.First;
    CD071Buffer.IndexFieldNames := 'CompCode;CodeNo';
  finally
    CD071Buffer.EnableControls;
  end;
  aReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CompareCH009WithCD071;
begin
  CD071Buffer.DisableControls;
  CH009Buffer.DisableControls;
  try
    CD071Buffer.First;
    while not CD071Buffer.Eof do
    begin
      if ( CH009Buffer.Locate( 'CompCode;GroupId', VarArrayOf( [
        CD071Buffer.FieldByName( 'CompCode' ).AsInteger,
        CD071Buffer.FieldByName( 'CodeNo' ).AsString] ), [] ) ) then
       CD071Buffer.Delete
     else
      CD071Buffer.Next;
    end;
  finally
    CH009Buffer.EnableControls;
    CD071Buffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.txtCompInfoPropertiesChange(Sender: TObject);
var
  aAppSo: TAppSo;
  aCompCode: String;
begin
  Screen.Cursor := crHourGlass;
  try
    aAppSo := ConfigModule.SoList[ConfigModule.SoList.IndexOf(
      txtCompInfo.Properties.Items[txtCompInfo.ItemIndex].Value )];
    aCompCode := EmptyStr;
    if Assigned( aAppSo ) then aCompCode := aAppSo.CompCode;
      FilterBuffer( CH009Buffer, aCompCode );
    FilterBuffer( CD071Buffer, aCompCode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CD071ToCH009(const ASndTree, ASrcTree:TcxDBTreeList; const
  ASndNode: TcxTreeListDataNode);
var
  aIndex: Integer;
  aRGroupId: String;
  aBuffer: TDataSet;
  aNode: TcxTreeListNode;
begin
  aRGroupId := EmptyStr;
  aBuffer := ASndTree.DataController.DataSet;
  if Assigned( ASndNode ) then
    aRGroupId := VarToStrDef( ASndNode.KeyValue, EmptyStr );
  aBuffer.DisableControls;
  try
    for aIndex := 0 to ASrcTree.SelectionCount - 1 do
    begin
      aNode := ASrcTree.Selections[aIndex];
      aBuffer.Append;
      aBuffer.FieldByName( 'CompCode' ).AsString := VarToStrDef( aNode.Values[CD071TreeCol1.ItemIndex], EmptyStr );
      aBuffer.FieldByName( 'CompName' ).AsString := VarToStrDef( aNode.Values[CD071TreeCol2.ItemIndex], EmptyStr );
      aBuffer.FieldByName( 'GroupId' ).AsString := VarToStrDef( aNode.Values[CD071TreeCol3.ItemIndex], EmptyStr );
      aBuffer.FieldByName( 'GroupName' ).AsString := VarToStrDef( aNode.Values[CD071TreeCol4.ItemIndex], EmptyStr );
      aBuffer.FieldByName( 'RGroupId' ).AsString := aRGroupId;
      aBuffer.FieldByName( 'StopFlag' ).AsString := VarToStrDef( aNode.Values[CD071TreeCol5.ItemIndex], EmptyStr );
      aBuffer.FieldByName( 'ImageIndex' ).AsInteger:= StyleModule.GroupImageIndex;
      aBuffer.Post;
    end;
    ASrcTree.DeleteSelection;
  finally
    aBuffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.ProcessChildNode(const ABuffer: TDataSet; ANode: TcxTreeListNode;
  const AIsChild: Boolean = False);
begin
  while Assigned( ANode ) do
  begin
    aBuffer.Append;
    aBuffer.FieldByName( 'CompCode' ).AsString := VarToStrDef( ANode.Values[CH009TreeCol1.ItemIndex], EmptyStr );
    aBuffer.FieldByName( 'CompName' ).AsString := VarToStrDef( ANode.Values[CH009TreeCol2.ItemIndex], EmptyStr );
    aBuffer.FieldByName( 'CodeNo' ).AsString := VarToStrDef( ANode.Values[CH009TreeCol3.ItemIndex], EmptyStr );
    aBuffer.FieldByName( 'Description' ).AsString := VarToStrDef( ANode.Values[CH009TreeCol4.ItemIndex], EmptyStr );
    aBuffer.FieldByName( 'StopFlag' ).AsString := VarToStrDef( ANode.Values[CH009TreeCol5.ItemIndex], EmptyStr );
    aBuffer.FieldByName( 'ImageIndex' ).AsInteger:= StyleModule.GroupImageIndex;
    aBuffer.Post;
    if ANode.HasChildren then
      ProcessChildNode( ABuffer, ANode.getFirstChild, True );
    if ( AIsChild ) then
      ANode := ANode.getNextSibling
    else
      ANode := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CH009ToCD071(const ASndTree, ASrcTree: TcxDBTreeList);
var
  aIndex: Integer;
  aBuffer: TDataSet;
  aNode: TcxTreeListNode;
begin
  aBuffer := ASndTree.DataController.DataSet;
  aBuffer.DisableControls;
  try
    for aIndex := 0 to ASrcTree.SelectionCount - 1 do
    begin
      aNode := ASrcTree.Selections[aIndex];
      ProcessChildNode( ABuffer, aNode );
    end;
    ASrcTree.DeleteSelection;
  finally
    aBuffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CH009TreeDragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  if ( Sender = Source ) then Exit;
  CD071ToCH009( TcxDBTreeList( Sender ), TcxDBTreeList( Source ),
    TcxTreeListDataNode( TcxDBTreeList( Sender ).HitTest.HitNode ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CH009TreeDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  TcxDBTreeList( Sender ).HitTest.ReCalculate( Point( X,Y ) );
  Accept := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CD071TreeDragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  CH009ToCD071( TcxDBTreeList( Sender ), TcxDBTreeList( Source ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CD071TreeDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := ( Sender <> Source );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CH009TreeCustomDrawCell(Sender: TObject;
  ACanvas: TcxCanvas; AViewInfo: TcxTreeListEditCellViewInfo;
  var ADone: Boolean);
var
  aIsStop: Boolean;
begin
  aIsStop := SameText(
    VarToStrDef( AViewInfo.Node.Values[CH009TreeCol5.ItemIndex], '0' ), '1' );
  if ( aIsStop ) then
  begin
    ACanvas.Font.Color := clGrayText;
    ACanvas.Font.Style := ( ACanvas.Font.Style + [fsStrikeOut] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.CD071TreeCustomDrawCell(Sender: TObject;
  ACanvas: TcxCanvas; AViewInfo: TcxTreeListEditCellViewInfo;
  var ADone: Boolean);
var
  aIsStop: Boolean;
begin
  aIsStop := SameText(
    VarToStrDef( AViewInfo.Node.Values[CH009TreeCol5.ItemIndex], '0' ), '1' );
  if ( aIsStop ) then
  begin
    ACanvas.Font.Color := clGrayText;
    ACanvas.Font.Style := ( ACanvas.Font.Style + [fsStrikeOut] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmGroupSet.SaveToCH009(const AppSo: TAppSo): Boolean;
var
  aInstSql, aDelSql: String;
begin
  Result := True;
  aInstSql :=
    ' INSERT INTO CH009 ( COMPCODE, GROUPID, GROUPNAME, RGROUPID ) ' +
    '  VALUES ( ''%s'', ''%s'', ''%s'', ''%s'' ) ';
  aDelSql :=
    ' DELETE FROM CH009 WHERE COMPCODE = ''%s'' ';
  CH009Buffer.DisableControls;
  try
    FOraModule.OraSQL.SQL.Text := Format( aDelSql, [AppSo.CompCode] );
    FOraModule.OraSQL.Execute;
    CH009Buffer.First;
    try
      while not CH009Buffer.Eof do
      begin
        FOraModule.OraSQL.SQL.Text := Format( aInstSql, [
          CH009Buffer.FieldByName( 'CompCode' ).AsString,
          CH009Buffer.FieldByName( 'GroupId' ).AsString,
          CH009Buffer.FieldByName( 'GroupName' ).AsString,
          CH009Buffer.FieldByName( 'RGroupId' ).AsString] );
        FOraModule.OraSQL.Execute;
        CH009Buffer.Next;
      end;
      FOraModule.OraSession.Commit;
    except
      on E: Exception do
      begin
        FOraModule.OraSession.Rollback;
        ErrorMsg( E.Message );
        Result := False;
      end;
    end;
  finally
    CH009Buffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmGroupSet.SaveAll: Boolean;
var
  aIndex: Integer;
  aAppSo: TAppSo;
begin
  Result := False;
  for aIndex := 0 to txtCompInfo.Properties.Items.Count - 1 do
  begin
    aAppSo := ConfigModule.SoList[ConfigModule.SoList.IndexOf(
      txtCompInfo.Properties.Items[aIndex].Value )];
    if Assigned( aAppSo ) then
    begin
      if BeginSession( aAppSo ) then
      begin
        txtCompInfo.ItemIndex := aIndex;
        Result := ( SaveToCH009( aAppSo ) and Result );
      end;
    end;
    EndSession;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmGroupSet.btnConfirmClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;
  try
    if SaveAll then
      Self.ModalResult := mrOk;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
