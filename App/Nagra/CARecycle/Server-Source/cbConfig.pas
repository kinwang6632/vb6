unit cbConfig;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ImgList, DB, ADODB, cbClass, cbAppClass,
  cxLookAndFeelPainters, cxPC, cxButtons, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxCheckBox, cxSpinEdit, cxTimeEdit, cxRadioGroup,
  cxGraphics, cxCustomData, cxStyles, cxTL, cxImageComboBox,
  cxInplaceContainer, cxDropDownEdit, Menus, OracleCI;



type
  TfmConfig = class(TForm)
    ConfigPage: TcxPageControl;
    Panel1: TPanel;
    Bevel1: TBevel;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    cxSoDb: TcxTabSheet;
    SoTree: TcxTreeList;
    SoTreeColumn1: TcxTreeListColumn;
    SoTreeColumn2: TcxTreeListColumn;
    SoTreeColumn3: TcxTreeListColumn;
    SoTreeColumn5: TcxTreeListColumn;
    SoTreeColumn6: TcxTreeListColumn;
    SoTreeColumn7: TcxTreeListColumn;
    SoTreeColumn9: TcxTreeListColumn;
    btnDeleteSo: TcxButton;
    btnAddSo: TcxButton;
    btnTest: TcxButton;
    lblTest: TLabel;
    HeaderPanel: TPanel;
    HeaderImage: TImage;
    SoTreeColumn10: TcxTreeListColumn;
    SoTreeColumn11: TcxTreeListColumn;
    SoTreeColumn12: TcxTreeListColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ConfigPageChange(Sender: TObject);
    procedure SoTreeFocusedNodeChanged(Sender: TObject; APrevFocusedNode,
      AFocusedNode: TcxTreeListNode);
    procedure SoTreeColumn1PropertiesEditValueChanged(Sender: TObject);
    procedure btnAddSoClick(Sender: TObject);
    procedure btnDeleteSoClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
  private
    { Private declarations }
    function GetPosNode(const aPos, aPosName: String; const aCanAdd: Boolean = True):
      TcxTreeListNode;
    function IsTNSExists(const aName: String): Boolean;  
    procedure GetTNSList;
    procedure GetSoList;
    procedure EnableControls;
    procedure DisableControls;
  public
    { Public declarations }
  end;

var
  fmConfig: TfmConfig;

implementation

uses cbMain, cbStyleModule, cbConfigModule;

var aGlobeKeyValue: String;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function GetSoImageIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbNoSelect: Result := 1;
  else
    Result := 2;
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsSelected(const aStatus: TDbConnectStatus): String;
begin
  Result := 'Y';
  if ( aStatus in [dbNoSelect] ) then Result := 'N';
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormCreate(Sender: TObject);
begin
  lblTest.Caption := EmptyStr;
  GetTNSList;
  GetSoList;
  ConfigPageChange( ConfigPage );
  DisableControls;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if not ( ssCtrl in Shift ) then aGlobeKeyValue := EmptyStr;
  if not ( Char( Key ) in ['a'..'z', 'A'..'Z'] ) then
    aGlobeKeyValue := EmptyStr
  else
    aGlobeKeyValue := ( aGlobeKeyValue + Char( Key ) );
  if UpperCase( aGlobeKeyValue ) = 'CABLESOFT' then
  begin
    EnableControls;
    aGlobeKeyValue := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormDestroy(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

function TfmConfig.GetPosNode(const aPos, aPosName: String; const aCanAdd: Boolean = True):
  TcxTreeListNode;
begin
  Result := nil;
  if ( aPosName = EmptyStr ) then Exit;
  Result := SoTree.FindNodeByText( aPosName, SoTreeColumn3 );
  if not Assigned( Result ) and aCanAdd then
  begin
    Result := SoTree.Add( nil );
    Result.AssignValues( [EmptyStr, EmptyStr, aPosName, EmptyStr,
      EmptyStr, EmptyStr, aPos] );
    Result.ImageIndex := 0;
    Result.SelectedIndex := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmConfig.IsTNSExists(const aName: String): Boolean;
begin
  Result := ( OracleAliasList.IndexOf( aName ) >= 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetTNSList;
begin
  TcxComboBoxProperties( SoTreeColumn7.Properties ).Items.Assign(
    OracleAliasList );
  TcxComboBoxProperties( SoTreeColumn12.Properties ).Items.Assign(
    OracleAliasList );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetSoList;
var
  aIndex: Integer;
  aSo: TAppSo;
  aNodePos, aNodeComp: TcxTreeListNode;
  aTNSName, aComTNSName, aSelected: String;
begin
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    aNodePos := GetPosNode( IntToStr( aSo.Pos ), aSo.PosName );
    aNodeComp := SoTree.AddChild( aNodePos );
    aSelected := IsSelected( aSo.DbConnectStatus );
    aTNSName := EmptyStr;
    if IsTNSExists( aSo.DbAliase ) then
      aTNSName := aSo.DbAliase;
    aComTNSName := EmptyStr;
    if IsTNSExists( aSo.ComDbAliase ) then
      aComTNSName := aSo.ComDbAliase; 
    aNodeComp.AssignValues( [aSelected, aSo.CompCode, aSo.CompName,
      aSo.DbLoginUser, aSo.DbLoginPass, aTNSName, aSo.Pos,
      aSo.ComDbLoginUser, aSO.ComDbLoginPass, aComTNSName] );
    aNodeComp.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
    aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
    aNodePos.Expanded := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.EnableControls;
begin
  SoTree.Bands[SoTree.Bands.Count - 1].Visible := True;
  SoTreeColumn5.Visible := True;
  SoTreeColumn6.Visible := True;
  SoTreeColumn7.Visible := True;
  SoTreeColumn10.Visible := True;
  SoTreeColumn11.Visible := True;
  SoTreeColumn12.Visible := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.DisableControls;
begin
  SoTree.Bands[SoTree.Bands.Count - 1].Visible := False;
  SoTreeColumn5.Visible := False;
  SoTreeColumn6.Visible := False;
  SoTreeColumn7.Visible := False;
  SoTreeColumn10.Visible := False;
  SoTreeColumn11.Visible := False;
  SoTreeColumn12.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.ConfigPageChange(Sender: TObject);
begin
  HeaderPanel.Parent := ConfigPage.ActivePage;
  StyleModule.ConfigImageList.GetIcon( ConfigPage.ActivePageIndex,
    HeaderImage.Picture.Icon );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.SoTreeFocusedNodeChanged(Sender: TObject;
  APrevFocusedNode, AFocusedNode: TcxTreeListNode);
begin
  if Assigned( AFocusedNode ) then
    SoTree.OptionsSelection.CellSelect := ( AFocusedNode.Level > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.SoTreeColumn1PropertiesEditValueChanged(
  Sender: TObject);
begin
  if ( VarToStrDef( SoTreeColumn1.EditValue, 'N' ) = 'N' ) then
    SoTree.FocusedNode.ImageIndex := 1
  else
    SoTree.FocusedNode.ImageIndex := 2;
  SoTree.FocusedNode.SelectedIndex := SoTree.FocusedNode.ImageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddSoClick(Sender: TObject);
var
  aParentNode, aNode: TcxTreeListNode;
begin
  if not Assigned( SoTree.FocusedNode ) then Exit;
  aParentNode := nil;
  if SoTree.FocusedNode.Level = 0 then
    aParentNode := SoTree.FocusedNode
  else begin
    if Assigned( SoTree.FocusedNode.Parent ) then
      aParentNode := SoTree.FocusedNode.Parent;
  end;
  if Assigned( aParentNode ) then
  begin
    aNode := aParentNode.AddChild;
    aNode.ImageIndex := 1;
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.AssignValues( ['N', EmptyStr, EmptyStr, EmptyStr, EmptyStr,
      EmptyStr, aParentNode.Values[SoTreeColumn9.ItemIndex],
      EmptyStr, EmptyStr, EmptyStr] );
    aNode.Focused := True;
    SoTree.SetFocus;
    SoTreeColumn2.Editing := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteSoClick(Sender: TObject);
var
  aParent: TcxTreeListNode;
begin
  if Assigned( SoTree.FocusedNode ) then
  begin
    if SoTree.FocusedNode.Level > 0 then
    begin
      aParent := SoTree.FocusedNode.Parent;
      SoTree.FocusedNode.Delete;
      if Assigned( aParent ) then aParent.Focused := True;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnTestClick(Sender: TObject);
var
  aCompName, aUser, aPass, aAlias, aMessage: String;
  aConnection: TADOConnection;
begin
  if not Assigned( SoTree.FocusedNode ) then Exit;
  if SoTree.FocusedNode.Level <= 0 then Exit;
  aCompName := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn3.ItemIndex], EmptyStr );
  aUser := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn5.ItemIndex], EmptyStr );
  aPass := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn6.ItemIndex], EmptyStr );
  aAlias := VarToStrDef( SoTree.FocusedNode.Values[SoTreeColumn7.ItemIndex], EmptyStr );
  if ( aUser = EmptyStr ) or ( aPass = EmptyStr ) or ( aAlias = EmptyStr ) then
  begin
    lblTest.Caption := Format( '系統台[%s],資料庫連結參數值輸值有誤。', [aCompName] );
    Exit;
  end;
  lblTest.Caption := Format( '系統台[%s],資料庫連結測試中.......。', [aCompName] );
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    aConnection := TADOConnection.Create( nil );
    aConnection.LoginPrompt := False;
    try
      aConnection.ConnectionString := Format(
        'Provider=MSDAORA.1;Persist Security Info=True;' +
        'Password=%s;User ID=%s;Data Source=%s;',
        [aPass, aUser, aAlias] );
      aMessage := EmptyStr;
      try
        aConnection.Connected := True;
        aMessage := Format( '系統台[%s],資料庫連結測試完成。', [aCompName] );
      except
        on E: Exception do
          aMessage := Format( '系統台[%s],資料庫連結測試錯誤, 訊息:%s。',
            [aCompName, E.Message] );
      end;
      if aConnection.Connected then aConnection.Connected := False;
      lblTest.Caption := aMessage;
    finally
      aConnection.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
