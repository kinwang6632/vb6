unit cbConfig;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, ExtCtrls,
{ App: }
  cbLanguage,
{ Deveoper Express: }
  cxLookAndFeelPainters, dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxStyles, cxControls, cxInplaceContainer,
  cxTL, cxPC, cxButtons, cxCheckBox, cxTextEdit, cxContainer, cxEdit,
  cxMaskEdit, cxSpinEdit, cxDropDownEdit, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary,
  dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin,
  dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2007Blue,
  dxSkinOffice2007Green, dxSkinOffice2007Pink, dxSkinOffice2007Silver,
  dxSkinSilver, dxSkinStardust, dxSkinSummer2008, dxSkinValentine,
  dxSkinXmas2008Blue;

type
  TfmConfig = class(TForm)
    Panel1: TPanel;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    Bevel1: TBevel;
    ConfigPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    HeaderPanel: TPanel;
    HeaderImage: TImage;
    cxTabSheet2: TcxTabSheet;
    lblBusyTimeReadFrquence: TLabel;
    cxTabSheet3: TcxTabSheet;
    lblTest: TLabel;
    SoTree: TcxTreeList;
    SoTreeCol1: TcxTreeListColumn;
    SoTreeCol2: TcxTreeListColumn;
    SoTreeCol3: TcxTreeListColumn;
    SoTreeCol4: TcxTreeListColumn;
    SoTreeCol5: TcxTreeListColumn;
    SoTreeCol6: TcxTreeListColumn;
    btnDeleteSo: TcxButton;
    btnAddSo: TcxButton;
    btnTestSo: TcxButton;
    Label1: TLabel;
    txtDbValidateSessionFreq: TcxSpinEdit;
    Label2: TLabel;
    txtDbGetPoolObjectRetryCount: TcxSpinEdit;
    Label3: TLabel;
    txtDbSyncUser: TcxSpinEdit;
    Label4: TLabel;
    txtUIGetUserListFreq: TcxSpinEdit;
    Label5: TLabel;
    txtAutoRefresh: TcxComboBox;
    Label6: TLabel;
    Label7: TLabel;
    txtAuthorizeRefreshRate: TcxSpinEdit;
    Label8: TLabel;
    txtAnnRefreshRate: TcxSpinEdit;
    Label9: TLabel;
    txtUserRefreshRate: TcxSpinEdit;
    SoTreeCol7: TcxTreeListColumn;
    SoTreeCol8: TcxTreeListColumn;
    Label10: TLabel;
    txtTryReconnectRate: TcxSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure ConfigPageChange(Sender: TObject);
    procedure btnAddSoClick(Sender: TObject);
    procedure btnDeleteSoClick(Sender: TObject);
    procedure btnTestSoClick(Sender: TObject);
    procedure txtAutoRefreshPropertiesChange(Sender: TObject);
  private
    { Private declarations }
    procedure InitLanguage;
    procedure InitControl;
    procedure GetCommEnv;
    procedure GetClientEnv;
    procedure GetSoList;
  public
    { Public declarations }
  end;

var
  fmConfig: TfmConfig;

implementation

{$R *.dfm}

uses cbUtilis, cbStyleModule, cbConfigModule, cbOraModule;


{ ---------------------------------------------------------------------------- }

procedure TfmConfig.FormCreate(Sender: TObject);
begin
  GetCommEnv;
  GetClientEnv;
  GetSoList;
  InitLanguage;
  InitControl;
  ConfigPage.ActivePageIndex := 0;
  ConfigPageChange( ConfigPage );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.ConfigPageChange(Sender: TObject);
begin
  HeaderPanel.Parent := ConfigPage.ActivePage;
  StyleModule.ImgList32.GetIcon( StyleModule.ConfigHeaderImage(
    ConfigPage.ActivePageIndex ), HeaderImage.Picture.Icon );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.InitLanguage;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ConfigPage.PageCount - 1 do
  begin
    ConfigPage.Pages[aIndex].Caption := LanguageManager.Get(
      ConfigPage.Pages[aIndex].Name );
  end;
  Label1.Caption := LanguageManager.Get( Label1.Name );
  Label2.Caption := LanguageManager.Get( Label2.Name );
  Label3.Caption := LanguageManager.Get( Label3.Name );
  Label4.Caption := LanguageManager.Get( Label4.Name );
  Label5.Caption := LanguageManager.Get( Label5.Name );
  Label6.Caption := LanguageManager.Get( Label6.Name );
  Label7.Caption := LanguageManager.Get( Label7.Name );
  Label8.Caption := LanguageManager.Get( Label8.Name );
  Label9.Caption := LanguageManager.Get( Label9.Name );
  Label10.Caption := LanguageManager.Get( Label10.Name );
  {}
  for aIndex := 0 to SoTree.ColumnCount - 1 do
  begin
    SoTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      SoTree.Columns[aIndex].Name );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.InitControl;
begin
  lblTest.Caption := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetCommEnv;
begin
  txtDbValidateSessionFreq.Value := ConfigModule.CommEnv.DbVaildateSessionFreq;
  txtDbGetPoolObjectRetryCount.Value := ConfigModule.CommEnv.DbGetPoolObjectRertyCount;
  txtDbSyncUser.Value := ConfigModule.CommEnv.DbSyncUser;
  txtUIGetUserListFreq.Value := ConfigModule.CommEnv.UIGetUserListFreq;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetClientEnv;
begin
  txtAutoRefresh.ItemIndex := 0;
  if ( not ConfigModule.ClientEnv.AutoRefresh ) then
    txtAutoRefresh.ItemIndex := 1;
  txtAuthorizeRefreshRate.Value := ConfigModule.ClientEnv.AuthorizeRefreshRate;
  txtAnnRefreshRate.Value := ConfigModule.ClientEnv.AnnRefreshRate;
  txtUserRefreshRate.Value := ConfigModule.ClientEnv.UserRefreshRate;
  txtTryReconnectRate.Value := ConfigModule.ClientEnv.TryReconnectRate;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.GetSoList;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
begin
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aNode := SoTree.Add;
    aNode.ImageIndex := StyleModule.SoImageIndex( ConfigModule.SoList.Items[aIndex] );
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.Values[SoTreeCol1.ItemIndex] := ConfigModule.SoList.Items[aIndex].Selected;
    aNode.Values[SoTreeCol2.ItemIndex] := ConfigModule.SoList.Items[aIndex].CompCode;
    aNode.Values[SoTreeCol3.ItemIndex] := ConfigModule.SoList.Items[aIndex].CompName;
    aNode.Values[SoTreeCol4.ItemIndex] := ConfigModule.SoList.Items[aIndex].DbUserId;
    aNode.Values[SoTreeCol5.ItemIndex] := ConfigModule.SoList.Items[aIndex].DbUserPass;
    aNode.Values[SoTreeCol6.ItemIndex] := ConfigModule.SoList.Items[aIndex].DbAliase;
    aNode.Values[SoTreeCol7.ItemIndex] := ConfigModule.SoList.Items[aIndex].SynData;
    aNode.Values[SoTreeCol8.ItemIndex] := ConfigModule.SoList.Items[aIndex].MaxSession;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnAddSoClick(Sender: TObject);
var
  aNode: TcxTreeListNode;
begin
  aNode := SoTree.Add;
  aNode.ImageIndex := 1;
  aNode.SelectedIndex := aNode.ImageIndex;
  aNode.Values[SoTreeCol1.ItemIndex] := False;
  aNode.Values[SoTreeCol2.ItemIndex] := EmptyStr;
  aNode.Values[SoTreeCol3.ItemIndex] := EmptyStr;
  aNode.Values[SoTreeCol4.ItemIndex] := EmptyStr;
  aNode.Values[SoTreeCol5.ItemIndex] := EmptyStr;
  aNode.Values[SoTreeCol6.ItemIndex] := EmptyStr;
  aNode.Values[SoTreeCol7.ItemIndex] := False;
  aNode.Values[SoTreeCol8.ItemIndex] := 0;
  aNode.MakeVisible;
  SoTree.SetFocus;
  aNode.Focused := True;
  SoTreeCol2.Editing := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnDeleteSoClick(Sender: TObject);
begin
  if Assigned( SoTree.FocusedNode ) then
    SoTree.FocusedNode.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.btnTestSoClick(Sender: TObject);
var
  aCompName, aUser, aPass, aAlias, aMessage: String;
  aOraModule: TOraModule;
begin
  if not Assigned( SoTree.FocusedNode ) then Exit;
  aCompName := VarToStrDef( SoTree.FocusedNode.Values[SoTreeCol3.ItemIndex], EmptyStr );
  aUser := VarToStrDef( SoTree.FocusedNode.Values[SoTreeCol4.ItemIndex], EmptyStr );
  aPass := VarToStrDef( SoTree.FocusedNode.Values[SoTreeCol5.ItemIndex], EmptyStr );
  aAlias := VarToStrDef( SoTree.FocusedNode.Values[SoTreeCol6.ItemIndex], EmptyStr );
  if ( aUser = EmptyStr ) or ( aPass = EmptyStr ) or ( aAlias = EmptyStr ) then
  begin
    lblTest.Caption := Format( LanguageManager.Get( 'SConfigDbTestParamIsEmpty' ),
      [aCompName] );
    Exit;
  end;
  lblTest.Caption := Format( LanguageManager.Get( 'SConfigDbTestNow' ), [aCompName] );
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  aOraModule := TOraModule.Create( nil );
  try
    aOraModule.OraSession.LoginPrompt := False;
    try
      aOraModule.OraSession.ConnectString := Format( '%s/%s@%s', [aUser, aPass, aAlias] );
      aMessage := EmptyStr;
      try
        aOraModule.OraSession.Connected := True;
        aMessage := Format( LanguageManager.Get( 'SConfigDbTestOk' ),
          [aCompName] );
      except
        on E: Exception do
          aMessage := Format( LanguageManager.Get( 'SConfigDbTestError' ),
            [aCompName, E.Message] );
      end;
      if ( aOraModule.OraSession.Connected ) then
        aOraModule.OraSession.Connected := False;
      lblTest.Caption := aMessage;
    finally
      aOraModule.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmConfig.txtAutoRefreshPropertiesChange(Sender: TObject);
var
  aSubItemEnable: Boolean;
begin
  aSubItemEnable := ( txtAutoRefresh.ItemIndex = 0 );
  txtAuthorizeRefreshRate.Enabled := aSubItemEnable;
  txtAnnRefreshRate.Enabled := aSubItemEnable;
  txtUserRefreshRate.Enabled := aSubItemEnable;
end;

{ ---------------------------------------------------------------------------- }

end.
