unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus, DB, Jpeg, ComCtrls, ImgList,
  dxBar, dxBarExtItems, cxClasses, cxGridStrs, cxEditConsts, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinsdxBarPainter, dxSkinscxPCPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee,ShellAPI,IdHTTP;

type
  TfrmMain = class(TForm)
    dxBarManager: TdxBarManager;
    MainImageList: TImageList;
    A01: TdxBarButton;
    A02: TdxBarButton;
    A04: TdxBarButton;
    A05: TdxBarButton;
    A06: TdxBarButton;
    A07: TdxBarButton;
    A08: TdxBarButton;
    A09: TdxBarButton;
    A0A: TdxBarButton;
    A0B: TdxBarButton;
    A0C: TdxBarButton;
    MenuGroupA: TdxBarSubItem;
    B01: TdxBarButton;
    B02: TdxBarButton;
    B03: TdxBarButton;
    B04: TdxBarButton;
    MenuGroupB: TdxBarSubItem;
    C01: TdxBarButton;
    C02: TdxBarButton;
    C03: TdxBarButton;
    C04: TdxBarButton;
    C05: TdxBarButton;
    MenuGroupC: TdxBarSubItem;
    D01: TdxBarButton;
    D02: TdxBarButton;
    D03: TdxBarButton;
    D04: TdxBarButton;
    D05: TdxBarButton;
    D06: TdxBarButton;
    D07: TdxBarButton;
    MenuGroupD: TdxBarSubItem;
    E01: TdxBarButton;
    E02: TdxBarButton;
    E04: TdxBarButton;
    E03: TdxBarButton;
    MenuGroupE: TdxBarSubItem;
    Z1: TdxBarButton;
    GroupA: TdxBarGroup;
    GroupMain: TdxBarGroup;
    GroupB: TdxBarGroup;
    GroupC: TdxBarGroup;
    GroupD: TdxBarGroup;
    GroupE: TdxBarGroup;
    dxBarDockControl1: TdxBarDockControl;
    stcCopyRight: TdxBarStatic;
    stcUser: TdxBarStatic;
    LoginTimer: TTimer;
    stcCompany: TdxBarImageCombo;
    ImgBg: TImage;
    stcStatic: TdxBarStatic;
    B05: TdxBarButton;
    D08: TdxBarButton;
    B06: TdxBarButton;
    B07: TdxBarButton;
    D09: TdxBarButton;
    B08: TdxBarButton;
    D0A: TdxBarButton;
    D0B: TdxBarButton;
    C06: TdxBarButton;
    B09: TdxBarButton;
    D0C: TdxBarButton;
    D0D: TdxBarButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure A05Click(Sender: TObject);
    procedure C01Click(Sender: TObject);
    procedure A06Click(Sender: TObject);
    procedure C05Click(Sender: TObject);
    procedure D01Click(Sender: TObject);
    procedure D02Click(Sender: TObject);
    procedure D03Click(Sender: TObject);
    procedure D04Click(Sender: TObject);
    procedure D05Click(Sender: TObject);
    procedure D06Click(Sender: TObject);
    procedure A01Click(Sender: TObject);
    procedure C03Click(Sender: TObject);
    procedure D07Click(Sender: TObject);
    procedure A09Click(Sender: TObject);
    procedure E02Click(Sender: TObject);
    procedure C04Click(Sender: TObject);
    procedure C02Click(Sender: TObject);
    procedure B02Click(Sender: TObject);
    procedure A0AClick(Sender: TObject);
    procedure B03Click(Sender: TObject);
    procedure E03Click(Sender: TObject);
    procedure B01Click(Sender: TObject);
    procedure A02Click(Sender: TObject);
    procedure A07Click(Sender: TObject);
    procedure E01Click(Sender: TObject);
    procedure A04Click(Sender: TObject);
    procedure A08Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Z1Click(Sender: TObject);
    procedure LoginTimerTimer(Sender: TObject);
    procedure stcCompanyChange(Sender: TObject);
    procedure A0BClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure E04Click(Sender: TObject);
    procedure A0CClick(Sender: TObject);
    procedure B04Click(Sender: TObject);
    procedure B05Click(Sender: TObject);
    procedure D08Click(Sender: TObject);
    procedure B06Click(Sender: TObject);
    procedure B07Click(Sender: TObject);
    procedure D09Click(Sender: TObject);
    procedure B08Click(Sender: TObject);
    procedure D0AClick(Sender: TObject);
    procedure D0BClick(Sender: TObject);
    procedure C06Click(Sender: TObject);
    procedure B09Click(Sender: TObject);
    procedure D0CClick(Sender: TObject);
    procedure D0DClick(Sender: TObject);
  private
    { Private declarations }
    FCompanyList: TList;
    FExpireDate: TDateTime;
    FFormCaption : String;
    procedure SetFormCompetence;
    procedure DoMenuItemEnable(const aFormId: String);
    procedure ChangeDevexComponentLang;
    function PostExample: string;
  public
    { Public declarations }
    function GetFormTitleString(const aFunctionId, aDesc: String): String;
    function ChangeInvoiceCompany(aConnSeq: Integer): Boolean;
    procedure BuildInvoiceCompanyItem;
    property ExpireDate: TDateTime read FExpireDate write FExpireDate;
    property CompanyList: TList read FCompanyList;
  end;

var
  frmMain: TfrmMain;

implementation

uses
  cbUtilis, dtmMainJU, dtmMainHU, Uotheru,
  frmLoginU, dtmMainU, xmlU,
  frmInvA01U, frmInvA02U, frmInvA04U, frmInvA05U, frmInvA06U, 
  frmInvA07_1U, frmInvA08U, frmInvA09U, frmInvA10U, frmInvA11U,
  frmInvB01U, frmInvB02U, frmInvB03U, frmInvB04U, frmInvB05U,
  frmInvB09U,
  frmInvC01U, frmInvC02U, frmInvC03U, frmInvC04U, frmInvC05U,
  frmInvD01U, frmInvD02U, frmInvD03U, frmInvD04U, frmInvD05U, frmInvD06U,
  frmInvD07U, frmInvD08U,
  frmInvE01U, frmInvE02U, frmInvE03U, frmInvE04U,frmInvB06U,frmInvB06_1U,
  frmInvD09U,frmInvD0AU,frmInvD0BU,frmInvD0CU,frmInvD0DU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormCreate(Sender: TObject);
var
  aPicPath: String;
begin
  FCompanyList := TList.Create;
  stcUser.Caption := Format( '?n?J?? : %s',
    [Nvl(dtmMain.getLoginUser, '?|???n?J' )] );
  aPicPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'Main.jpg';
  if FileExists( aPicPath ) then
    ImgBg.Picture.LoadFromFile( aPicPath );
  ForceDirectories( IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + GRID_FOLDER );
  { ???? Devex ????, ???^?????????? }
  ChangeDevexComponentLang;
{$IFDEF DEBUG}
   if ( Screen.MonitorCount > 1 ) then
   begin
     Self.Position := poDesigned;
     Self.Left := Screen.Monitors[1].Left + ( ( Screen.Monitors[1].Width - Width ) div 2 );
     Self.Top := Screen.Monitors[1].Top + ( ( Screen.Monitors[1].Height - Height) div 2);
   end;
{$ENDIF}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormShow(Sender: TObject);
begin
  LoginTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  FCompanyList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.LoginTimerTimer(Sender: TObject);
var
  aModalResult: Integer;
  aCompetence: String;
begin
  LoginTimer.Enabled := False;
  frmLogin := TfrmLogin.Create( Application );
  try
    frmLogin.CompanyList := FCompanyList;
    aModalResult := frmLogin.ShowModal;
  finally
    frmLogin.Free;
  end;
  if aModalResult <> mrOK then
  begin
    Application.Terminate;
    ExitProcess( 0 );
    Exit;
  end;
  Application.ProcessMessages;
  Self.Caption := Format( '?}???????o???t?? %s', [VERSION] );
  FFormCaption := Self.Caption;
  Application.Title := Self.Caption;
  stcUser.Caption := Format( '?n?J?? : %s',
    [Nvl(dtmMain.getLoginUser, '?|???n?J' )] );
  BuildInvoiceCompanyItem;
  Application.ProcessMessages;  
  aCompetence := dtmMain.LoadCompetence( dtmMain.getGroupId );
  Application.ProcessMessages;
  if ( aCompetence = EmptyStr ) then
  begin
    WarningMsg( '?z?L?????v???i?????o???t???C' );
    Exit;
  end;
  Screen.Cursor := crSQLWait;
  try
    setFormCompetence
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.GetFormTitleString(const aFunctionId, aDesc: String): String;
begin
  Result := Format( '[%s]-%s', [aFunctionId, aDesc] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.ChangeInvoiceCompany(aConnSeq: Integer): Boolean;
begin
  //#5358 ?W?[ShowFaci ?P?_?b?}???o?????O?_?n?N?]???????g?JMemo1 By Kin 2009/12/14
  Result := dtmMain.ConnectToINV( dtmMain.GetSoInfo( aConnSeq ) );
  if ( Result ) and ( dtmMain.GetLinkToMIS ) then
    Result := dtmMain.ConnectToSO( dtmMain.GetSoInfo( aConnSeq ) );
  if not Result then Exit;    
  dtmMain.sG_CompID := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).CompanyId;
  dtmMain.sG_CompName := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).CompanyName;
  dtmMain.sG_CompLongName := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).CompanyLongName;
  dtmMain.sG_ServiceTypeStr := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).ServiceType;
  dtmMain.sG_LinkToMIS := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).LinkToMIS;
  dtmMain.sG_ConnSeq := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).ConnSeq;
  dtmMain.sG_AutoCreateNum := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).AutoCreateNum;
  dtmMain.sG_IfPrintTitle := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).IfPrintTitle;
  dtmMain.sG_IfPrintAddr := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).IfPrintAddr;
  dtmMain.sG_IfPrintCheck := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).IfPrintCheck;
  dtmMain.sG_ExpAddrType := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).ExpAddrType;
  dtmMain.sG_CompTel := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).Tel;
  dtmMain.sG_ShowFaci := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).ShowFaci;
  dtmMain.sG_StarEInvoice := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarEInvoice;
  dtmMain.sG_StarEmail := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarEmail;
  dtmMain.sG_StarMessage := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarMessage;
  dtmMain.sG_StarTVmail :=PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarTVMail;
  dtmMain.sG_StarCMSend :=PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarCMSend;
  dtmMain.sG_StarCMTVmail := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarCMTVMail;
  dtmMain.sG_StarAllUpload := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).StarAllUpload;
  dtmMain.sG_UploadType := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).UploadType;
  dtmMain.sG_AESKey := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).AESKey;
  dtmMain.sG_UpdLimtDate := PCompany(
    FCompanyList[stcCompany.ItemIndex] ).UpdLimtDate;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.BuildInvoiceCompanyItem;
var
  aIndex, aSelected: Integer;
begin
  aSelected := -1;
  stcCompany.OnChange := nil;
  try
    stcCompany.Items.Clear;
    for aIndex := 0 to FCompanyList.Count - 1  do
    begin
      stcCompany.Items.Add( PCompany( FCompanyList[aIndex] ).CompanyName );
      if PCompany( FCompanyList[aIndex] ).CompanyId = dtmMain.getCompID then
        aSelected := aIndex;
      stcCompany.ImageIndexes[stcCompany.Items.Count - 1] := 8;
    end;
    if aSelected >= 0 then
    begin
     stcCompany.ItemIndex := aSelected;
     ChangeInvoiceCompany( PCompany( FCompanyList[aSelected] ).ConnSeq );
    end;
  finally

    {#6764 ???Y?W?[?????n?????T By Kin 2014/05/02}
    {#6764 ??????OK,?W?[???????q?O?N?X By Kin 2014/07/14}
    Self.Caption := Format('%s <<???q?N?X : %s, %s -- %s ( %s ) >>',
            [FFormCaption,dtmMain.getCompID, stcCompany.Text,
            dtmMain.getLoginUser,dtmMain.getLoginUserName]);
    //Self.Caption := Self.Caption + ' << ' +   stcCompany.Text + ' -- '   >>';
    stcCompany.OnChange := stcCompanyChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.SetFormCompetence;
var
  aFormId, aCompetence : String;
begin
  dtmMain.cdsCompetence.First;
  while not dtmMain.cdsCompetence.Eof do
  begin
    aFormId := dtmMain.cdsCompetence.FieldByName( 'ID' ).AsString;
    aCompetence := dtmMain.cdsCompetence.FieldByName( 'Query' ).AsString;
    if ( aCompetence = 'Y' ) then
      DoMenuItemEnable( aFormId );
    dtmMain.cdsCompetence.Next;
  end;
  { ?????]?w?v???Y?i???????\?? }
  { E04.?????K?X }
  E04.Visible := ivAlways;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.DoMenuItemEnable(const aFormId: String);
var
  aIndex: Integer;
  aItem: TdxBarItem;
  aItemLinks: TdxBarItemLinks;
begin
  aItem := dxBarManager.GetItemByName( aFormId );
  if not Assigned( aItem ) then Exit;
  for aIndex := 0 to GroupMain.Count - 1 do
  begin
    aItemLinks := TdxBarSubItem( GroupMain.Items[aIndex] ).ItemLinks;
    //if ( aItemLinks.IsReferencedBy( aItem as IdxBarLinksOwner ) ) then
    if ( aItemLinks.HasItem( aItem ) ) then
    begin
      TdxBarSubItem( GroupMain.Items[aIndex] ).Visible := ivAlways;
      Break;
    end;
  end;
  aItem.Visible := ivAlways;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ChangeDevexComponentLang;
begin
  cxSetResourceString( @scxGridDeletingConfirmationCaption, '?T?{' );
  cxSetResourceString( @scxGridDeletingFocusedConfirmationText, '?R???????????' );
  cxSetResourceString( @scxGridDeletingSelectedConfirmationText, '?R???????????????' );
  cxSetResourceString( @scxGridCustomizationFormCaption, '????????' );
  cxSetResourceString( @scxGridCustomizationFormBandsPageCaption, '????' );
  cxSetResourceString( @scxGridCustomizationFormColumnsPageCaption, '????' );
  cxSetResourceString( @scxGridNoDataInfoText, '?L?????i????' );
  cxSetResourceString( @cxSDatePopupToday, '????' );
  cxSetResourceString( @cxSDatePopupClear, '?M??' );
  cxSetResourceString( @cxSDatePopupNow, '?{?b' );
  cxSetResourceString( @cxSDateError, '???????~' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  aIndex: Integer;
begin
  CanClose := ConfirmMsg( '?T?{?o???????t???' );
  if CanClose then
  begin
    for aIndex := 0 to FCompanyList.Count - 1 do
    begin
      if Assigned( FCompanyList.Items[aIndex] ) then
      begin
        Dispose( PCompany( FCompanyList.Items[aIndex] ) );
        FCompanyList.Items[aIndex] := nil;
      end;
    end;
    FCompanyList.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A05Click(Sender: TObject);
begin
    frmInvA05 := TfrmInvA05.Create(Application);
    frmInvA05.ShowModal;
    frmInvA05.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.C01Click(Sender: TObject);
begin
    frmInvC01 := TfrmInvC01.Create(Application);
    frmInvC01.ShowModal;
    frmInvC01.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A06Click(Sender: TObject);
begin
    frmInvA06 := TfrmInvA06.Create(Application);
    frmInvA06.ShowModal;
    frmInvA06.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.C05Click(Sender: TObject);
begin
    frmInvC05 := TfrmInvC05.Create(Application);
    frmInvC05.ShowModal;
    frmInvC05.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D01Click(Sender: TObject);
begin
  frmInvD01 := TfrmInvD01.Create(Application);
  frmInvD01.ShowModal;
  frmInvD01.Free;
  stcCompany.OnChange := nil;
  try
    dtmMain.GetInvoiceCompany( FCompanyList );
    Self.BuildInvoiceCompanyItem;
  finally
    stcCompany.OnChange := stcCompanyChange;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D02Click(Sender: TObject);
begin
    frmInvD02 := TfrmInvD02.Create(Application);
    frmInvD02.ShowModal;
    frmInvD02.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D03Click(Sender: TObject);
begin
  frmInvD03 := TfrmInvD03.Create( Application );
  try
    frmInvD03.ShowModal;
  finally
    frmInvD03.Free;
  end;  
end;

procedure TfrmMain.D04Click(Sender: TObject);
begin
  frmInvD04 := TfrmInvD04.Create( Application );
  try
    frmInvD04.ShowModal;
  finally
    frmInvD04.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D05Click(Sender: TObject);
begin
  frmInvD05 := TfrmInvD05.Create( Application );
  try
    frmInvD05.ShowModal;
  finally
    frmInvD05.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D06Click(Sender: TObject);
begin
  frmInvD06 := TfrmInvD06.Create( Application );
  try
    frmInvD06.ShowModal;
  finally
    frmInvD06.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A01Click(Sender: TObject);

begin
  {
  ShellApplication('D:\App\Invoice\Bin\ElectronInv.exe','TEST',
  [dtmMain.getDbSID,dtmMain.getDbUserID,dtmMain.getDbPassword,dtmMain.getCompID,dtmMain.getMisDbOwner]);
  }
  
  frmInvA01 := TfrmInvA01.Create(Application);
  try
    frmInvA01.ShowModal;
  finally
    frmInvA01.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.C03Click(Sender: TObject);
begin
  frmInvC03 := TfrmInvC03.Create(Application);
  try
    frmInvC03.ShowModal;
  finally
    frmInvC03.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D07Click(Sender: TObject);
begin
  frmInvD07 := TfrmInvD07.Create(Application);
  try
    frmInvD07.ShowModal;
  finally
    frmInvD07.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.D08Click(Sender: TObject);
begin
  frmInvD08 := TfrmInvD08.Create(Application);
  try
    frmInvD08.ShowModal;
  finally
    frmInvD08.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A09Click(Sender: TObject);
begin
  frmInvA09 := TfrmInvA09.Create(Application);
  try
    frmInvA09.ShowModal;
  finally
    frmInvA09.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.E02Click(Sender: TObject);
begin
  frmInvE02 := TfrmInvE02.Create(Application);
  try
    frmInvE02.ShowModal;
  finally
    frmInvE02.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.C04Click(Sender: TObject);
begin
  frmInvC04 := TfrmInvC04.Create(Application);
  try
    frmInvC04.ShowModal;
  finally
    frmInvC04.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.C02Click(Sender: TObject);
begin
  frmInvC02 := TfrmInvC02.Create( Application );
  try
    frmInvC02.ShowModal;
  finally 
    frmInvC02.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.B02Click(Sender: TObject);
begin
  frmInvB02 := TfrmInvB02.Create(Application);
  try
    frmInvB02.InvId := EmptyStr;
    frmInvB02.ShowModal;
  finally
    frmInvB02.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A0AClick(Sender: TObject);
begin
  frmInvA10 := TfrmInvA10.Create(Application);
  try
    frmInvA10.ShowModal;
  finally
    frmInvA10.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.B03Click(Sender: TObject);
begin
  
  frmInvB03 := TfrmInvB03.Create(Application);
  try
    frmInvB03.ShowModal;
  finally
    frmInvB03.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.B04Click(Sender: TObject);
begin
  if not dtmMain.GetLinkToMIS then
  begin
    WarningMsg( Format( '?????q?O:%s?????]?w???P???A?t??????, ???\?????i?????C',
      [stcCompany.Text] ) );
    Exit;
  end;
  frmInvB04 := TfrmInvB04.Create( Application );
  try
    frmInvB04.ShowModal;
  finally
    frmInvB04.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.B05Click(Sender: TObject);
begin
  frmInvB05 := TfrmInvB05.Create( Application );
  try
    frmInvB05.ShowModal;
  finally
    frmInvB05.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.E03Click(Sender: TObject);
begin
  frmInvE03 := TfrmInvE03.Create(Application);
  try
    frmInvE03.ShowModal;
  finally
    frmInvE03.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.E04Click(Sender: TObject);
begin
  frmInvE04 := TfrmInvE04.Create( Application );
  try
    frmInvE04.ShowModal;
  finally
    frmInvE04.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.B01Click(Sender: TObject);
begin
  frmInvB01 := TfrmInvB01.Create(Application);
  try
    frmInvB01.ShowModal;
  finally
    frmInvB01.Free;
  end;  
  stcCompany.OnChange := nil;
  try
    dtmMain.GetInvoiceCompany( FCompanyList );
    Self.BuildInvoiceCompanyItem;
  finally
    stcCompany.OnChange := stcCompanyChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A02Click(Sender: TObject);
begin
  frmInvA02 := TfrmInvA02.Create(Application);
  try
    frmInvA02.CallFromProgram := nil;
    frmInvA02.ShowModal;
  finally
    frmInvA02.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A07Click(Sender: TObject);
begin

  frmInvA07_1 := TfrmInvA07_1.Create( Application );
  try
    frmInvA07_1.ShowModal;
  finally
    frmInvA07_1.Free;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.E01Click(Sender: TObject);
begin
  frmInvE01 := TfrmInvE01.Create( Application );
  try
    frmInvE01.ShowModal;
  finally
    frmInvE01.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A04Click(Sender: TObject);
begin
  if not dtmMain.GetLinkToMIS then
  begin
    WarningMsg( Format( '?????q?O:%s?????]?w???P???A?t??????, ???\?????i?????C',
      [stcCompany.Text] ) );
    Exit;
  end;
  frmInvA04 := TfrmInvA04.Create(Application);
  try
    frmInvA04.ShowModal;
  finally
    frmInvA04.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A08Click(Sender: TObject);
begin
  frmInvA08 := TfrmInvA08.Create(Application);
  try
    frmInvA08.ShowModal;
  finally
    frmInvA08.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A0BClick(Sender: TObject);
begin
  frmInvA11 := TfrmInvA11.Create( Application );
  try
    frmInvA11.ShowModal;
  finally
    frmInvA11.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.A0CClick(Sender: TObject);
begin
  frmInvA02 := TfrmInvA02.Create( Application );
  try
    frmInvA02.CallFromProgram := Self.ClassType;
    frmInvA02.ShowModal;
  finally
    frmInvA02.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Z1Click(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.stcCompanyChange(Sender: TObject);
var
  aCompName: String;
  aSoInfo: PSMSDb;

  { ------------------------------------------------------ }

  procedure AbortChange;
  var
    aIndex: Integer;
  begin
    for aIndex := 0 to FCompanyList.Count - 1 do
    begin
      if ( PCompany( FCompanyList[aIndex] ).CompanyId = dtmMain.getCompID ) then
      begin
        stcCompany.ItemIndex := aIndex;
        Break;
      end;
    end;
  end;

  { ------------------------------------------------------ }

begin
  aCompName := PCompany( FCompanyList[stcCompany.ItemIndex] ).CompanyName;
  stcCompany.OnChange := nil;
  try
    if not ConfirmMsg(
      Format( '?T?{???????q?O?? ?i%s?j ?', [aCompName] ) ) then
    begin
      AbortChange;
      Exit;
    end;
    aSoInfo := dtmMain.GetSoInfo( PCompany( FCompanyList[stcCompany.ItemIndex] ).ConnSeq );
    if not Assigned( aSoInfo ) then
    begin
      ErrorMsg( '?L?k???????q?O, ???]:?????????t???x?C' );
      AbortChange;
      Exit;
    end;      
    Screen.Cursor := crSQLWait;
    try
      Sleep( 300 );
      if not ChangeInvoiceCompany( PCompany(
        FCompanyList[stcCompany.ItemIndex] ).ConnSeq ) then
      begin
        AbortChange;
      end;
    finally

      Screen.Cursor := crDefault;
    end;
  finally
    {#6764 ??????OK,?n?P?BForm.Caption By Kin 2014/07/14}
    Self.Caption := Format('%s <<???q?N?X : %s, %s -- %s ( %s ) >>',
            [FFormCaption,dtmMain.getCompID, stcCompany.Text,
            dtmMain.getLoginUser,dtmMain.getLoginUserName]);
    stcCompany.OnChange := stcCompanyChange;
  end;
end;

{ ---------------------------------------------------------------------------- }



procedure TfrmMain.B06Click(Sender: TObject);

begin
  try

    frmInvB06_1 := TfrmInvB06_1.Create( Application );
    frmInvB06_1.ShowModal;
  finally
    dtmMain.EInvConnection.Connected := False;
    frmInvB06_1.Free;

  end;


 {
  aSoInfo := PSMSDb(
      dtmMain.GetSoInfo( dtmMain.GetConnSeq ) ) ;
  dtmMain.ConnectToEINV( aSoInfo);
  frmInvB06 := TfrmInvB06.Create( Application );

  try
    dtmMainH.getInv035Data(EmptyStr);

    frmInvB06.ShowModal;
  finally
    dtmMain.EInvConnection.Connected := False;

    frmInvB06.Free;
  end;
  }
end;

procedure TfrmMain.B07Click(Sender: TObject);
var
  aExeFileName: String;
  aParam : array[0..16] of string;

begin
  aExeFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_ELECTRON_UPLOAD_EXE;
  aParam[0] := dtmMain.getDbSID;
  aParam[1] := dtmMain.getDbUserID;
  aParam[2] := dtmMain.getDbPassword;
  aParam[3] := dtmMain.getCompID;
  aParam[4] := dtmMain.getMisDbOwner;
  
  if dtmMain.sG_StarAutoNotify then
    aParam[5] := '1'
  else
    aParam[5] := '0';

  if dtmMain.GetUseCOM then
  begin
    aParam[6] := dtmMain.GetEInvDb;
    aParam[7] := dtmMain.GetEInvUser;
    aParam[8] := dtmMain.GetEInvPassword;
  end else
  begin
    aParam[6] := dtmMain.getDbSID;
    aParam[7] := dtmMain.getDbUserID;
    aParam[8] := dtmMain.getDbPassword;
  end;
  aParam[9] := dtmMain.getLoginUser;
  aParam[10] := dtmMain.getLoginUserName;
  aParam[11] :='X';
  if dtmMain.GetMisDbLink <> EmptyStr then
  begin
    aParam[11] := dtmMain.GetMisDbLink;
  end;
  {#6370 ?P?_?O?_????StartAllUpload By Kin 2012/11/23}
  if dtmMain.sG_StarAllUpload then
    aParam[12] := '1'
  else
    aParam[12] := '0';

  {?W?[MisDBSid?BMisDBUserId?BMisDBPassword?BMisOwner For Miggie By Kin 2013/01/18}
  aParam[13] := 'X';
  aParam[14] := 'X';
  aParam[15] := 'X';
  aParam[16] := 'X';
  if dtmMain.GetMisDBSid <> '' then
    aParam[13] :=  dtmMain.GetMisDBSid;
  if dtmMain.GetMisDBUserId <> '' then
    aParam[14] := dtmMain.GetMisDBUserId;
  if dtmMain.GetMisDBPassword <> '' then
    aParam[15] := dtmMain.GetMisDBPassword;
  if dtmMain.getMisDbOwner <> '' then
    aParam[16] := dtmMain.getMisDbOwner;
  dtmMain.ShellApplication(aExeFileName,EmptyStr, aParam );

end;

procedure TfrmMain.D09Click(Sender: TObject);
begin
  frmInvD09 := TfrmInvD09.Create(Application);
  try
    frmInvD09.ShowModal;
  finally
    frmInvD09.Free;
  end;
end;

procedure TfrmMain.B08Click(Sender: TObject);
var
  aExeFileName: String;
  aParam : array[0..20] of string;
begin
  if not dtmMain.GetUseCOM then
  begin
    WarningMsg( 'INI???]?w???~?I?L?k????' );
  end;
  aExeFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_ELECTRON_PROClIST_EXT;
  aParam[0] := dtmMain.getCompID;
  aParam[1] := dtmMain.getDbSID;
  aParam[2] := dtmMain.getDbUserID;
  aParam[3] := dtmMain.getDbPassword;
  if dtmMain.GetUseCOM then
  begin
    aParam[4] := dtmMain.GetEInvDb;
    aParam[5] := dtmMain.GetEInvUser;
    aParam[6] := dtmMain.GetEInvPassword;
  end else
  begin
    aParam[4] := dtmMain.getDbSID;
    aParam[5] := dtmMain.getDbUserID;
    aParam[6] := dtmMain.getDbPassword;
  end;
  aParam[7] := dtmMain.getLoginUser;
  aParam[8] := dtmMain.getLoginUserName;
  aParam[9] := '0';
  aParam[10] := '0';
  aParam[11] := '0';
  aParam[12] := '0';
  aParam[13] := '0';
  aParam[14] := '0';
  aParam[15] := 'X';

  if dtmMain.GetMisDbLink <> EmptyStr then
  begin
    aParam[15] := dtmMain.GetMisDbLink;
  end;
  {#6396 ?W?[UploadTime,?]???q?o???t???????i?JB08,?????n??UploadTime ???H??X By Kin 2012/12/18}
  aParam[16] := 'X';


  {?W?[MisDBSid?BMisDBUserId?BMisDBPassword?BMisOwner For Miggie By Kin 2013/01/18}
  aParam[17] := 'X';
  aParam[18] := 'X';
  aParam[19] := 'X';
  aParam[20] := 'X';
  if dtmMain.GetMisDBSid <> '' then
    aParam[17] :=  dtmMain.GetMisDBSid;
  if dtmMain.GetMisDBUserId <> '' then
    aParam[18] := dtmMain.GetMisDBUserId;
  if dtmMain.GetMisDBPassword <> '' then
    aParam[19] := dtmMain.GetMisDBPassword;
  if dtmMain.getMisDbOwner <> '' then
    aParam[20] := dtmMain.getMisDbOwner;
  dtmMain.ShellApplication(aExeFileName,EmptyStr, aParam );
//  dtmMain.ShellApplication( aExeFileName,EmptyStr,aParam );
  //aParam[1] := dtmMain.getDbUserID;
  //aParam[2] := dtmMain.getDbPassword;
  //aParam[3] := dtmMain.getCompID;
  //aParam[4] := dtmMain.getMisDbOwner;
  //dtmMain.ShellApplication(aExeFileName,EmptyStr, aParam );
end;

procedure TfrmMain.D0AClick(Sender: TObject);
begin
  frmInvD0A := TfrmInvD0A.Create(Application);
  try
    frmInvD0A.ShowModal;
  finally
    frmInvD0A.Free;
  end;
end;

procedure TfrmMain.D0BClick(Sender: TObject);
begin
  if not dtmMain.sG_StarAutoNotify then
  begin
    MessageDlg('?q?l?o???q?????{?]?w???????A?????]?w?I',mtInformation,[mbOk],0);
    Exit;
  end;
  frmInvD0B := TfrmInvD0B.Create(Application);
  try
    frmInvD0B.ShowModal;
  finally
    frmInvD0B.Free;
  end;
end;

procedure TfrmMain.C06Click(Sender: TObject);
var 
  aExeFileName: String;
  aParam : array[0..5] of string;

begin
  aExeFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_NOTIFYRPT_C06_EXE;
  aParam[0] := dtmMain.getDbSID;  
  aParam[1] := dtmMain.getDbUserID;
  aParam[2] := dtmMain.getDbPassword;
  aParam[3] := dtmMain.getCompID;
  aParam[4] := dtmMain.getLoginUserName;
  aParam[5] := IncludeTrailingPathDelimiter( ExtractFilePath(
              Application.ExeName ) ) + IncludeTrailingPathDelimiter(
              REPORT_FOLDER );
  dtmMain.ShellApplication(aExeFileName,EmptyStr, aParam );

end;

procedure TfrmMain.B09Click(Sender: TObject);
begin
  try
    frmInvB09 := TfrmInvB09.Create( Application );
    frmInvB09.ShowModal;
  finally
    frmInvB09.Free;
  end;
end;

procedure TfrmMain.D0CClick(Sender: TObject);
begin

  frmInvD0C := TfrmInvD0C.Create(Application);
  try
    frmInvD0C.ShowModal;
  finally
    frmInvD0C.Free;
  end;
end;

procedure TfrmMain.D0DClick(Sender: TObject);
begin
   frmInvD0D := TfrmInvD0D.Create( Application );
  try
    frmInvD0D.ShowModal;
  finally
    frmInvD0D.Free;
  end;
end;

function TfrmMain.PostExample: string;
  var
  lHTTP: TIdHTTP;
  lParamList: TStringList;
  sResult:string;
begin
  lParamList := TStringList.Create;
  lParamList.Add('id=1');

  lHTTP := TIdHTTP.Create(nil);
  try
    sResult :=lHTTP.Get('www.yahoo.com.tw');
    Result := sResult;
  finally
    FreeAndNil(lHTTP);
    FreeAndNil(lParamList);
  end;

end;

end.
