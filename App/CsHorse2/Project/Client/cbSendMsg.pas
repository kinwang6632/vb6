unit cbSendMsg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, ComCtrls, StdCtrls, Menus, ImgList, ActnList, DB,
{ App: }
  cbStyleModule, cbClientClass, cbDataClientModule, cbControls, cbHrHelper,
  cbROClientModule,
{ ODAC: }
  MemDS, VirtualTable,
{ Developer Express: }
  cxContainer, cxControls, cxClasses, cxGraphics, cxStyles, cxLookAndFeels,
  dxOffice11, cxGeometry, cxLabel,
  dxRibbon, dxRibbonForm, dxRibbonGallery, dxRibbonGalleryFilterEd,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxRibbonPainter, dxSkinscxPCPainter,
  dxSkinsdxBarPainter, cxLookAndFeelPainters,
  dxBar, cxPC, cxListView, cxEdit, cxGroupBox, cxTextEdit, dxBarExtItems,
  dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkiniMaginary, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Silver,
  cxTL,
{ RemObject: }
  uROTypes, CsHorse2Library_Intf,
{ RichView: }
  RVStyle, RVScroll, RichView, RVEdit;

type

  TfmSendMsg = class(THrClientForm)
    BarManager: TdxBarManager;
    dxBarButtonOpen: TdxBarLargeButton;
    dxBarButtonExit: TdxBarLargeButton;
    dxBarFontSize: TdxBarCombo;
    dxBarFontBold: TdxBarLargeButton;
    dxBarFontItalic: TdxBarLargeButton;
    dxBarFontUnderline: TdxBarLargeButton;
    dxBarFontBullets: TdxBarLargeButton;
    dxBarFontAlignLeft: TdxBarLargeButton;
    dxBarFontAlignCenter: TdxBarLargeButton;
    dxBarFontAlignRight: TdxBarLargeButton;
    OpenDialog: TOpenDialog;
    FontDialog: TFontDialog;
    dxBarPopupMenu: TdxRibbonPopupMenu;
    dxBarFontName: TdxBarFontNameCombo;
    tabEditMsg: TdxRibbonTab;
    Ribbon: TdxRibbon;
    dxBarGroup1: TdxBarGroup;
    Bar1: TdxBar;
    Bar2: TdxBar;
    Bar3: TdxBar;
    rgiFontColor: TdxRibbonGalleryItem;
    rgiColorTheme: TdxRibbonGalleryItem;
    dxBarFontColor: TdxBarButton;
    dxRibbonDropDownGallery: TdxRibbonDropDownGallery;
    RStyle: TRVStyle;
    MsgPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    BfRecver: TVirtualTable;
    LvUserList: TcxListView;
    Bar4: TdxBar;
    dxBarViewUsers: TdxBarLargeButton;
    dxBarSend: TdxBarLargeButton;
    dxBarPriority: TdxBarLargeButton;
    cxGroupBox1: TcxGroupBox;
    RView: TRichViewEdit;
    cxGroupBox2: TcxGroupBox;
    txtMsgSubject: TcxTextEdit;
    Timer1: TTimer;
    cxGroupBox3: TcxGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Image1: TImage;
    Label4: TLabel;
    Label5: TcxLabel;
    procedure FormCreate(Sender: TObject);
    procedure RViewStyleConversion(Sender: TCustomRichViewEdit; StyleNo,
      UserData: Integer; AppliedToText: Boolean; var NewStyleNo: Integer);
    procedure dxBarFontNameChange(Sender: TObject);
    procedure dxBarFontSizeChange(Sender: TObject);
    procedure Bar2CaptionButtons0Click(Sender: TObject);
    procedure dxBarFontColorClick(Sender: TObject);
    procedure RViewCurParaStyleChanged(Sender: TObject);
    procedure RViewCurTextStyleChanged(Sender: TObject);
    procedure dxBarFontBoldClick(Sender: TObject);
    procedure dxBarFontItalicClick(Sender: TObject);
    procedure dxBarFontUnderlineClick(Sender: TObject);
    procedure dxBarFontAlignLeftClick(Sender: TObject);
    procedure dxBarFontAlignCenterClick(Sender: TObject);
    procedure dxBarFontAlignRightClick(Sender: TObject);
    procedure RViewParaStyleConversion(Sender: TCustomRichViewEdit; StyleNo,
      UserData: Integer; AppliedToText: Boolean; var NewStyleNo: Integer);
    procedure dxBarFontBulletsClick(Sender: TObject);
    procedure RViewCaretMove(Sender: TObject);
    procedure dxBarButtonOpenClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dxBarViewUsersClick(Sender: TObject);
    procedure dxBarSendClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    FColorPickerController: TColorPickerController;
    FFontName: String;
    FFontSize: Integer;
    FAlreadySend: Boolean;
    FMsgFormType: Integer;
    FMsgInfo: TMsgInfo;
    FMsgNode: TcxTreeListNode;
    procedure AssignFontColorGlyph;
    procedure FontColorChanged(Sender: TObject);
    procedure BuildMsgRecver;
    procedure SetMsgFormType;
    procedure DeleteNoneNeedFont;
    function CheckRViewIsEmpty: Boolean;
    function DoSendMsg: Boolean;
    function GetMsg: Boolean;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent);
    procedure ChangeReadyState(const AReady: Boolean);
    property MsgFormType: Integer read FMsgFormType write FMsgFormType;
    property MsgInfo: TMsgInfo read FMsgInfo write FMsgInfo;
    property MsgNode: TcxTreeListNode read FMsgNode write FMsgNode;
  end;

var
  fmSendMsg: TfmSendMsg;

implementation

uses cbUtilis, cbMain;

{$R *.DFM}

{ ---------------------------------------------------------------------------- }

{ TfmSendMsg }

constructor TfmSendMsg.Create(AOwner: TComponent);
begin
  inherited Create( AOwner, ftMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.FormCreate(Sender: TObject);
begin
  FColorPickerController := TColorPickerController.Create( rgiFontColor,
    rgiColorTheme, dxRibbonDropDownGallery );
  FColorPickerController.OnColorChanged := FontColorChanged;
  AssignFontColorGlyph;
  Ribbon.ApplicationButton.Visible := False;
  DataClientModule.AddForm( Self );
  BarManager.ImageOptions.Images := StyleModule.SmallImages;
  BarManager.ImageOptions.LargeImages := StyleModule.LargeImages;
  dxBarFontName.Text := RStyle.TextStyles[0].FontName;
  dxBarFontSize.Text := IntToStr( RStyle.TextStyles[0].Size );
  LvUserList.SmallImages := StyleModule.SmallImages;
  dxBarFontAlignLeft.Down := True;
  Self.Caption := '訊息';
  ChangeColorSchema( StyleModule.ColorName );
  TBufferHelper.CreateFieldDefs( biRecvList, BfRecver );
  MsgPage.HideTabs := True;
  MsgPage.ActivePageIndex := 0;
  dxBarViewUsers.Down := False;
  FAlreadySend := False;
  FMsgFormType := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.FormShow(Sender: TObject);
begin
  Screen.Cursor := crAppStart;
  try
    SetMsgFormType;
    if ( FMsgFormType = 0 ) then
    begin
      DeleteNoneNeedFont;
      BuildMsgRecver;
    end else
      Timer1.Enabled := True;  
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  if ( FMsgFormType = 0 ) then
  begin
    if ( not FAlreadySend ) and ( not CheckRViewIsEmpty ) then
    begin
      if Application.MessageBox( '確定不發送訊息?', '確認', MB_OKCANCEL + MB_ICONQUESTION ) = ID_CANCEL then
        Action := caNone
    end;
  end;
  if ( Action = caFree ) then
  begin
    if ( FMsgFormType = 1 ) and Assigned( FMsgNode ) and Assigned( FMsgInfo ) then
    begin
      if ( FMsgInfo.IsRead ) then DataClientModule.DeleteMsgItem( FMsgNode );
    end;
    DataClientModule.RemoveForm( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.FormDestroy(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.BuildMsgRecver;
var
  aItem: TListItem;
begin
  LvUserList.Clear;
  LvUserList.Items.BeginUpdate;
  try
    BfRecver.First;
    while not BfRecver.Eof do
    begin
      aItem := LvUserList.Items.Add;
      aItem.Caption := BfRecver.FieldByName( 'DisplayText' ).AsString;
      aItem.ImageIndex := StyleModule.GetUserStatusImageIndex(
        BfRecver.FieldByName( 'Status' ).AsInteger );
      BfRecver.Next;
    end;
  finally
    LvUserList.Items.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmSendMsg.CheckRViewIsEmpty: Boolean;
var
  aIndex: Integer;
  aText: String;
begin
  Result := True;
  for aIndex := 0 to RView.ItemCount - 1 do
  begin
    if ( RView.GetItemStyle( aIndex ) >= 0 ) then
    begin
      aText := ( aText + Trim( RView.GetItemText( aIndex ) ) );
      Result := ( aText = EmptyStr );
    end else
      Result := False;
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.SetMsgFormType;
begin
  if ( FMsgFormType = 1 ) then
  begin
    RView.ReadOnly := True;
    txtMsgSubject.Properties.ReadOnly := True;
    cxGroupBox2.Visible := False;
    cxGroupBox3.Visible := True;
    Label1.Caption := EmptyStr;
    Label5.Caption := EmptyStr;
    Ribbon.ViewInfo.ForceHide := True;
  end else
  begin
    cxGroupBox2.Visible := True;
    cxGroupBox3.Visible := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.DeleteNoneNeedFont;
var
  aIndex: Integer;
begin
  for aIndex := dxBarFontName.Items.Count - 1 downto 0  do
  begin
    if IsDelimiter( Chr( 64 ), dxBarFontName.Items[aIndex], 1 ) then
      dxBarFontName.Items.Delete( aIndex );
  end;
  aIndex := dxBarFontName.Items.IndexOf( 'Arial' );
  if aIndex >= 0 then
    dxBarFontName.ItemIndex := aIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.AssignFontColorGlyph;
begin
  dxBarFontColor.Glyph := rgiFontColor.Glyph;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.ChangeReadyState(const AReady: Boolean);
begin
  dxBarSend.Enabled := AReady;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewCaretMove(Sender: TObject);
var
  aItemNo: Integer;
begin
  aItemNo := RView.CurItemNo;
  if ( aItemNo < 0 ) then Exit;
  while not ( RView.IsParaStart( aItemNo ) ) do
    Dec( aItemNo );
  dxBarFontBullets.Down := ( RView.GetItemStyle( aItemNo ) = rvsListMarker )
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewCurParaStyleChanged(Sender: TObject);
begin
  dxBarFontAlignLeft.OnClick := nil;
  dxBarFontAlignRight.OnClick := nil;
  dxBarFontAlignCenter.OnClick := nil;
  try
    case RStyle.ParaStyles[RView.CurParaStyleNo].Alignment of
      rvaLeft:
        dxBarFontAlignLeft.Down := True;
      rvaRight:
        dxBarFontAlignRight.Down := True;
      rvaCenter:
        dxBarFontAlignCenter.Down := True;
    end;
  finally
    dxBarFontAlignLeft.OnClick := dxBarFontAlignLeftClick;
    dxBarFontAlignRight.OnClick := dxBarFontAlignRightClick;
    dxBarFontAlignCenter.OnClick := dxBarFontAlignCenterClick;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewCurTextStyleChanged(Sender: TObject);
var
  aFontInfo: TFontInfo;
begin
  dxBarFontSize.OnChange := nil;
  dxBarFontName.OnChange := nil;
  try
    aFontInfo := RStyle.TextStyles[RView.CurTextStyleNo];
    dxBarFontSize.Text := IntToStr( aFontInfo.Size );
    dxBarFontName.ItemIndex := dxBarFontName.Items.IndexOf( aFontInfo.FontName );
    dxBarFontBold.Down := ( fsBold in aFontInfo.Style );
    dxBarFontItalic.Down := ( fsItalic in aFontInfo.Style );
    dxBarFontUnderline.Down := ( fsUnderline in aFontInfo.Style );
  finally
    dxBarFontSize.OnChange := dxBarFontSizeChange;
    dxBarFontName.OnChange := dxBarFontNameChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ( ssCtrl in Shift ) and ( Chr( Key ) = 'A' ) then
    try
      RView.SelectAll;
    except
      {}
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewParaStyleConversion(Sender: TCustomRichViewEdit;
  StyleNo, UserData: Integer; AppliedToText: Boolean; var NewStyleNo: Integer);
var
  aParaInfo: TParaInfo;
begin
  aParaInfo := TParaInfo.Create(nil);
  try
    aParaInfo.Assign( RStyle.ParaStyles[StyleNo] );
    case UserData of
      PARA_ALIGNMENT_LEFT:
        aParaInfo.Alignment := rvaLeft;
      PARA_ALIGNMENT_CENTER:
        aParaInfo.Alignment := rvaCenter;
      PARA_ALIGNMENT_RIGHT:
        aParaInfo.Alignment := rvaRight;
    end;
    NewStyleNo := RStyle.ParaStyles.FindSuchStyle(StyleNo, aParaInfo, RVAllParaInfoProperties );
    if NewStyleNo =-1 then
    begin
      RStyle.ParaStyles.Add;
      NewStyleNo := RStyle.ParaStyles.Count - 1;
      RStyle.ParaStyles[NewStyleNo].Assign( aParaInfo );
      RStyle.ParaStyles[NewStyleNo].Standard := False;
    end;
  finally
    aParaInfo.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.RViewStyleConversion(Sender: TCustomRichViewEdit; StyleNo,
  UserData: Integer; AppliedToText: Boolean; var NewStyleNo: Integer);
var
  FontInfo: TFontInfo;
begin
  FontInfo := TFontInfo.Create(nil);
  try
    FontInfo.Charset := DEFAULT_CHARSET;
    FontInfo.Unicode := True;
    FontInfo.Assign(RStyle.TextStyles[StyleNo]);
    case UserData of
      TEXT_BOLD:
        if dxBarFontBold.Down then
          FontInfo.Style := FontInfo.Style+[fsBold]
        else
          FontInfo.Style := FontInfo.Style-[fsBold];
      TEXT_ITALIC:
        if dxBarFontItalic.Down then
          FontInfo.Style := FontInfo.Style+[fsItalic]
        else
          FontInfo.Style := FontInfo.Style-[fsItalic];
      TEXT_UNDERLINE:
        if dxBarFontUnderline.Down then
          FontInfo.Style := FontInfo.Style+[fsUnderline]
        else
          FontInfo.Style := FontInfo.Style-[fsUnderline];
      TEXT_APPLYFONTNAME:
        FontInfo.FontName := FFontName;
      TEXT_APPLYFONTSIZE:
        FontInfo.Size     := FFontSize;
      TEXT_APPLYFONT:
        FontInfo.Assign( FontDialog.Font );
      TEXT_COLOR:
        FontInfo.Color := FColorPickerController.Color;
      TEXT_BACKCOLOR:
        FontInfo.BackColor := FColorPickerController.Color;
    end;
    NewStyleNo := RStyle.TextStyles.FindSuchStyle( StyleNo, FontInfo, RVAllFontInfoProperties );
    if ( NewStyleNo = -1 ) then
    begin
      RStyle.TextStyles.Add;
      NewStyleNo := ( RStyle.TextStyles.Count - 1 );
      RStyle.TextStyles[NewStyleNo].Assign( FontInfo );
      RStyle.TextStyles[NewStyleNo].Standard := False;
    end;
  finally
    FontInfo.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.FontColorChanged(Sender: TObject);
begin
  AssignFontColorGlyph;
  dxBarFontColorClick( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontNameChange(Sender: TObject);
begin
  if ( dxBarFontName.ItemIndex >=0 ) then
  begin
    FFontName := dxBarFontName.Items[dxBarFontName.ItemIndex];
    RView.ApplyStyleConversion( TEXT_APPLYFONTNAME );
  end;
  if Visible then
    RView.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontSizeChange(Sender: TObject);
begin
  if ( dxBarFontSize.Text <> EmptyStr ) then
  begin
    FFontSize := StrToIntDef( dxBarFontSize.Text, 10 );
    RView.ApplyStyleConversion( TEXT_APPLYFONTSIZE );
  end;
  if Visible then
    RView.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.Bar2CaptionButtons0Click(Sender: TObject);
begin
  FontDialog.Font.Assign( RStyle.TextStyles[RView.CurTextStyleNo] );
  if FontDialog.Execute then RView.ApplyStyleConversion( TEXT_APPLYFONT );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontColorClick(Sender: TObject);
begin
  RView.ApplyStyleConversion( TEXT_COLOR );
  if Visible then
    RView.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontBoldClick(Sender: TObject);
begin
  RView.ApplyStyleConversion( TEXT_BOLD );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontItalicClick(Sender: TObject);
begin
  RView.ApplyStyleConversion( TEXT_ITALIC );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontUnderlineClick(Sender: TObject);
begin
  RView.ApplyStyleConversion( TEXT_UNDERLINE );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarViewUsersClick(Sender: TObject);
begin
  if dxBarViewUsers.Down then
    MsgPage.ActivePageIndex := 1
  else
    MsgPage.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontAlignLeftClick(Sender: TObject);
begin
  RView.ApplyParaStyleConversion( PARA_ALIGNMENT_LEFT );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontAlignCenterClick(Sender: TObject);
begin
  RView.ApplyParaStyleConversion( PARA_ALIGNMENT_CENTER );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontAlignRightClick(Sender: TObject);
begin
  RView.ApplyParaStyleConversion( PARA_ALIGNMENT_RIGHT );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarFontBulletsClick(Sender: TObject);
begin
  RView.ApplyListStyle( 0, 0, 1, False, False );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarButtonOpenClick(Sender: TObject);
var
  aOpenResult: Boolean;
  aCurTextStyleNo, aCurParaStyleNo: Integer;
begin
  if OpenDialog.Execute then
  begin
    aOpenResult := True;
    Screen.Cursor := crHourGlass;
    try
      aCurTextStyleNo := RView.CurTextStyleNo;
      aCurParaStyleNo := RView.CurParaStyleNo;
      RView.Clear;
      RView.CurTextStyleNo := aCurTextStyleNo;
      RView.CurParaStyleNo := aCurParaStyleNo;
      case OpenDialog.FilterIndex of
        1: aOpenResult := RView.LoadRTF( OpenDialog.FileName );
        2: aOpenResult := RView.LoadText( OpenDialog.FileName, aCurTextStyleNo, aCurParaStyleNo, False );
      end;
      if not aOpenResult then
      begin
        Application.MessageBox( '無法開啟檔案。', '錯誤', MB_OK + MB_ICONERROR );
        Exit;
      end;
      RView.Format;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmSendMsg.DoSendMsg: Boolean;

  { --------------------------------------------- }

  function RTFToBinary: Binary;
  begin
    Result := Binary.Create;
    RView.SaveRTFToStream( Result, False );
  end;

  { --------------------------------------------- }

  function CreateAndInitMsgInfo: TMsgInfo;
  begin
    Result := TMsgInfo.Create;
    Result.CompCode := ROClientModule.LoginInfo.CompCode;
    Result.CompName := ROClientModule.LoginInfo.CompName;
    Result.MsgPriority := '0';
    if ( dxBarPriority.Down ) then
      Result.MsgPriority := '1';
    Result.MsgSubject := txtMsgSubject.Text;
    Result.MsgTime := EmptyStr;
    Result.MsgReply := '0';
    Result.MsgSenderId := ROClientModule.LoginInfo.UserId;
    Result.MsgSenderName := ROClientModule.LoginInfo.UserName;
    Result.MsgSenderWorkClass := ROClientModule.LoginInfo.WorkClass;
    Result.MsgSenderWorkName := ROClientModule.LoginInfo.GroupName;
  end;

  { --------------------------------------------- }

var
  aMsgBin: Binary;
  aMsgInfo: TMsgInfo;
  aErrMsg: string;
begin
  { Free in ROClientModule }
  aMsgBin := RTFToBinary;
  { Free in ROClientModule }
  aMsgInfo := CreateAndInitMsgInfo;
  if not ROClientModule.SendMsg( BfRecver, aMsgBin, aMsgInfo, aErrMsg ) then
  begin
    aErrMsg := Format( '無法傳送訊息, 原因:%s。', [aErrMsg] );
    Application.MessageBox( PChar( aErrMsg ), '錯誤', MB_OK + MB_ICONERROR );
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.dxBarSendClick(Sender: TObject);
begin
  if CheckRViewIsEmpty then Exit;
  dxBarSend.OnClick := nil;
  try
    Screen.Cursor := crHourGlass;
    try
      if ( FMsgFormType = 0 ) then
      begin
        if DoSendMsg then
        begin
          FAlreadySend := True;
          Self.Close;
        end;
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    dxBarSend.OnClick := dxBarSendClick;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmSendMsg.GetMsg: Boolean;
var
  aBinary: Binary;
begin
  Result := False;
  Delay( 300 );
  aBinary := ROClientModule.GetMsg( FMsgInfo );
  if Assigned( aBinary ) then
  begin
    RView.Clear;
    RView.LoadRTFFromStream( aBinary );
    aBinary.Free;
    RView.Format;
    Result := True;
  end;
  if ( Result ) then
  begin
    Label1.Caption := Format( '%s(%s) - %s    %s', [FMsgInfo.MsgSenderName,
      FMsgInfo.MsgSenderWorkName, FMsgInfo.CompName, FMsgInfo.MsgTime] );
    Label5.Caption := FMsgInfo.MsgSubject;
    if ( FMsgInfo.MsgPriority = '1' ) then
      StyleModule.SmallImages.GetIcon( StyleModule.GetMsgItemsStateIndex( FMsgInfo.MsgPriority ), Image1.Picture.Icon )
    else
      Label3.Caption := '重要性: 一般';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmSendMsg.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
    if ( GetMsg ) then
    begin
      DataClientModule.UpdateMsgItem( FMsgNode );
      ROClientModule.SetMsgRead( FMsgInfo );
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.

