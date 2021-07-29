unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxRibbon, dxRibbonForm, dxBar, dxSkinsCore, dxSkinsDefaultPainters,
  cxClasses, cxControls, cxGeometry, cxLookAndFeels, Menus,
  cxLookAndFeelPainters, StdCtrls, cxButtons, dxBarExtItems, ImgList, cxGraphics,
  ExtCtrls, dxSkinscxPCPainter, dxBarStrs, cxStyles, cxCustomData,
  cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData, cxGridLevel,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, dxmdaset, cxGridBandedTableView, cxGridDBBandedTableView,
  cxMemo, cxTextEdit, cxBarEditItem, dxStatusBar,
  dxRibbonStatusBar, cxButtonEdit, cxContainer, cxMaskEdit,
  dxRibbonSkins, dxBarSkinConsts, dxGDIPlusClasses, cxImageComboBox, cxGridStrs,
  OleCtrls, SHDocVw, dxSkinsdxRibbonPainter, dxSkinsdxBarPainter;

type

  TServerConnectState = (sctOnLine, sctOffLine, sctProcDate);
  
  TForm1 = class(TdxRibbonForm)
    dxBarManager1: TdxBarManager;
    RibbonBar: TdxRibbon;
    rbTabAppearance: TdxRibbonTab;
    barQuickAccess: TdxBar;
    barMsg: TdxBar;
    cxLookAndFeelController: TcxLookAndFeelController;
    barSoItem: TdxBarCombo;
    barMsgGroupItem: TdxBarCombo;
    Bar16ImageList: TcxImageList;
    SmallImages: TcxImageList;
    barAppearance: TdxBar;
    barBlue: TdxBarLargeButton;
    barBlack: TdxBarLargeButton;
    barSilver: TdxBarLargeButton;
    LargeImages: TcxImageList;
    GroupBuffer: TdxMemData;
    dsGroup: TDataSource;
    dsMsg: TDataSource;
    MsgBuffer: TdxMemData;
    GroupBufferGroupId: TStringField;
    GroupBufferGroupText: TStringField;
    GroupBufferGroupType: TStringField;
    MsgBufferMsgPriority: TStringField;
    MsgBufferMsgIsRead: TStringField;
    MsgBufferMsgId: TStringField;
    MsgBufferGroupId: TStringField;
    GroupBufferCompCode: TStringField;
    MsgBufferCompCode: TStringField;
    MsgBufferMsgText: TStringField;
    barFilterMsgIsReadItem: TdxBarLargeButton;
    rbTabeView: TdxRibbonTab;
    cxStyleRepository1: TcxStyleRepository;
    cxStyleSkinBlueLargeGroup: TcxStyle;
    EditStyle: TcxEditStyleController;
    ServerPricTimer: TTimer;
    MsgBufferCompName: TStringField;
    barChangeView: TdxBarLargeButton;
    TrayIcon1: TTrayIcon;
    ImageList1: TImageList;
    TrayIconTimer: TTimer;
    dxBarPopupMenu1: TdxBarPopupMenu;
    barShowApp: TdxBarButton;
    barAppExit: TdxBarButton;
    MsgGrid: TcxGrid;
    gvMsgGroup: TcxGridDBBandedTableView;
    gvMsgGroupRecId: TcxGridDBBandedColumn;
    gvMsgGroupCompCode: TcxGridDBBandedColumn;
    gvMsgGroupGroupId: TcxGridDBBandedColumn;
    gvMsgGroupGroupType: TcxGridDBBandedColumn;
    gvMsgGroupGroupText: TcxGridDBBandedColumn;
    gvMsgText: TcxGridDBBandedTableView;
    gvMsgTextRecId: TcxGridDBBandedColumn;
    gvMsgTextCompCode: TcxGridDBBandedColumn;
    gvMsgTextCompName: TcxGridDBBandedColumn;
    gvMsgTextGroupId: TcxGridDBBandedColumn;
    gvMsgTextMsgId: TcxGridDBBandedColumn;
    gvMsgTextMsgPriority: TcxGridDBBandedColumn;
    gvMsgTextMsgIsRead: TcxGridDBBandedColumn;
    gvMsgTextMsgText: TcxGridDBBandedColumn;
    glMsgGroup: TcxGridLevel;
    glMsgText: TcxGridLevel;
    WebBrowser: TWebBrowser;
    ServerStateImage: TImage;
    cxLabel1: TLabel;
    txtSearch: TcxTextEdit;
    btnSearch: TcxButton;
    Shape1: TShape;
    procedure RibbonBarDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure barBlueClick(Sender: TObject);
    procedure barBlackClick(Sender: TObject);
    procedure barSilverClick(Sender: TObject);
    procedure barFilterMsgIsReadItemClick(Sender: TObject);
    procedure MsgBufferFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure txtSearchEnter(Sender: TObject);
    procedure txtSearchExit(Sender: TObject);
    procedure ServerPricTimerTimer(Sender: TObject);
    procedure barSoItemChange(Sender: TObject);
    procedure GroupBufferFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure barMsgGroupItemChange(Sender: TObject);
    procedure barChangeViewClick(Sender: TObject);
    procedure MainPageChange(Sender: TObject);
    procedure gvMsgTextCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure TrayIconTimerTimer(Sender: TObject);
    procedure TrayIcon1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure barShowAppClick(Sender: TObject);
    procedure barAppExitClick(Sender: TObject);
    procedure RibbonBarResize(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
    FWindowShutDownOrLogOff: Boolean;
    FSrvConnectSate: TServerConnectState;
    FMarqueeTemplate: String;
    procedure WMQueryEndSession(var Message: TWMQueryEndSession); message WM_QUERYENDSESSION;
    procedure InitUIStyle;
    procedure InitGroupBuffer;
    procedure InitMsgBuffer;
    procedure DrawServerState(const aSatte: TServerConnectState);
    procedure OnApplicationMessage(var Msg: TMsg; var Handled: Boolean);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}


{ ---------------------------------------------------------------------------- }

procedure TForm1.FormCreate(Sender: TObject);
begin
  FWindowShutDownOrLogOff := True;
  FSrvConnectSate := sctOnLine;
  InitUIStyle;
  GroupBuffer.Active := True;
  MsgBuffer.Active := True;
  InitMsgBuffer;
  InitGroupBuffer;
  FMarqueeTemplate := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'MarqueeTemplate.htm';
  Application.OnMessage := OnApplicationMessage;
  RibbonBarResize( RibbonBar );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( not FWindowShutDownOrLogOff ) then
  begin
    Action := caNone;
    Self.Hide;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormDestroy(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormResize(Sender: TObject);
begin
  RibbonBarResize( RibbonBar );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.WMQueryEndSession(var Message: TWMQueryEndSession);
begin
  FWindowShutDownOrLogOff := True;
  Message.Result := Ord( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.InitUIStyle;
begin

  WebBrowser.Visible := False;
  WebBrowser.Align := alNone;
  WebBrowser.Height := 0;
  MsgGrid.Visible := True;
  MsgGrid.Align := alClient;

  cxSetResourceString(@dxSBAR_SHOWBELOWRIBBON, '快取列顯示在下方' );
  cxSetResourceString(@dxSBAR_SHOWABOVERIBBON, '快取列顯示在上方' );
  cxSetResourceString(@dxSBAR_MINIMIZERIBBON, '自動隱藏功能列' );
  {}
  cxSetResourceString(@scxGridNoDataInfoText, '無符合資料' );
  {}

  EditStyle.Style.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );

  EditStyle.StyleDisabled.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );

  EditStyle.StyleFocused.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );

  EditStyle.StyleHot.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_HOT );

  gvMsgGroup.OptionsView.BandHeaders := False;
  gvMsgText.OptionsView.BandHeaders := False;

  txtSearchExit( txtSearch );

  ServerStateImage.Picture.Icon := TIcon.Create;
  ServerPricTimer.Enabled := True;
  {}
  barSoItem.ItemIndex := 0;
  barMsgGroupItem.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.MainPageChange(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.MsgBufferFilterRecord(DataSet: TDataSet; var Accept: Boolean);
begin
  Accept := True;
  if ( barSoItem.Text <> '全部' ) then
  begin
    Accept := ( DataSet.FieldByName( 'CompName' ).AsString = barSoItem.Text );
  end;
  if ( barFilterMsgIsReadItem.Down ) and ( Accept ) then
    Accept := ( DataSet.FieldByName( 'MsgIsRead' ).AsString = '0' );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.OnApplicationMessage(var Msg: TMsg; var Handled: Boolean);
begin
 if ( Msg.message = WM_RBUTTONDOWN ) or ( Msg.message = WM_RBUTTONUP ) then
    Handled := ( PtInRect( WebBrowser.ClientRect, ScreenToClient( Msg.pt ) ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.GroupBufferFilterRecord(DataSet: TDataSet;
  var Accept: Boolean);
begin
  Accept := True;
  if ( barMsgGroupItem.Text <> '全部' ) then
    Accept := ( DataSet.FieldByName( 'GroupText' ).AsString = barMsgGroupItem.Text );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.InitGroupBuffer;
begin
  gvMsgGroup.ViewData.Expand( False );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.InitMsgBuffer;
var
  aIndex: Integer;
begin
  aIndex := 0;
  MsgBuffer.DisableControls;
  try
    while not MsgBuffer.Eof do
    begin
      Inc( aIndex );
      MsgBuffer.Edit;
      if ( ( aIndex mod 2) = 0 ) then
        MsgBuffer.FieldByName( 'MsgText' ).AsString :=
         'NN43A,NN44B及NN45光機告警, 類比:教育台斷訊 ' +
         '日期：2005-02-22  發佈人：王大明 聯絡人員：王大明　聯絡電話：1234567890 ' +
         '主旨：豐盟機房施工(2/22) ' +
         '影響範圍：影響範圍: ' +
         '主場:豐盟(得易購1,2,3,TOP,5台,富邦購物,東森新聞斷訊) ' +
         '副場:AA,BB,CC,DD(得易購1,2,3,TOP,5台)另外BB還有影響東森新聞跟慈濟大愛二台 ' +
         '影響時間：2005-02-22 00:00:00 ~ 2005-02-22 03:00:00'
      else
        MsgBuffer.FieldByName( 'MsgText' ).AsString :=
         '太陽黑子干擾部份頻道收視異常, 類比國衛台無聲音 ' +
         '日期：2005-03-11 ' +
         '發佈人：李小聰 ' +
         '聯絡人員：  李小聰　聯絡電話：0800019069 ' +
         '主旨：類比國衛台無聲音 ' +
         '影響範圍：影響類比訊號:CH100台收視戶 ' +
         '影響時間：2005-03-11 08:13:00 ~ 2005-03-11 08:16:00 ';
      MsgBuffer.Post;
      MsgBuffer.Next;
    end;
    MsgBuffer.First;
  finally
    MsgBuffer.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RibbonBarDblClick(Sender: TObject);
var
  aPoint: TPoint;
begin
  GetCursorPos( aPoint );
  aPoint := ScreenToClient( aPoint );
  if cxRectPtIn( RibbonBar.ViewInfo.ApplicationButtonImageBounds,
    aPoint.X, aPoint.Y ) then
  begin
    TrayIcon1.Animate := not TrayIcon1.Animate;
    if not TrayIcon1.Animate then TrayIcon1.IconIndex := 0;
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RibbonBarResize(Sender: TObject);
begin
  if WebBrowser.Visible then
    WebBrowser.Navigate( FMarqueeTemplate );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barBlueClick(Sender: TObject);
begin
  RibbonBar.ColorSchemeName := 'Blue';
  cxLookAndFeelController.SkinName := 'MoneyTwins';
  EditStyle.Style.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );
  EditStyle.StyleDisabled.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );
  EditStyle.StyleFocused.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );
  EditStyle.StyleHot.Color := TdxBlueRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_HOT );
  txtSearchExit( txtSearch );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barBlackClick(Sender: TObject);
begin
  RibbonBar.ColorSchemeName := 'Black';
  cxLookAndFeelController.SkinName := 'Caramel';
  {}
  EditStyle.Style.Color := TdxBlackRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );
  EditStyle.StyleDisabled.Color := TdxBlackRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );
  EditStyle.StyleFocused.Color := TdxBlackRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );
  EditStyle.StyleHot.Color := TdxBlackRibbonSkin( RibbonBar.ColorScheme ).GetPartColor(
    DXBAR_EDIT_BACKGROUND, DXBAR_HOT );
  Shape1.Pen.Color := clBlack;  
  txtSearchExit( txtSearch );  
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barSilverClick(Sender: TObject);
begin
  RibbonBar.ColorSchemeName := 'Silver';
  cxLookAndFeelController.SkinName := EmptyStr;
  Shape1.Pen.Color := clGray;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barFilterMsgIsReadItemClick(Sender: TObject);
begin
  gvMsgText.BeginUpdate;
  try
    MsgBuffer.Filtered := False;
    MsgBuffer.Filtered := True;
  finally
   gvMsgText.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.txtSearchEnter(Sender: TObject);
begin
  txtSearch.Clear;
  EditStyle.Style.Font.Color := clWindowText;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.txtSearchExit(Sender: TObject);
begin
  txtSearch.Text := '請輸入搜尋關鍵字';
  if ( barBlue.Down ) then
    EditStyle.Style.Font.Color := $DEC1AB
  else if ( barBlack.Down ) then
    EditStyle.Style.Font.Color := $898989
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.DrawServerState(const aSatte: TServerConnectState);
begin
  case ( aSatte ) of
    sctOnLine:
      Bar16ImageList.GetIcon( 23, ServerStateImage.Picture.Icon );
    sctProcDate:
      Bar16ImageList.GetIcon( 25, ServerStateImage.Picture.Icon );
  else
    Bar16ImageList.GetIcon( 21, ServerStateImage.Picture.Icon );
  end;
  ServerStateImage.Invalidate;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barShowAppClick(Sender: TObject);
begin
  if ( not Self.Visible) then
    Self.Show;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barAppExitClick(Sender: TObject);
begin
  FWindowShutDownOrLogOff := True;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barChangeViewClick(Sender: TObject);
var
  aEnabled: Boolean;
begin
  if barChangeView.Down then
  begin
    MsgGrid.Visible := False;
    MsgGrid.Align := alNone;
    MsgGrid.Height := 0;
    WebBrowser.Visible := True;
    if ( WebBrowser.Align <> alClient ) then WebBrowser.Align := alClient;
   RibbonBarResize( RibbonBar );
   RibbonBar.Visible := False;
  end else
  begin
    WebBrowser.Visible := False;
    WebBrowser.Align := alNone;
    WebBrowser.Height := 0;
    MsgGrid.Visible := True;
    if ( MsgGrid.Align <> alClient ) then MsgGrid.Align := alClient;
       RibbonBar.Visible := True;
  end;
  aEnabled := ( MsgGrid.Visible );
  barFilterMsgIsReadItem.Enabled := aEnabled;
  barSoItem.Enabled := aEnabled;
  barMsgGroupItem.Enabled := aEnabled;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ServerPricTimerTimer(Sender: TObject);
begin
  if ( FSrvConnectSate = sctOffLine ) then
    FSrvConnectSate := sctOnLine
  else if ( FSrvConnectSate = sctOnLine ) then
    FSrvConnectSate := sctProcDate
  else
    FSrvConnectSate := sctOnLine;
  DrawServerState( FSrvConnectSate );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barSoItemChange(Sender: TObject);
begin
  gvMsgText.BeginUpdate;
  try
    MsgBuffer.Filtered := False;
    MsgBuffer.Filtered := True;
  finally
    gvMsgText.EndUpdate;
  end;
  gvMsgGroup.ViewData.Expand( False );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.barMsgGroupItemChange(Sender: TObject);
begin
  gvMsgGroup.BeginUpdate;
  try
    GroupBuffer.Filtered := False;
    GroupBuffer.Filtered := True;
  finally
    gvMsgGroup.EndUpdate;
  end;
  gvMsgGroup.ViewData.Expand( False );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.gvMsgTextCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aItem: TcxGridDBBandedColumn;
  aIsRead: String;
begin
  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
  if ( UpperCase( aItem.DataBinding.FieldName ) = 'MSGTEXT' ) then
  begin
    if ( VarToStrDef( AViewInfo.GridRecord.Values[
      gvMsgTextMsgIsRead.Index], EmptyStr ) <> '1' ) then
    begin
      ACanvas.Font.Color := clRed;
      if ( AViewInfo.Focused or AViewInfo.Selected ) then
        ACanvas.Font.Color := clHighlightText;
      ACanvas.Font.Style := ( ACanvas.Font.Style + [fsBold] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
  barShowApp.Click;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.TrayIcon1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if ( Button = mbRight ) then
  begin
    SetForegroundWindow( Application.Handle );
    Application.ProcessMessages;
    dxBarPopupMenu1.PopupFromCursorPos;
  end;

end;

procedure TForm1.TrayIconTimerTimer(Sender: TObject);
begin
  if not TrayIcon1.Visible then
    TrayIcon1.Visible := True;
  TrayIcon1.ShowBalloonHint;
  TrayIconTimer.Enabled := False;
  TrayIcon1.Animate := True;
end;

{ ---------------------------------------------------------------------------- }

end.
