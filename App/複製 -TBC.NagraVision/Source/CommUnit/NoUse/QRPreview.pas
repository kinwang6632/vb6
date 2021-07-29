{ ---------------------------------------------------------------------------- }
{                                                                              }
{ PC HOME ONLINE Copyright (c) 2002-2003                                       }
{                                                                              }
{ Project: PC home online 網路家庭( EC2 ) ERP Program                          }
{ Unit: Quick Report 自定預覽                                                  }
{ Author: Bug                                                                  }
{ Date: 2003/08/23                                                             }
{                                                                              }
{ ---------------------------------------------------------------------------- }


unit QRPreview;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms, Dialogs, Controls,
  StdCtrls, Qrprntr, Quickrpt, ImgList,
  { EC2 ERP Project Units }
  ApClass, ApConst, DM,
  { Raize CodeSite, http://www.raize.com }
  {$IFDEF DEBUG} CsIntf, {$ENDIF}
  { Developer ExpressBar, http://www.devexpress.com }
  dxBar, dxBarExtItems;

type
  TfrmQRPreview = class(TForm)
    dxBarManager: TdxBarManager;
    dxGotoPage: TdxBarSpinEdit;
    dxFirstPage: TdxBarButton;
    dxPreviousPage: TdxBarButton;
    dxNextPage: TdxBarButton;
    dxLastPage: TdxBarButton;
    dxZoom: TdxBarSpinEdit;
    dxZoomToFit: TdxBarButton;
    dxZoomTo100: TdxBarButton;
    dxZoomToWidth: TdxBarButton;
    dxPrintSetup: TdxBarButton;
    dxPrint: TdxBarButton;
    dxBarDockControl1: TdxBarDockControl;
    dxExit: TdxBarButton;
    QRPreview: TQRPreview;
    dxStatus: TdxBarStatic;
    dxReportProgress: TdxBarProgressItem;
    dxCurrentPrinter: TdxBarStatic;
    dxReportPaperSize: TdxBarStatic;
    ImageList: TImageList;
    SaveDialogExport: TSaveDialog;
    procedure InitialiQRPrinter;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure QRPreviewPageAvailable(Sender: TObject; PageNum: Integer);
    procedure BitBtnExitClick(Sender: TObject);
    procedure CancelReport;
    procedure dxFirstPageClick(Sender: TObject);
    procedure dxPreviousPageClick(Sender: TObject);
    procedure dxNextPageClick(Sender: TObject);
    procedure dxLastPageClick(Sender: TObject);
    procedure dxZoomToFitClick(Sender: TObject);
    procedure dxZoomTo100Click(Sender: TObject);
    procedure dxZoomToWidthClick(Sender: TObject);
    procedure dxPrintSetupClick(Sender: TObject);
    procedure dxPrintClick(Sender: TObject);
    procedure dxExitClick(Sender: TObject);
    procedure QRPreviewProgressUpdate(Sender: TObject; Progress: Integer);
    procedure dxGotoPageChange(Sender: TObject);
    procedure dxZoomCurChange(Sender: TObject);
  private
    { Private declarations }
    FTotalPages: Integer;
    FQuickReport: TQuickRep;
    FParentForm: TCustomForm;
    property TotalPages: Integer read FTotalPages write FTotalPages;
    procedure ShowPreviewCaption(const ACaption: String);
    procedure ShowCurrentPrinterName(const APrinterIndex: Integer);
    procedure ShowCurrentReportPaperSize(Sender: TQuickRep);
    procedure SetMasterReport(Sender: TQuickRep);
    procedure WMSysCommand(var Message: TWMSysCommand); message WM_SYSCOMMAND;
  public
    { Public declarations }
    property MasterReport: TQuickRep read FQuickReport write SetMasterReport;
  end;


var
  frmQRPreview: TfrmQRPreview;

implementation

{$R *.dfm}

uses Main;

var
  
  bPleaseInit: Boolean;

  strFormatPage: String = '頁次: 第%d頁/共%d頁';

  strCurrentPrinter: String = '印表機: %s';

  strPrintProgress: String = '報表列印中 ... %d%s';

  strReportPaperSize: String = '紙張: %s';

  PaperSizeName: array [Default..Custom] of String =
    ( 'Default', 'Letter', 'LetterSmall', 'Tabloid', 'Ledger',
      'Legal', 'Statement', 'Executive', 'A3', 'A4', 'A4Small',
      'A5', 'B4', 'B5', 'Folio', 'Quarto', 'qr10X14', 'qr11X17',
      'Note', 'Env9', 'Env10', 'Env11', 'Env12', 'Env14', 'CSheet',
      'DSheet', 'ESheet', 'Custom' );


{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.InitialiQRPrinter;
begin
  bPleaseInit := False;
  dxReportProgress.Max := Self.TotalPages;
  dxGotoPage.MaxValue := Self.TotalPages;
  dxGotoPage.MinValue := 1;
  dxGotoPage.Value := 1;
  QRPreview.Zoom := 100;
  dxZoomTo100.Down := True;
  dxZoom.Value := QRPreview.Zoom;
  ShowPreviewCaption( QRPreview.QRPrinter.Title );
  ShowCurrentPrinterName( QRPreview.QRPrinter.PrinterIndex );
  ShowCurrentReportPaperSize( MasterReport );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.FormShow(Sender: TObject);
begin
  bPleaseInit := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.QRPreviewPageAvailable(Sender: TObject;
  PageNum: Integer);
const
  AProgressString = '報表產生中 ... 第 %d 頁 ';
begin
  if bPleaseInit then InitialiQRPrinter;
  dxGotoPage.MaxValue := PageNum;
  FTotalPages := QRPreview.QRPrinter.PageCount;
  dxStatus.Caption := Format( AProgressString, [FTotalPages] );
  { Here's one way to show the current status of the report }
  case QRPreview.QRPrinter.Status of
    mpFinished:
       begin
         Screen.Cursor := crDefault;
         dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
           Self.TotalPages] );
         dxReportProgress.Position := 0;
         dxReportProgress.Max := 0;
       end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.FormCreate(Sender: TObject);
begin
  FParentForm := nil;
  FTotalPages := 0;
  if dxBarManager.UseSystemFont then
    dxBarManager.BarByName( 'StatusBar' ).Font := Screen.HintFont;
  Self.WindowState := wsMaximized;
  frmMain.BeginReportPreview;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.BitBtnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.CancelReport;
begin
  if MasterReport.qrprinter.status = mpBusy then
  begin
    if Application.MessageBox( '中止產生此份報表?', '確認', MB_OKCANCEL +
      MB_ICONQUESTION ) = IDOK then
    begin
      QRPreview.QRPrinter.Cancel;
      QRPreview.QrPrinter := nil;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  QRPreview.QRPrinter.Cancel;
  Application.ProcessMessages;
  MasterReport := nil;
  QRPreview.QRPrinter := nil;
  if Assigned( FParentForm ) then FParentForm.Close;
  frmMain.EndReportPreview;
  Action := caFree;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxFirstPageClick(Sender: TObject);
begin
  Application.ProcessMessages;
  dxGotoPage.Value := 1;
  QRPreview.PageNumber := Trunc( dxGotoPage.Value );
  dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
    Self.TotalPages] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxPreviousPageClick(Sender: TObject);
begin
  Application.ProcessMessages;
  if dxGotoPage.Value > 1 then
  begin
    dxGotoPage.Value := dxGotoPage.Value - 1;
    QRPreview.PageNumber := Trunc( dxGotoPage.Value );
  end;
  dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
    Self.TotalPages] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxNextPageClick(Sender: TObject);
begin
  Application.ProcessMessages;
  if dxGotoPage.Value < QRPreview.QRPrinter.PageCount then
  begin
    dxGotoPage.Value := dxGotoPage.Value + 1;
    QRPreview.PageNumber := Trunc( dxGotoPage.Value );
  end;
  dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
    Self.TotalPages] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxLastPageClick(Sender: TObject);
begin
  Application.ProcessMessages;
  dxGotoPage.Value := QRPreview.QRPrinter.PageCount;
  QRPreview.PageNumber := Trunc( dxGotoPage.Value );
  dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
    Self.TotalPages] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxZoomToFitClick(Sender: TObject);
begin
  Application.ProcessMessages;
  QRPreview.ZoomToFit;
  dxZoom.Value := QRPreview.Zoom;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxZoomTo100Click(Sender: TObject);
begin
  Application.ProcessMessages;
  QRPreview.Zoom := 100;
  dxZoom.Value := QRPreview.Zoom;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxZoomToWidthClick(Sender: TObject);
begin
  Application.ProcessMessages;
  QRPreview.ZoomToWidth;
  dxZoom.Value := QRPreview.Zoom;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxPrintSetupClick(Sender: TObject);
begin
  {
     With 2.0j, QuickReport will set the report's tag property to -1 if the
     user cancels the printersetup.  By checking for this value, we can call
     the dxPrint method directly from setup if the users selects OK
  }
  { Just in case you are using an older version }
  MasterReport.Tag := -1;
  MasterReport.PrinterSetup;
  ShowCurrentPrinterName( MasterReport.PrinterSettings.PrinterIndex );
  if MasterReport.Tag = 0 then dxPrintClick( dxPrint );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxPrintClick(Sender: TObject);
begin
  dxReportProgress.Max := 100;
  QRPreview.QRPrinter.Print;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxExitClick(Sender: TObject);
begin
  FParentForm := GetParentForm( MasterReport );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.QRPreviewProgressUpdate(Sender: TObject;
  Progress: Integer);
begin
  dxReportProgress.Position := QRPreview.QRPrinter.PageCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxGotoPageChange(Sender: TObject);
begin
  Application.ProcessMessages;
  QRPreview.PageNumber := Trunc( dxGotoPage.CurValue );
  dxGotoPage.Text := dxGotoPage.CurText;
  dxStatus.Caption := Format( strFormatPage, [QRPreview.PageNumber,
    Self.TotalPages] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.dxZoomCurChange(Sender: TObject);
begin
  Application.ProcessMessages;
  QRPreview.Zoom := Trunc( dxZoom.CurValue );
  dxZoom.Text := dxZoom.CurText;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.ShowPreviewCaption(const ACaption: String);
begin
  Self.Caption := '報表預覽';
  if ACaption <> '' then
    Self.Caption := ACaption + '(預覽)';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.ShowCurrentPrinterName(const APrinterIndex: Integer);
begin
 if APrinterIndex <> -1 then
   dxCurrentPrinter.Caption := Format( strCurrentPrinter,
   [QRPreview.QRPrinter.Printers[APrinterIndex] ] )
 else
   dxCurrentPrinter.Caption := Format( strCurrentPrinter,
   [QRPreview.QRPrinter.Printers[QRPreview.QRPrinter.PrinterIndex]] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.ShowCurrentReportPaperSize(Sender: TQuickRep);
var
  strPaperName: String;
begin
  strPaperName := PaperSizeName[Sender.Page.PaperSize];
  if Sender <> nil then
    dxReportPaperSize.Caption := Format( strReportPaperSize,
    [strPaperName + ' [ ' + FloatToStr( Sender.Page.Length / 10 ) + ' mm * ' +
    FloatToStr( Sender.Page.Width / 10 ) + ' mm ] '] )
 else
  dxReportPaperSize.Caption := Format( strReportPaperSize, [''] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.SetMasterReport(Sender: TQuickRep);
begin
  FQuickReport := Sender;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmQRPreview.WMSysCommand(var Message: TWMSysCommand);
begin
  if ( Message.CmdType and $FFF0 = SC_NEXTWINDOW ) or
     ( Message.CmdType and $FFF0 = SC_PREVWINDOW ) then
    Message.Result := 0
  else
    Inherited;
end;

{ ---------------------------------------------------------------------------- }

end.
