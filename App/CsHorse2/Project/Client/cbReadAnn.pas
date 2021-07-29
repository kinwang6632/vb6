unit cbReadAnn;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, cbSyncObjs, DB,
{ App: }
  cbStyleModule, cbClientClass, cbDataClientModule, cbHrHelper,
{ Developer Express: }
  cxClasses, cxControls, cxContainer,
  cxEdit, cxGroupBox, cxPC,
  dxRibbon, dxSkinsdxRibbonPainter, dxRibbonForm, dxSkinsCore,
  dxSkinsDefaultPainters, cxLookAndFeelPainters, dxSkinscxPCPainter,
  dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkiniMaginary, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Silver,
{ ODAC: }
  MemDS, VirtualTable,
{ RichView }
  RVScroll, RichView, RVStyle, dxSkinSummer2008;

type
  TfmReadAnn = class(THrClientForm)
    RibbonTab1: TdxRibbonTab;
    Ribbon: TdxRibbon;
    HeaderPage: TcxPageControl;
    TabSO021: TcxTabSheet;
    TabSO022: TcxTabSheet;
    TabCD042: TcxTabSheet;
    cxGroupBox1: TcxGroupBox;
    RView: TRichView;
    cxGroupBox2: TcxGroupBox;
    cxGroupBox3: TcxGroupBox;
    Timer1: TTimer;
    lbGistTime: TLabel;
    lbGistMan: TLabel;
    lbFaultTime: TLabel;
    lbFixTime: TLabel;
    lbFaultReason: TLabel;
    lbAcceptMan: TLabel;
    lbActDate: TLabel;
    lbDescription: TLabel;
    RVStyle: TRVStyle;
    lbGist: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FAnnBuffer: TVirtualTable;
    FSubBuffer: TVirtualTable;
    FAnnSyncMutex: TMutex;
    procedure DoShowAnn;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent);
    property AnnSyncMutex: TMutex read FAnnSyncMutex write FAnnSyncMutex;
    procedure SetBuffer(AMaster, ADetail: TDataSet);
  end;

var
  fmReadAnn: TfmReadAnn;

implementation

uses cbUtilis, cbMain;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

{ TfmReadAnn }

constructor TfmReadAnn.Create(AOwner: TComponent);
begin
  inherited Create( AOwner, ftAnn );
  FAnnBuffer := TVirtualTable.Create( nil );
  FSubBuffer := TVirtualTable.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.FormCreate(Sender: TObject);
begin
  Ribbon.ViewInfo.ForceHide := True;
  HeaderPage.HideTabs := True;
  ChangeColorSchema( StyleModule.ColorName );
  DataClientModule.AddForm( Self );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DataClientModule.RemoveForm( Self );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.FormDestroy(Sender: TObject);
begin
  FAnnBuffer.Free;
  FSubBuffer.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.SetBuffer(AMaster, ADetail: TDataSet);
begin
  FAnnBuffer.Name := AMaster.Name;
  if Assigned( ADetail ) then FSubBuffer.Name := ADetail.Name;
  if ( UpperCase( FAnnBuffer.Name ) = 'BFSO021' ) then
  begin
    HeaderPage.ActivePageIndex := 0;
    lbGistTime.Caption := EmptyStr;
    lbGistMan.Caption := EmptyStr;
    lbGist.Caption := EmptyStr;
    TBufferHelper.CreateFieldDefs( biSO021, TVirtualTable( FAnnBuffer ) );
    FAnnSyncMutex.Acquire;
    try
      FAnnBuffer.Append;
      FAnnBuffer.FieldByName( 'BoardTime' ).AsString := AMaster.FieldByName( 'BoardTime' ).AsString;
      FAnnBuffer.FieldByName( 'BoardEn' ).AsString := AMaster.FieldByName( 'BoardEn' ).AsString;
      FAnnBuffer.FieldByName( 'Subject' ).AsString := AMaster.FieldByName( 'Subject' ).AsString;
      FAnnBuffer.FieldByName( 'Content' ).AsString := AMaster.FieldByName( 'Content' ).AsString;
      FAnnBuffer.Post;
    finally
      FAnnSyncMutex.Release;
    end;
  end else
  if ( UpperCase( FAnnBuffer.Name ) = 'BFSO022' ) then
  begin
    HeaderPage.ActivePageIndex := 1;
    lbFaultTime.Caption := EmptyStr;
    lbFixTime.Caption := EmptyStr;
    lbFaultReason.Caption := EmptyStr;
    lbAcceptMan.Caption := EmptyStr;
    TBufferHelper.CreateFieldDefs( biSO022, TVirtualTable( FAnnBuffer ) );
    if Assigned( FSubBuffer ) then
      TBufferHelper.CreateFieldDefs( biSO023, TVirtualTable( FSubBuffer ) );
    FAnnSyncMutex.Acquire;
    try
      FAnnBuffer.Append;
      FAnnBuffer.FieldByName( 'ErrorTime' ).AsString := AMaster.FieldByName( 'ErrorTime' ).AsString;
      FAnnBuffer.FieldByName( 'EndTime' ).AsString := AMaster.FieldByName( 'EndTime' ).AsString;
      FAnnBuffer.FieldByName( 'MfName' ).AsString := AMaster.FieldByName( 'MfName' ).AsString;
      FAnnBuffer.FieldByName( 'AcceptName' ).AsString := AMaster.FieldByName( 'AcceptName' ).AsString;
      FAnnBuffer.FieldByName( 'Description' ).AsString := AMaster.FieldByName( 'Description' ).AsString;
      FAnnBuffer.Post;
      {}
      FSubBuffer.Append;
      FSubBuffer.FieldByName( 'Address' ).AsString := ADetail.FieldByName( 'Address' ).AsString;
      FSubBuffer.Post;
    finally
      FAnnSyncMutex.Release;
    end;
  end else
  if ( UpperCase( FAnnBuffer.Name ) = 'BFCD042' ) then
  begin
    HeaderPage.ActivePageIndex := 2;
    lbActDate.Caption := EmptyStr;
    lbDescription.Caption := EmptyStr;
    TBufferHelper.CreateFieldDefs( biCD042, TVirtualTable( FAnnBuffer ) );
    FAnnSyncMutex.Acquire;
    try
      FAnnBuffer.Append;
      FAnnBuffer.FieldByName( 'ActDateStToEd' ).AsString := AMaster.FieldByName( 'ActDateStToEd' ).AsString;
      FAnnBuffer.FieldByName( 'Description' ).AsString := AMaster.FieldByName( 'Description' ).AsString;
      FAnnBuffer.FieldByName( 'Note' ).AsString := AMaster.FieldByName( 'Note' ).AsString;
      FAnnBuffer.Post;
    finally
      FAnnSyncMutex.Release;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.FormShow(Sender: TObject);
begin
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Screen.Cursor := crAppStart;
  try
    DoShowAnn;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmReadAnn.DoShowAnn;
var
  aContentText: String;
begin
  if ( UpperCase( FAnnBuffer.Name ) = 'BFSO021' ) then
  begin
    lbGistTime.Caption := FAnnBuffer.FieldByName( 'BoardTime' ).AsString;
    lbGistMan.Caption := FAnnBuffer.FieldByName( 'BoardEn' ).AsString;
    lbGist.Caption := FAnnBuffer.FieldByName( 'Subject' ).AsString;
    aContentText := FAnnBuffer.FieldByName( 'Content' ).AsString;
    Self.Caption := '¤½§GÄæ';
  end else
  if ( UpperCase( FAnnBuffer.Name ) = 'BFSO022' ) then
  begin
    lbFaultTime.Caption := FAnnBuffer.FieldByName( 'ErrorTime' ).AsString;
    lbFixTime.Caption := FAnnBuffer.FieldByName( 'EndTime' ).AsString;
    lbFaultReason.Caption := FAnnBuffer.FieldByName( 'MfName' ).AsString;
    lbAcceptMan.Caption := FAnnBuffer.FieldByName( 'AcceptName' ).AsString;
    aContentText := FAnnBuffer.FieldByName( 'Description' ).AsString;
    aContentText := ( aContentText + #13#10 + #13#10 );
    aContentText := ( aContentText + FSubBuffer.FieldByName( 'Address' ).AsString );
    Self.Caption := '¬G»Ù°Ï°ì';
  end else
  if ( UpperCase( FAnnBuffer.Name ) = 'BFCD042' ) then
  begin
    lbActDate.Caption := FAnnBuffer.FieldByName( 'ActDateStToEd' ).AsString;
    lbDescription.Caption := FAnnBuffer.FieldByName( 'Description' ).AsString;
    aContentText := FAnnBuffer.FieldByName( 'Note' ).AsString;
    Self.Caption := '«P®×';
  end;
  RView.AddNL( aContentText, -1, -1 );
  RView.Format;
end;

{ ---------------------------------------------------------------------------- }

end.
