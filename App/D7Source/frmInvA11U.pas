unit frmInvA11U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons;

type
  TfrmInvA11 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    GroupBox1: TGroupBox;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInvA11: TfrmInvA11;

implementation

uses frmMainU, dtmMainU, dtmReportModule;

{$R *.dfm}

procedure TfrmInvA11.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA11.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'A0B', '電子計算機發票套印調整' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA11.btnOKClick(Sender: TObject);
var
  aPath: String;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + IncludeTrailingPathDelimiter( REPORT_FOLDER );
  if RadioButton1.Checked then
    aPath := IncludeTrailingPathDelimiter( aPath ) + 'FrptInvA05_1.fr3'
  else if RadioButton2.Checked then
    aPath := IncludeTrailingPathDelimiter( aPath ) + 'FrptInvA05_2.fr3'
  else if RadioButton3.Checked then
    aPath := IncludeTrailingPathDelimiter( aPath ) + 'FrptInvA05_3.fr3'
  else
    Exit;
  dtmReport.AddA05DataField( dtmMain.GetAutoCreateNum );  
  dtmReport.frxMainReport.LoadFromFile( aPath );
  dtmReport.frxMainReport.DesignReport;
end;

{ ---------------------------------------------------------------------------- }

end.
