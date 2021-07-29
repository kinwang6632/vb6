unit frmInvA05_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DBCtrls, ComCtrls;

type
  TfrmInvA05_2 = class(TForm)
    edtFilePath: TEdit;
    btnEnter: TButton;
    SpeedButton1: TSpeedButton;
    OpenDialog1: TOpenDialog;
    btnExit: TButton;
    ProgressBar1: TProgressBar;
    rgpInvformat: TRadioGroup;
    Panel2: TPanel;
    Label1: TLabel;
    ListBox1: TListBox;
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private

    { Private declarations }
  public
    { Public declarations }
    sG_UserID : String ;
    sG_CompID : String ;
    procedure setProgress;
    procedure setProgressMax(nI_RecordCount:Integer);
  end;

var
  frmInvA05_2: TfrmInvA05_2;

implementation

uses cbUtilis, dtmMainJU, frmMainU;

{$R *.dfm}

procedure TfrmInvA05_2.SpeedButton1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    edtFilePath.Text := OpenDialog1.FileName;
  end;
end;

procedure TfrmInvA05_2.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_2.btnEnterClick(Sender: TObject);
var
  aPath, aFullName : String ;
  aHowToCreate: Integer ;
begin
  // 0 表示選全部 1 表示不含現場開立
  aHowToCreate := rgpInvformat.ItemIndex ;
  aFullName := edtFilePath.Text;
  aPath := ExtractFileDir( aFullName );

  if ( aFullName = '' ) then
  begin
    WarningMsg( '請輸入或選取檔案路徑。' );
    edtFilePath.SetFocus;
    Exit;
  end;

  if ( aPath = '' ) then
  begin
    WarningMsg( '請輸入檔案存放路徑。' );
    edtFilePath.SetFocus;
    Exit
  end;

  ForceDirectories( aPath );

  if FileExists( aFullName ) then
  begin
    if not ConfirmMsg( '檔案已存在！是否覆蓋舊檔案？' ) then
    begin
      edtFilePath.SetFocus;
      Exit;
    end;
  end;

  ProgressBar1.Min := 0;
  ListBox1.Clear;
  Screen.Cursor := crSQLWait;
  try
    btnEnter.Enabled := False;
    btnExit.Enabled := False;
    try
      dtmMainJ.Trans2Txt( aFullName, aHowToCreate );
    finally
      btnEnter.Enabled := True;
      btnExit.Enabled := True;
    end;
  finally
    Screen.Cursor := crDefault;
  end;  

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_2.setProgress;
begin
  ProgressBar1.StepIt;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_2.setProgressMax(nI_RecordCount: Integer);
begin
  ProgressBar1.Max := nI_RecordCount;
end;

procedure TfrmInvA05_2.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'A05_2','轉文字檔存檔路徑' );
  ListBox1.Clear;
end;

end.
