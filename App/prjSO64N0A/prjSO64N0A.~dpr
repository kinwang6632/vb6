program prjSO64N0A;

uses
  Forms,
  Windows,
  Messages,
  frmcd042 in 'frmcd042.pas' {Form1},
  frmMultiSelectU in 'frmMultiSelectU.pas' {frmMultiSelect},
  print_frmcd042 in 'print_frmcd042.pas' {QuickReport1: TQuickRep},
  frmGiftU in 'frmGiftU.pas' {frmGift},
  frmcd0421 in 'frmcd0421.pas' {Form2},
  frmMultiSelectU2 in 'frmMultiSelectU2.pas' {frmMultiSelect2},
  frmQryCD042U in 'frmQryCD042U.pas' {frmQryCD042};

{$R *.res}

var
  Wnds:Hwnd;

begin
  Wnds := FindWindow( 'TForm1', '促銷活動管理 [SO64N0A]' );
  if ( Wnds = 0 ) or ( FindWindow('TAppBuilder', nil )> 0 ) then
  begin
    Application.Initialize;
    Application.Title := '促銷活動管理 [SO64N0A]';
    Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TfrmMultiSelect2, frmMultiSelect2);
  Application.CreateForm(TfrmQryCD042, frmQryCD042);
  Application.Run;
  end else
  begin
    SendMessage( Wnds, WM_USER + 999, 0, 0 );
    SetForegroundWindow(Wnds);
    Application.Terminate;
  end;
end.
