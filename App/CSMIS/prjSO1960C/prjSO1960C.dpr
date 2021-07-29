program prjSO1960C;

uses
  Forms,
  Windows,
  Messages,
  cbMain in 'cbMain.pas' {fmSO1960C},
  cbUtilis in '..\..\..\..\CommoUnit\cbUtilis.pas';

{$R *.res}

var
  Wnds: Hwnd;
begin
  Wnds := FindWindow( 'fmSO1960C', 'EMTA/CM資料管理 [SO1960C]' );
  if ( Wnds = 0 ) or ( FindWindow('TAppBuilder', nil ) > 0 ) then
  begin
    Application.Initialize;
    Application.Title := 'EMTA/CM資料管理 [SO1960C]';
    Application.CreateForm(TfmSO1960C, fmSO1960C);
    Application.Run;
  end else
  begin
    SendMessage( Wnds,WM_USER + 999, 0, 0 );
    SetForegroundWindow( Wnds );
    Application.Terminate;
  end;
end.
