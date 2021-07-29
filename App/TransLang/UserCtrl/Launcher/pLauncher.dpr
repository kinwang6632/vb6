program pLauncher;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  utlUserInoU in '..\Common\utlUserInoU.pas',
  utlUCU in '..\Common\utlUCU.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
