program Loader;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSetPathU in 'frmSetPathU.pas' {frmSetPath};

{$R *.RES}

begin
  Application.Initialize;
  Application.ShowMainForm := False;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
