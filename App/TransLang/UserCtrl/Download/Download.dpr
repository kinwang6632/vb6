program Download;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain};

{$R *.RES}

begin
  Application.Initialize;
  Application.ShowMainForm := False;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
