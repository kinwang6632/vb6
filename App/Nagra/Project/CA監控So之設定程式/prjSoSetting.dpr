program prjSoSetting;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
