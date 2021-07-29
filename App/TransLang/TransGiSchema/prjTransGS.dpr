program prjTransGS;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
