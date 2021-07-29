program prjEditIni;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {Form1},
  Encryption_TLB in 'Encryption_TLB.pas',
  DevPower_TLB in 'DevPower_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
