program prjSGSEditIni;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {Form1},
  DevPower_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPower_TLB.pas',
  DevPowerEncrypt_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPowerEncrypt_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
