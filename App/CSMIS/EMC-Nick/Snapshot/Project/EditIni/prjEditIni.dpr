program prjEditIni;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {Form1},
  Encryption_TLB in '..\..\..\..\Program Files\Borland\Delphi6\Imports\Encryption_TLB.pas',
  DevPower_TLB in '..\..\..\..\Program Files\Borland\Delphi6\Imports\DevPower_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
