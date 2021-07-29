program prjEditIni;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {Form1},
  Encryption_TLB in 'C:\Program Files\Borland\Delphi6\Imports\Encryption_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
