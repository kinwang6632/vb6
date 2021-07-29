program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  frmSO18B5U in 'frmSO18B5U.pas' {frmSO18B5},
  frmDbMultiSelectu in '..\Common\frmDbMultiSelectu.pas' {frmDbMultiSelect},
  Ustru in '..\Common\Ustru.pas',
  UdbU in '..\Common\UdbU.pas',
  frmDbSelectu in '..\Common\frmDbSelectu.pas' {frmDbSelect};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
