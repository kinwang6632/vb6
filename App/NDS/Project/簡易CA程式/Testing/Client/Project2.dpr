program Project2;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Project1_TLB in 'C:\Program Files\Borland\Delphi6\Imports\Project1_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
