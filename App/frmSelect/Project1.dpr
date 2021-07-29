program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  frmMultiSelectU2 in 'frmMultiSelectU2.pas' {frmMultiSelect2};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TfrmMultiSelect2, frmMultiSelect2);
  Application.Run;
end.
