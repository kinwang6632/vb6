program CARecycle;

uses
  Forms,
  cbMain in 'cbMain.pas' {fmMain},
  cbSoModule in 'cbSoModule.pas' {SoDataModule: TDataModule},
  cbStyleModule in 'cbStyleModule.pas' {StyleModule: TDataModule},
  cbLogin in 'cbLogin.pas' {fmLogin},
  cbAppClass in 'cbAppClass.pas',
  cbUtilis in '..\..\..\CommoUnit\cbUtilis.pas',
  Encryption_TLB in '..\..\..\CommoUnit\Encryption_TLB.pas',
  cbOption in 'cbOption.pas' {fmOption};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TStyleModule, StyleModule);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
