program prjSimpleSGS;

uses
  Forms,
  frmSimpleSGSU in 'frmSimpleSGSU.pas' {frmSimpleSGS},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  ConstU in '..\Common\ConstU.pas',
  DevPower_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPower_TLB.pas',
  DevPowerEncrypt_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPowerEncrypt_TLB.pas',
  Ustru in '..\Common\Ustru.pas',
  fraYMDU in '..\Common\fraYMDU.pas' {fraYMD: TFrame},
  UdateTimeu in '..\Common\UdateTimeu.pas',
  xmlU in '..\Common\xmlU.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmSimpleSGS, frmSimpleSGS);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
