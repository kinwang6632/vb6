program prjSnapShot;

uses
  Forms,
  frmSnapShotU in 'frmSnapShotU.pas' {frmSnapShot},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  fraChineseYMDU in '..\Common\fraChineseYMDU.pas' {fraChineseYMD: TFrame},
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  UCommonU in '..\Common\UCommonU.pas',
  UObjectu in '..\Common\UObjectu.pas',
  DevPower_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPower_TLB.pas',
  DevPowerEncrypt_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPowerEncrypt_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmSnapShot, frmSnapShot);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
