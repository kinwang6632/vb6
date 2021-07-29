program prjSGS;

uses
  Forms,
  frmLoginU in 'frmLoginU.pas' {frmLogin},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule},
  frmMainU in 'frmMainU.pas' {frmMain},
  frmSysParamU in 'frmSysParamU.pas' {frmSysParam},
  ConstU in '..\Common\ConstU.pas',
  DevPower_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPower_TLB.pas',
  DevPowerEncrypt_TLB in 'C:\Program Files\Borland\Delphi6\Imports\DevPowerEncrypt_TLB.pas',
  frmDbInfoU in 'frmDbInfoU.pas' {frmDbInfo},
  Ustru in '..\Common\Ustru.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  xmlU in '..\Common\xmlU.pas',
  Uotheru in '..\Common\UotherU.pas',
  UmsgCodeu in '..\Common\UmsgCodeu.pas',
  frmNightRunLogU in 'frmNightRunLogU.pas' {frmNightRunLog},
  fraYMDU in '..\Common\fraYMDU.pas' {fraYMD: TFrame},
  THttpThreadU in 'THttpThreadU.pas',
  TNightRunThreadU in 'TNightRunThreadU.pas',
  frmQueryLogU in 'frmQueryLogU.pas' {frmQueryLog},
  frmLoadDataU in 'frmLoadDataU.pas' {frmLoadData};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmMain, dtmMain);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
