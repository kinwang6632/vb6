program projLancelot;

{%File '..\..\Bin\UC.dll'}

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  dtmBasicU in '..\..\Common\dtmBasicU.pas' {dtmBasic: TDataModule},
  dtmTL0020U in 'dtmTL0020U.pas' {dtmTL0020: TDataModule},
  frmTL0020U in 'frmTL0020U.pas' {frmTL0020},
  Ustru in '..\..\Common\Ustru.pas',
  UCommonU in '..\..\Common\UCommonU.pas',
  frmTL0030U in 'frmTL0030U.pas' {frmTL0030},
  dtmTL0030U in 'dtmTL0030U.pas' {dtmTL0030: TDataModule},
  dtmCommonU in '..\..\Common\dtmCommonU.pas' {dtmCommon: TDataModule},
  UObjectu in '..\..\Common\UObjectu.pas',
  utlUserInoU in '..\..\UserCtrl\Common\utlUserInoU.pas',
  utlUCU in '..\..\UserCtrl\Common\utlUCU.pas',
  dllUCU in '..\..\UserCtrl\Common\dllUCU.pas',
  frmTL0040U in 'frmTL0040U.pas' {frmTL0040},
  dtmTL0040U in 'dtmTL0040U.pas' {dtmTL0040: TDataModule};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TdtmCommon, dtmCommon);
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmTL0040, dtmTL0040);
  Application.Run;
end.
