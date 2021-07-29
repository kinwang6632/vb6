library prjE2C;

uses
  ComServ,
  prjE2C_TLB in 'prjE2C_TLB.pas',
  E2CImp in 'E2CImp.pas' {E2C: CoClass},
  UdateTimeu in 'UdateTimeu.pas',
  Ustru in 'Ustru.pas',
  xmlU in 'xmlU.pas',
  MSXML_TLB in 'MSXML_TLB.pas',
  DevPower_TLB in 'DevPower_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
