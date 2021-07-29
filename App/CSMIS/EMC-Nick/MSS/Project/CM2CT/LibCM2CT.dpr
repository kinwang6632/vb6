library LibCM2CT;

uses
  ComServ,
  LibCM2CT_TLB in 'LibCM2CT_TLB.pas',
  cbCM2CT in 'cbCM2CT.pas' {CM2CT: CoClass};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
