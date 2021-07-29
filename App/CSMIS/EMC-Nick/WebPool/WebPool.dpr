library WebPool;

uses
  ComServ,
  WebPool_TLB in 'WebPool_TLB.pas',
  cbPoolManager in 'cbPoolManager.pas' {PoolManager: TMtsDataModule} {PoolManager: CoClass},
  Encryption_TLB in 'Encryption_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
