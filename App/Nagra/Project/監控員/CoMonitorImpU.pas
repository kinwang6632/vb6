// �ɦ�2003/01/20....�o�� function ���٨S���Q����e�ݵ{���Ψ�
unit CoMonitorImpU;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, prjMonitor_TLB, StdVcl;

type
  TMonitor = class(TAutoObject, IMonitor)
  protected
    procedure RunCa(sI_IniFileName: OleVariant); safecall;
    { Protected declarations }
  end;

implementation

uses ComServ, frmMainU;


procedure TMonitor.RunCa(sI_IniFileName: OleVariant);
begin
    frmMain.RunCA(sI_IniFileName);

end;

initialization
  TAutoObjectFactory.Create(ComServer, TMonitor, Class_Monitor,
    ciMultiInstance, tmApartment);
end.
