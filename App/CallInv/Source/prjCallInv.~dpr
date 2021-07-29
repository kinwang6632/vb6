program prjCallInv;

uses
  Forms,
  Dialogs,
  frmBrowInvU in 'frmBrowInvU.pas' {frmBrowINV},
  dtmCommU in 'dtmCommU.pas' {dtmComm: TDataModule},
  CommU in 'CommU.pas',
  Encryption_TLB in '..\..\CommUnit\Encryption_TLB.pas',
  frmBrowLogU in 'frmBrowLogU.pas' {frmLog};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdtmComm, dtmComm);
  if ParamCount=5 then
  begin
    if AppCreate then
    begin
      if ParamStr(2) = '1' then
      begin
        Application.CreateForm(TfrmBrowINV,frmBrowINV);
      end else
      begin
        if dtmComm.GetLogData(ParamStr(3),dtmComm.GetCompID) then
        begin
          if dtmComm.adoLog.RecordCount = 0 then
          begin
            MessageDlg( '無任何傳送記錄！',mtInformation,[mbOk],0);
            Exit;
          end;
        end;
        Application.CreateForm(TfrmLog,frmLog);
      end;



    end;
  end else
  begin
    MessageDlg('傳入的參數數量不正確！',mtError,[mbOK],0);
  end;
  Application.Run;
end.
