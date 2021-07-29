program AFA;

uses
  Forms,
  cbMain in 'cbMain.pas' {Main},
  cbDataControler in 'cbDataControler.pas' {DataControler: TDataModule},
  cbFTPControler in 'cbFTPControler.pas' {FTPControler: TDataModule},
  cbFtpThread in 'cbFtpThread.pas',
  Encryption_TLB in 'Encryption_TLB.pas',
  cbDataThread in 'cbDataThread.pas',
  cbClass in 'cbClass.pas',
  cbWatchDogThread in 'cbWatchDogThread.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMain, Main);
  Application.Run;
end.
