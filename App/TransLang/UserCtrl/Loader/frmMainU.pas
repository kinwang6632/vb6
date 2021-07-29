unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

type
  TfrmMain = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    function RunShell(ExeName, Param: string; Wait: Boolean): Boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses frmSetPathU, IniFiles, FileCtrl;

{$R *.DFM}


procedure TfrmMain.FormCreate(Sender: TObject);
var
  Count: Integer;
  Inif: TIniFile;
  IniName, UCPath, DownPath: string;
begin
  IniName := ExtractFilePath(Application.ExeName) + 'UC.INI';
  Inif := TIniFile.Create(IniName);
  UCPath := Inif.ReadString('SYSTEM', 'UCPATH', '');
  DownPath := Inif.ReadString('SYSTEM', 'DOWNLOAD', '');
  frmSetPath := TfrmSetPath.Create(nil);
  Count := 0;

  repeat
    if Count <> 0 then
    begin
      if not DirectoryExists(UCPath) then
        MessageDlg('系統檔案路徑不存在', mtError, [mbOk], 0);
      if not DirectoryExists(DownPath) then
        MessageDlg('軟體儲存路徑不存在', mtError, [mbOk], 0);
    end;
    frmSetPath.DownPath := DownPath;
    frmSetPath.UCPath := UCPath;
    if frmSetPath.ShowModal = mrCancel then
    begin
      Inif.Free;
      frmSetPath.Free;
      Application.Terminate;
      Exit;
    end;
    UCPath := frmSetPath.UCPath;
    DownPath := frmSetPath.DownPath;
    Count := Count + 1;
  until DirectoryExists(UCPath) and DirectoryExists(DownPath);

  frmSetPath.Free;
  Inif.WriteString('SYSTEM', 'UCPATH', UCPath);
  Inif.WriteString('SYSTEM', 'DOWNLOAD', DownPath);
  Inif.Free;

  RunShell('Download.exe', 'pLauncher.exe', True);
  WinExec('pLauncher.exe', SW_SHOWNORMAL);
  Application.Terminate;
end;

function TfrmMain.RunShell(ExeName, Param: string; Wait: Boolean): Boolean;
var
  pi: TProcessInformation;
  si: TStartupInfo;
  bRet: Boolean;
begin

  Result := False;

  FillChar(si, SizeOf(TStartupInfo), 0);
  si.cb := SizeOf(TStartupInfo);
  si.wShowWindow := SW_NORMAL;

  bRet := CreateProcess(nil, PChar(ExeName + ' ' + Param), nil, nil, False,
    0, nil, nil, si, pi);

  if not bRet then Exit;

  if not Wait then
  begin
    Result := True;
    Exit;
  end;

  { wait until the process terminate}
  WaitForSingleObject(pi.hProcess, INFINITE);

  Result := True ;
end;


end.
