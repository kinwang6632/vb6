unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, ImgList, utlUserInoU, utlUCU;

type

  TfrmMain = class(TForm)
    pnlCaption: TPanel;
    stbInfo: TStatusBar;
    pnlFunc: TPanel;
    lblUserName: TLabel;
    lblTime: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Timer1: TTimer;
    Label4: TLabel;
    lblDate: TLabel;
    trvApp: TTreeView;
    Splitter1: TSplitter;
    lsvFunc: TListView;
    ImageList1: TImageList;
    Panel1: TPanel;
    trvUserGrp: TTreeView;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure trvAppChange(Sender: TObject; Node: TTreeNode);
    procedure lsvFuncDblClick(Sender: TObject);
    procedure lsvFuncKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
    procedure pnlCaptionDblClick(Sender: TObject);
  private
    { Private declarations }
    FUserInfoShm: TSharedMem;
    FCurUser: TUser;
    procedure ShowSystemInfo;
    function RunShell(ExeName, Param: string; WithUI: Boolean): Boolean;
    procedure DownloadFiles(FileNames: string);
    procedure DownloadFuncPackage(Func: TFunc);
    procedure RegisterFuncDataFiles(Func: TFunc);
    procedure Clear;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses IniFiles;

{$R *.DFM}

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  FUserInfoShm := TSharedMem.Create('UCINFO', SizeOf(UserInfo));
  DownloadFiles('Loader.exe Login.exe Uc.dll');
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
  UserId, UserPass, UserName, UserGrp: string;
begin
  Width := Screen.Width;
  Height := Screen.Height;

  GetUserInfo(UserId, UserName, UserPass, UserGrp);
  if UserId = '' then {user fisrt login or log out}
  begin
    RunShell('Login.exe', '', False);
    GetUserInfo(UserId, UserName, UserPass, UserGrp);
    if UserId = '' then
    begin
      Application.Terminate;
      Exit;
    end;
  end;

  if not Assigned(UCMgr) then
    UCMgr := TUCMgr.Create(trvApp.Items, trvUserGrp.Items);
  FCurUser := UCMgr.HasUser(UserId, UserPass);
  ShowSystemInfo;
end;

procedure TfrmMain.Clear;
begin
  trvApp.Items.Clear;
  trvUserGrp.Items.Clear;
  lsvFunc.Items.Clear;
end;

procedure TfrmMain.ShowSystemInfo;
var
  Inif: TIniFile;
begin
  if Assigned(FCurUser) then lblUserName.Caption := FCurUser.Desc;
  Inif := TIniFile.Create(ExtractFilePath(Application.ExeName)+'\UC.ini');
  pnlCaption.Caption := Inif.ReadString('SYSTEM', 'NAME', '');
  stbInfo.SimpleText := Inif.ReadString('SYSTEM', 'CONTACT', '');
  Inif.Free;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
  lblDate.Caption := DateToStr(Date);
  lblTime.Caption := TimeToStr(Time);
end;

function TfrmMain.RunShell(ExeName, Param: string; WithUI: Boolean): Boolean;
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

  if WithUI then Hide;

  { wait until the process terminate}
  WaitForSingleObject(pi.hProcess, INFINITE);

  if WithUI then Show;

  Result := True ;
end;

procedure TfrmMain.DownloadFiles(FileNames: string);
begin
  RunShell('Download', FileNames, False);
end;

procedure TfrmMain.DownloadFuncPackage(Func: TFunc);
var
  ii: Integer;
  FileNames: string;
begin
  FileNames := Func.ExeName;
  for ii:=0 to Func.DataFileCount-1 do
  begin
    if FileNames <> '' then FileNames := FileNames + ' ';
    FileNames := FileNames + Func.DataFiles[ii].FileName;
  end;
  DownloadFiles(FileNames);
end;

procedure TfrmMain.RegisterFuncDataFiles(Func: TFunc);
var
  ii: Integer;
begin
  for ii:=0 to Func.DataFileCount-1 do
  begin
    if Func.DataFiles[ii].AutoReg then
    begin
      if Pos('.DLL', UpperCase(Func.DataFiles[ii].FileName)) > 0 then
        RunShell('Regsvr32', '/s ' + Func.DataFiles[ii].FileName, False)
      else if Pos('.EXE', UpperCase(Func.DataFiles[ii].FileName)) > 0 then
        RunShell(Func.DataFiles[ii].FileName, '/regserver', False);
    end;
  end;
end;

procedure TfrmMain.trvAppChange(Sender: TObject; Node: TTreeNode);
var
  ii: Integer;
  App: TApp;
  ListItem: TListItem;
begin
  lsvFunc.Items.BeginUpdate;
  lsvFunc.Items.Clear;
  App := TApp(Node.Data);
  for ii:=0 to App.FuncCount-1 do
  begin
    if not UCMgr.HasUserFunc(FCurUser, App.Funcs[ii]) then Continue;
    ListItem := lsvFunc.Items.Add;
    ListItem.Data := App.Funcs[ii];
    UCMgr.DisplayFunc(ListItem);
  end;
  lsvFunc.Items.EndUpdate;
end;

procedure TfrmMain.lsvFuncDblClick(Sender: TObject);
var
  Func: TFunc;
begin
  if not Assigned(lsvFunc.Selected) then Exit;
  Func := TFunc(lsvFunc.Selected.Data);
  DownloadFuncPackage(Func);
  RegisterFuncDataFiles(Func);
  if not RunShell(Func.ExeName, Func.Param, True) then
    MessageDlg('°õ¦æ¥¢±Ñ', mtError, [mbOk], 0);
end;

procedure TfrmMain.lsvFuncKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then lsvFuncDblClick(Sender);
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  UCMgr.Free;
  FUserInfoShm.Free;
end;

procedure TfrmMain.pnlCaptionDblClick(Sender: TObject);
var
  UserCode, UserPass: string;
begin
  UserCode := FCurUser.Code;
  UserPass := FCurUser.Password;
  Clear;
  UCMgr.Free;
  UCMgr := TUCMgr.Create(trvApp.Items, trvUserGrp.Items);
  FCurUser := UCMgr.HasUser(UserCode, UserPass);
  trvApp.Selected := trvApp.Items[0];
end;

end.
