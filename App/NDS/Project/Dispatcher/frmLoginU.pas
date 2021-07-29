unit frmLoginU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type
  TfrmLogin = class(TForm)
    Panel1: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    edtUserID: TEdit;
    edtPasswd: TEdit;
    bbnConfirm: TBitBtn;
    bbnCancel: TBitBtn;
    Panel2: TPanel;
    Panel3: TPanel;
    pnlPrgName: TPanel;
    btnReSet: TBitBtn;
    procedure bbnCancelClick(Sender: TObject);
    procedure btnReSetClick(Sender: TObject);
    procedure bbnConfirmClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    function IsDataOK : boolean;
  public
    { Public declarations }
    sG_UserID : String;
    function checkPasswd(sI_UserID,sI_Passwd : String; var sI_UserName:STring) : Boolean;
  end;

var
  frmLogin: TfrmLogin;

implementation

uses dtmMainU, frmMainU, frmSysParamU;

{$R *.dfm}

procedure TfrmLogin.bbnCancelClick(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmLogin.btnReSetClick(Sender: TObject);
begin
    edtUserID.Text := '';
    edtPasswd.Text := '';
    edtUserID.SetFocus ;
end;

procedure TfrmLogin.bbnConfirmClick(Sender: TObject);
var
  sL_UserID,sL_Passwd, sL_UserName : String;
  bL_IsPasswordOK : Boolean;
begin
    if IsDataOK then
    begin
      sL_UserID := edtUserID.Text;
      sL_Passwd := edtPasswd.Text;

      bL_IsPasswordOK := checkPasswd(sL_UserID,sL_Passwd, sL_UserName);
      if bL_IsPasswordOK then
      begin
        frmLogin.Visible := false;
        frmMain := TfrmMain.Create(Application);
        frmMain.sG_OperatorName :=  sL_UserName;
        frmMain.ShowModal;
        frmMain.Free;
      end
      else
      begin
        edtUserID.SetFocus;
        edtPasswd.Text := '';
        MessageDlg('帳號及密碼錯誤',mtError, [mbOK],0);
        exit;
      end;
    end
    else
    begin
      edtUserID.SetFocus;
      MessageDlg('請輸入帳號及密碼',mtError, [mbOK],0);
      exit;
    end;

end;

function TfrmLogin.IsDataOK: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if edtUserID.Text = '' then
    begin
      bL_Result := false;
    end;
    if edtPasswd.Text = '' then
    begin
      bL_Result := false;
    end;
    result := bL_Result;
end;



procedure TfrmLogin.FormCreate(Sender: TObject);
begin
    dtmMain.activeCDS;
    dtmMain.getAllUserID;


//暫時先加
{
dtmMain.cdsParam.Append;

dtmMain.cdsParam.FieldByName('sServerIP').AsString := '5';
dtmMain.cdsParam.FieldByName('nSPortNo').AsInteger := 5;
dtmMain.cdsParam.FieldByName('nRPortNo').AsInteger := 5;
dtmMain.cdsParam.FieldByName('sSysName').AsString := '5';
dtmMain.cdsParam.FieldByName('sSysVersion').AsString := '5';
dtmMain.cdsParam.FieldByName('sLogPath').AsString := 'c:\temp';
dtmMain.cdsParam.FieldByName('nTimeOut').AsInteger := 5;
dtmMain.cdsParam.FieldByName('nMaxCommandCount').AsInteger := 5;
dtmMain.cdsParam.FieldByName('bCommandLog').AsBoolean := true;
dtmMain.cdsParam.FieldByName('bResponseLog').AsBoolean := true;
dtmMain.cdsParam.FieldByName('dUptTime').AsString := '2003/04/26';
dtmMain.cdsParam.FieldByName('sUptName').AsString := '5';
dtmMain.cdsParam.FieldByName('nCmdRefreshRate1').AsInteger := 5;
dtmMain.cdsParam.FieldByName('nCmdRefreshRate2').AsInteger := 5;
dtmMain.cdsParam.FieldByName('nCmdResentTimes').AsInteger := 5;
dtmMain.cdsParam.FieldByName('bShowUI').AsBoolean := true;
dtmMain.cdsParam.FieldByName('sSHotTime').AsString := '5';
dtmMain.cdsParam.FieldByName('sEHotTime').AsString := '5';

dtmMain.cdsParam.FieldByName('nVersion').AsInteger := 5;
dtmMain.cdsParam.FieldByName('nFormID').AsInteger := 6;
dtmMain.cdsParam.FieldByName('nToID').AsInteger := 7;
dtmMain.cdsParam.FieldByName('sSecuritytype').AsString := 'M';


dtmMain.cdsParam.FieldByName('nConnectionID').AsInteger := 5;
dtmMain.cdsParam.FieldByName('sKey').AsString := '6';
dtmMain.cdsParam.FieldByName('nForwardTolerance').AsInteger := 7;
dtmMain.cdsParam.FieldByName('nBackwardTolerance').AsInteger := 8;
dtmMain.cdsParam.FieldByName('nCurrency').AsInteger := 8;


dtmMain.cdsParam.Post;

SetCurrentDir(ExtractFilePath(Application.ExeName));

if (dtmMain.cdsParam.ChangeCount>0) then
  dtmMain.cdsParam.SaveToFile('D:\NDS\Project\Src\param.dat');

frmSysParam.preSaveDB;
}

end;

procedure TfrmLogin.FormShow(Sender: TObject);
begin
    frmLogin.Caption := '系統登入';
end;

function TfrmLogin.checkPasswd(sI_UserID, sI_Passwd: String; var sI_UserName:STring): Boolean;
var
    sL_ExePath : String;
begin
    sG_UserID := sI_UserID;

    if dtmMain.cdsUser.Locate('SID;SPASSWORD',VarArrayOf([sI_UserID, sI_Passwd]) , []) then
    begin
      sI_UserName:= dtmMain.cdsUser.fieldbyname('sName').asstring;    
      Result := true;
    end
    else
      Result := false;
end;

end.

