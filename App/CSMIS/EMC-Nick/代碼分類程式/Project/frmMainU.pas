unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

const
    PARAM_COUNT = 3;
type
  TfrmMain = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private


    { Private declarations }
  public
    { Public declarations }
    sG_LoginID, sG_LoginName : String;
    //sG_DbUserID, sG_DbPasswd, sG_DbAlias : String;

  end;

var
  frmMain: TfrmMain;

implementation

uses frmLoginU, dtmMainU, frmCateCode1U, frmReportMainU;

{$R *.dfm}

procedure TfrmMain.BitBtn3Click(Sender: TObject);
begin

    Close;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
    bL_LoginOk : boolean;
begin
    if dtmMain.sG_NeedLogin='N' then   //各系統台使用
    begin

      if not dtmMain.connToDb(dtmMain.sG_DbAlias,dtmMain.sG_DbUserName,dtmMain.sG_DbPassword) then
      begin
        MessageDlg('資料庫連接失敗!程式即將結束.', mtError, [mbOK],0);
        Application.Terminate;
      end;

    end
    else    //營管使用
    begin
      frmLogin := TfrmLogin.Create(Application);
      with frmLogin do
      begin

        if ShowModal=mrCancel then
        begin
          bL_LoginOk := false;
        end
        else
        begin
          bL_LoginOk := true;
          sG_LoginID := sG_USerID;
          sG_LoginName := sG_UserName;
          Close;
          free;
        end;
      end;

      if not bL_LoginOk then
        Application.Terminate
      else
      begin
{
        if ParamCount <> PARAM_COUNT then
        begin
          MessageDlg('傳入之參數個數錯誤!', mtError, [mbOK],0);
          Application.Terminate;
        end
        else
        begin
}
          if not dtmMain.connToDb(dtmMain.sG_DbAlias,dtmMain.sG_DbUserName,dtmMain.sG_DbPassword) then
          begin
            MessageDlg('資料庫連接失敗!程式即將結束.', mtError, [mbOK],0);
            Application.Terminate;
          end;
        end;

//      end;
    end;

end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    frmCateCode1 := TfrmCateCode1.Create(Application);
    frmCateCode1.ShowModal;
    frmCateCode1.Free;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
    frmReportMain := TfrmReportMain.Create(Application);
    frmReportMain.ShowModal;
    frmReportMain.Free;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    bL_IsEnable : Boolean;
begin
    self.Caption := dtmMain.setCaption('SO8F00','[二階代碼分類程式]');

    //down, 設定權限...
    if (dtmMain.sG_IsSupervisor='Y') or (dtmMain.sG_NeedLogin='Y') then
    begin
      BitBtn2.Enabled := true;
    end
    else
    begin
      dtmMain.activePrivDataset(dtmMain.sG_CompCode,dtmMain.sG_UserID);
      bL_IsEnable := dtmMain.checkPriv('SO8F20');
      BitBtn2.Enabled := bL_IsEnable;
    end;
    //up, 設定權限...

end;

end.
