unit frmSo18C0U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

const
    PARAM_COUNT= 8;
type
  TfrmSo18C0 = class(TForm)
    btnSo8C10: TBitBtn;
    btnExit: TBitBtn;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    btnSo8C20: TBitBtn;
    btnSo8C30: TBitBtn;
    btnSo8C40: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSo8C10Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure btnSo8C20Click(Sender: TObject);
    procedure btnSo8C30Click(Sender: TObject);
    procedure btnSo8C40Click(Sender: TObject);
  private
    { Private declarations }

  public
    { Public declarations }
    sG_IsSupervisor, sG_OperatorID, sG_Operator, sG_CompID, sG_CompName, sG_DbUserID, sG_DbPasswd, sG_DbAlias : String;    
    function getCaption(sI_FormID, sI_FormName : String):String;
  end;

var
  frmSo18C0: TfrmSo18C0;

implementation

uses frmSo18C1U, dtmMainU, UdateTimeu, frmSo18C2U, frmSo18C3U, frmSo18C4U;

{$R *.dfm}

procedure TfrmSo18C0.FormCreate(Sender: TObject);
begin
    if ParamCount <> PARAM_COUNT then
    begin
      MessageDlg('傳入之參數個數錯誤!', mtError, [mbOK],0);
      Application.Terminate;
    end
    else
    begin
      //N 1 Howard 1 觀昇有限 gicmis may mis
      sG_IsSupervisor := Uppercase(ParamStr(1));
      sG_OperatorID := ParamStr(2);
      sG_Operator := ParamStr(3);
      sG_CompID := ParamStr(4);
      sG_CompName := ParamStr(5);
      sG_DbUserID := ParamStr(6);
      sG_DbPasswd := ParamStr(7);
      sG_DbAlias := ParamStr(8);


{
showmessage(sG_Operator);
showmessage(sG_CompID);
showmessage(sG_CompName);
showmessage(sG_DbUserID);
showmessage(sG_DbPasswd);
showmessage(sG_DbAlias);
}
    end;

end;

function TfrmSo18C0.getCaption(sI_FormID, sI_FormName: String): String;
begin
    result := sG_CompName + '-' + sG_Operator + '--' + sI_FormName + '-[' +  sI_FormID + ']';
end;

procedure TfrmSo18C0.FormShow(Sender: TObject);
begin
//    showmessage('form show');
    if not dtmMain.connectDB(sG_DbAlias, sG_DbUserID, sG_DbPasswd) then
    begin
      MessageDlg('資料庫連結失敗!',mtWarning, [mbOK],0 );
      Application.Terminate;
      Exit;
    end;  

    //down, 設定權限...
    if sG_IsSupervisor='Y' then
    begin
      btnSo8C10.Enabled :=  true;
      btnSo8C20.Enabled :=  true;
      btnSo8C30.Enabled :=  true;
      btnSo8C40.Enabled :=  true;      
    end
    else
    begin
      dtmMain.activePrivDataset(sG_CompID,sG_OperatorID);
      btnSo8C10.Enabled :=  dtmMain.checkPriv('SO8C10');
      btnSo8C20.Enabled :=  dtmMain.checkPriv('SO8C20');
      btnSo8C30.Enabled :=  dtmMain.checkPriv('SO8C30');
      btnSo8C40.Enabled :=  dtmMain.checkPriv('SO8C40');
    end;
    //up, 設定權限...

    self.Caption := getCaption('SO8C00','手開單據管理');
end;

procedure TfrmSo18C0.btnExitClick(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmSo18C0.btnSo8C10Click(Sender: TObject);
var
    L_Frm : TfrmSo18C1;
begin
    if dtmMain.connectDB(sG_DbAlias, sG_DbUserID, sG_DbPasswd) then
    begin
      L_Frm := TfrmSo18C1.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end
    else
      MessageDlg('資料庫連結失敗!',mtWarning, [mbOK],0 );
end;

procedure TfrmSo18C0.Button1Click(Sender: TObject);
var
    dL_TmpDate : TDate;
begin
    dL_TmpDate := TUdateTime.CDate2Date(Label1.Caption);
    Label2.Caption := DateToStr(dL_TmpDate);
end;

procedure TfrmSo18C0.Button2Click(Sender: TObject);
begin
    Label1.Caption := TUdateTime.CDateStr(date,9);
end;

procedure TfrmSo18C0.btnSo8C20Click(Sender: TObject);
var
    L_Frm : TfrmSo18C2;
begin
    if dtmMain.connectDB(sG_DbAlias, sG_DbUserID, sG_DbPasswd) then
    begin
      L_Frm := TfrmSo18C2.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end
    else
      MessageDlg('資料庫連結失敗!',mtWarning, [mbOK],0 );
end;

procedure TfrmSo18C0.btnSo8C30Click(Sender: TObject);
var
    L_Frm : TfrmSo18C3;
begin
    if dtmMain.connectDB(sG_DbAlias, sG_DbUserID, sG_DbPasswd) then
    begin
      L_Frm := TfrmSo18C3.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end
    else
      MessageDlg('資料庫連結失敗!',mtWarning, [mbOK],0 );
end;

procedure TfrmSo18C0.btnSo8C40Click(Sender: TObject);
var
    L_Frm : TfrmSo18C4;
begin
    if dtmMain.connectDB(sG_DbAlias, sG_DbUserID, sG_DbPasswd) then
    begin
      L_Frm := TfrmSo18C4.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end
    else
      MessageDlg('資料庫連結失敗!',mtWarning, [mbOK],0 );
end;

end.
