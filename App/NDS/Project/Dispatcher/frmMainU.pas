unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, ExtCtrls;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    AboutSMSGateway1: TMenuItem;
    N5: TMenuItem;
    DB1: TMenuItem;
    procedure N3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure AboutSMSGateway1Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure DB1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure terminateApp;
  public
    { Public declarations }
    sG_OperatorName : STring;
  end;

var
  frmMain: TfrmMain;

implementation

uses frmSysParamU, frmLoginU, dtmMainU, frmSendCommandU, frmAboutU,
  frmDbInfoU;

{$R *.dfm}

procedure TfrmMain.N3Click(Sender: TObject);
begin
    terminateApp;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
    frmMain.Caption := 'NDS Dispatcher 主功能畫面';
end;

procedure TfrmMain.N2Click(Sender: TObject);
begin
    frmSendCommand := TfrmSendCommand.Create(Application);
    frmSendCommand.ShowModal;
    frmSendCommand.Free;
end;

procedure TfrmMain.AboutSMSGateway1Click(Sender: TObject);
begin
    frmAbout.ShowModal;
end;

procedure TfrmMain.terminateApp;
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      Application.Terminate;
end;

procedure TfrmMain.N5Click(Sender: TObject);
begin
    frmSysParam := TfrmSysParam.Create(Application);
    frmSysParam.ShowModal;
    frmSysParam.Free;
end;

procedure TfrmMain.DB1Click(Sender: TObject);
begin
    frmDbInfo := TfrmDbInfo.Create(Application);
    frmDbInfo.ShowModal;
    frmDbInfo.Free;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    ;
end;

end.
