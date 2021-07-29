unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin;

type
  TfrmMain = class(TForm)
    CoolBar1: TCoolBar;
    ToolBar5: TToolBar;
    ToolButton36: TToolButton;
    ToolBar6: TToolBar;
    ToolButton38: TToolButton;
    ToolButton39: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    procedure ToolButton36Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ToolButton39Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton38Click(Sender: TObject);
  private
    { Private declarations }
    bG_LoginOk : boolean;
  public
    { Public declarations }
    function getCaption(sI_FunName:String):String;
  end;

var
  frmMain: TfrmMain;

implementation

uses frmLoginU, dtmMainU, frmGroupDataU, frmCompDataU, dtmCompDataU,
  dtmDeptDataU, frmDeptDataU, frmUserDataU;

{$R *.dfm}

procedure TfrmMain.ToolButton36Click(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var

    L_frmLogin: TfrmLogin;
begin

    L_frmLogin := TfrmLogin.Create(Application);
    with L_frmLogin do
    begin

      if ShowModal=mrCancel then
      begin

        bG_LoginOk := false;
      end
      else
      begin

        bG_LoginOk := true;
        Close;
        free;
      end;
    end;



    if bG_LoginOk then
    begin

    end
    else
    begin
      L_frmLogin.Free;
      Application.Terminate;
    end;

end;

function TfrmMain.getCaption(sI_FunName: String): String;
begin
    result := 'CSIS-' + sI_FunName + ' - ' + dtmMain.getUserName;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin

    self.Caption := getCaption('功能畫面');
end;

procedure TfrmMain.ToolButton39Click(Sender: TObject);

begin

    frmGroupData := TfrmGroupData.Create(Application);
    frmGroupData.ShowModal;
    frmGroupData.Free;

end;

procedure TfrmMain.ToolButton1Click(Sender: TObject);
begin
    dtmCompData := TdtmCompData.Create(Application);
    frmCompData := TfrmCompData.Create(Application);
    frmCompData.ShowModal;
    frmCompData.Free;
    dtmCompData.Free;

end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
begin
    dtmDeptData := TdtmDeptData.Create(Application);
    frmDeptData := TfrmDeptData.Create(Application);
    frmDeptData.ShowModal;
    frmDeptData.Free;
    dtmDeptData.Free;

end;

procedure TfrmMain.ToolButton38Click(Sender: TObject);
begin
    frmUserData := TfrmUserData.Create(Application);
    frmUserData.ShowModal;
    frmUserData.Free;

end;

end.
