unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, menus;

type
  TfrmMain = class(TForm)
    btnTL0020: TBitBtn;
    btnExit: TBitBtn;
    btnTL0030: TBitBtn;
    btnTL0040: TBitBtn;
    procedure btnExitClick(Sender: TObject);
    procedure btnTL0020Click(Sender: TObject);
    procedure btnTL0030Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnTL0040Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses dtmTL0020U, frmTL0020U, dtmTL0030U, frmTL0030U, dllUCU, UCommonU,
  dtmTL0040U, frmTL0040U;

{$R *.DFM}

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
        Application.Terminate;
end;

procedure TfrmMain.btnTL0020Click(Sender: TObject);
begin
    dtmTL0020 := TdtmTL0020.Create(Application);
    frmTL0020 := TfrmTL0020.Create(Application);
    frmTL0020.ShowModal;
    frmTL0020.Free;
    dtmTL0020.Free;
end;

procedure TfrmMain.btnTL0030Click(Sender: TObject);
begin
    dtmTL0030 := TdtmTL0030.Create(Application);
    frmTL0030 := TfrmTL0030.Create(Application);
    frmTL0030.ShowModal;
    frmTL0030.Free;
    dtmTL0030.Free;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
   ExeFileName, ExePath : String;

   aComp: TComponent;
begin
  //down, get user data...
  GetCurUser(sC_UserID, sC_UserName, sC_Password, sC_UserGrp);
  //up, get user data

  if ParamStr(1) <> '' then
  begin
    Application.ShowMainForm := False;
    aComp := FindComponent(ParamStr(1));
    if Assigned(aComp) and (aComp IS TBitBtn) then
    begin
      TBitBtn(aComp).Click;
      Application.Terminate;
      Exit;
    end
    else if Assigned(aComp) and (aComp IS TMenuItem)then
    begin
      TMenuItem(aComp).Click;
      Application.Terminate;
      Exit;
    end;
  end;

end;

procedure TfrmMain.btnTL0040Click(Sender: TObject);
begin
    dtmTL0040 := TdtmTL0040.Create(Application);
    frmTL0040 := TfrmTL0040.Create(Application);
    frmTL0040.ShowModal;
    frmTL0040.Free;
    dtmTL0040.Free;
end;

end.
