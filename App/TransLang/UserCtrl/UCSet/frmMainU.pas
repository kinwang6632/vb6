unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ToolWin, IniFiles, utlUCU, ExtCtrls, StdCtrls, Buttons, Menus;

type
  TfrmMain = class(TForm)
    stbInfo: TStatusBar;
    mnuMain: TMainMenu;
    mnuFile: TMenuItem;
    E1: TMenuItem;
    V1: TMenuItem;
    mnuAppFunc: TMenuItem;
    mnuUserGrpUser: TMenuItem;
    mnuSep1: TMenuItem;
    mnuExit: TMenuItem;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure mnuExitClick(Sender: TObject);
    procedure SelectObject(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses frmAppU, frmUserGrpU;

{$R *.DFM}

procedure TfrmMain.FormShow(Sender: TObject);
begin
  UCMgr := TUCMgr.Create(frmApp.trvApp.Items, frmUserGrp.trvUserGrp.Items);
  frmApp.BringToFront;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if MessageDlg('要儲存先前的異動嗎??', mtConfirmation,
    [mbYes, mbNo], 0) = mrYes then
    UCMgr.SaveToFile;
  UCMgr.Free;
end;

procedure TfrmMain.mnuExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.SelectObject(Sender: TObject);
begin
  if Sender = mnuAppFunc then frmApp.BringToFront
  else if Sender = mnuUserGrpUser then frmUserGrp.BringToFront;
end;

end.
