unit frmDispatcherInfoU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ImgList, ComCtrls, ExtCtrls;

type
  TfrmDispatcherInfo = class(TForm)
    Panel1: TPanel;
    StaticText1: TStaticText;
    livDispatcherInfo: TListView;
    ImageList1: TImageList;
    BitBtn1: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDispatcherInfo: TfrmDispatcherInfo;

implementation

uses frmMainU, ConstU;

{$R *.dfm}

procedure TfrmDispatcherInfo.FormShow(Sender: TObject);
var
    sL_Ip, sL_CompName : String;
    ii : Integer;
    L_ListItem : TListItem;
    L_Obj : TCompAndIpInfoObj;

    L_TmpStrList : TStringList;

begin
{
    Caption := '連線中的 Dispatcher 資訊';

    livDispatcherInfo.Clear;
    L_TmpStrList := TStringList.Create;
    for ii:=0 to frmMain.G_RemoteInfoStrList.Count -1 do
    begin

      L_Obj := (frmMain.G_RemoteInfoStrList.Objects[ii] As TCompAndIpInfoObj);
      sL_CompCode :=  Format('%.2d',[StrToIntDef(frmMain.G_RemoteInfoStrList.Strings[ii],-1)]);
      sL_Ip := L_Obj.IP;
      L_ListItem := livDispatcherInfo.Items.Add;
      L_ListItem.Caption := '系統台代碼:' + sL_CompCode + ' IP:' + sL_Ip;

      if L_TmpStrList.IndexOf(sL_Ip)<0 then
        L_TmpStrList.Add(sL_Ip);
    end;
    StaticText1.Caption := '目前共管理 ' + IntToStr(frmMain.G_RemoteInfoStrList.Count) + ' 個系統台. 分屬於 ' + IntToStr(L_TmpStrList.count) + ' 個 Dispatcher ';

    L_TmpStrList.Free;
}

    Caption := '簡易 CA 連線資訊';


    livDispatcherInfo.Clear;
    L_TmpStrList := TStringList.Create;
    for ii:=0 to frmMain.G_ListenerThread.nG_AryLength -1 do
    begin
      sL_Ip := frmMain.G_ListenerThread.G_ClientArray[ii].IP;
      sL_CompName := frmMain.G_ListenerThread.G_ClientArray[ii].CompName;
      L_ListItem := livDispatcherInfo.Items.Add;
      L_ListItem.Caption := '系統台:' + sL_CompName + '   IP:' + sL_Ip;

      if L_TmpStrList.IndexOf(sL_CompName)<0 then
        L_TmpStrList.Add(sL_CompName);
    end;
    StaticText1.Caption := '目前共管理 ' + IntToStr(frmMain.G_ListenerThread.nG_AryLength) + ' 個IP. 分屬於 ' + IntToStr(L_TmpStrList.count) + ' 個系統台';

    L_TmpStrList.Free;

end;

procedure TfrmDispatcherInfo.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

end.
