unit fraYMDU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Mask;
const
    SEP = '/';
type
  TfraYMD = class(TFrame)
    mseYMD: TMaskEdit;
    procedure mseYMDExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function getYMD:String;
    procedure setYMD(sI_YMD:String);
  end;

implementation

uses UdateTimeu, Ustru, dtmMainU;

{$R *.dfm}

{ TfraYMD }

function TfraYMD.getYMD: STring;
var
    sL_Year, sL_Month, sL_Day : String;
begin
    sL_Year := Trim(Copy(mseYMD.Text,1,4));
    sL_Month := Trim(Copy(mseYMD.Text,6,2));
    sL_Day := Trim(Copy(mseYMD.Text,9,2));
    
    if sL_Year <>'' then
      sL_Year := Format('%.4d',[StrToInt(sL_Year)])
    else
      sL_Year := '    ';

    if sL_Month <>'' then
      sL_Month := Format('%.2d',[StrToInt(sL_Month)])
    else
      sL_Month := '  ';

    if sL_Day <>'' then
      sL_Day := Format('%.2d',[StrToInt(sL_Day)])
    else
      sL_Day := '  ';
      
    result := Format('%s/%s/%s', [sL_Year, sL_Month, sL_Day]);
end;

procedure TfraYMD.setYMD(sI_YMD: String);
begin
    mseYMD.Text := sI_YMD;
end;

procedure TfraYMD.mseYMDExit(Sender: TObject);
begin

    if (Trim(TUstr.replaceStr(mseYMD.Text,SEP,''))<>'') and (not TUdateTime.IsDateStr(mseYMD.Text,SEP)) then
    begin
      MessageDlg(dtmMain.getLanguageSettingInfo('common_Msg_1_Content'),mtError,[mbOK],0);//請輸入正確的日期格式!
      mseYMD.SetFocus;
      abort;
    end;

end;

end.
