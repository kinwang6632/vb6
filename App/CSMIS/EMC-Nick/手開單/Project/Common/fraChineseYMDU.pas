unit fraChineseYMDU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Mask;
const
    SEP = '/';
type
  TfraChineseYMD = class(TFrame)
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

uses UdateTimeu, Ustru;

{$R *.dfm}

{ TfraYMD }

function TfraChineseYMD.getYMD: STring;
var
    sL_Year, sL_Month, sL_Day : String;
begin
//091/08/07
    sL_Year := Trim(Copy(mseYMD.Text,1,3));
    sL_Month := Trim(Copy(mseYMD.Text,5,2));
    sL_Day := Trim(Copy(mseYMD.Text,8,2));

    if sL_Year <>'' then
      sL_Year := Format('%.3d',[StrToInt(sL_Year)])
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

procedure TfraChineseYMD.setYMD(sI_YMD: String);
begin
    mseYMD.Text := sI_YMD;
end;

procedure TfraChineseYMD.mseYMDExit(Sender: TObject);
var
    sL_DateValue : String;
    L_StrList : TStringList;
    sL_Year, sL_Month, sL_Day : String;
begin
    sL_DateValue := mseYMD.Text;
    if length(Trim(TUstr.replaceStr(sL_DateValue,SEP,''))) =0 then Exit;
    L_StrList := TUstr.ParseStrings(sL_DateValue,SEP);
    if L_StrList.Count =3 then
    begin
      sL_Year := IntToStr(StrToIntDef(L_StrList.Strings[0],0) + 1911);
      sL_Month := L_StrList.Strings[1];
      sL_Day := L_StrList.Strings[2];
      sL_DateValue := sL_Year + SEP + sL_Month + SEP +sL_Day;
    end;

    L_StrList := TStringList.Create;

    if (Trim(TUstr.replaceStr(sL_DateValue,SEP,''))<>'') and (not TUdateTime.IsDateStr(sL_DateValue,SEP)) then
    begin
      MessageDlg('請輸入正確的日期格式!',mtError,[mbOK],0);
      mseYMD.SetFocus;
      abort;
    end;

end;

end.
