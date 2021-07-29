unit fraChineseYMU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Mask;
const
    SEP = '/';
type
  TfraChineseYM = class(TFrame)
    mseYM: TMaskEdit;
    procedure mseYMExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function getYM:String;
    procedure setYM(sI_YM:String);
  end;

implementation

uses UdateTimeu, Ustru;

{$R *.dfm}

{ TfraYM }

procedure TfraChineseYM.setYM(sI_YM: String);
begin
    mseYM.Text := sI_YM;
end;

procedure TfraChineseYM.mseYMExit(Sender: TObject);
var
    sL_DateValue : String;
    L_StrList : TStringList;
    sL_Year, sL_Month : String;
    aDate: TDateTime;
begin
  sL_DateValue := mseYM.Text;
  if length(Trim(TUstr.replaceStr(sL_DateValue,SEP,''))) =0 then Exit;
  L_StrList := TUstr.ParseStrings(sL_DateValue,SEP);
  if L_StrList.Count =2 then
  begin
    sL_Year := IntToStr(StrToIntDef(L_StrList.Strings[0],0) + 1911);
    sL_Month := L_StrList.Strings[1];
    sL_DateValue := sL_Year + SEP + sL_Month + SEP + '01';
  end;
  L_StrList.Free;
  if not TryStrToDate( sL_DateValue, aDate ) then
  begin
    MessageDlg('請輸入正確的日期格式!', mtError, [mbOK], 0 );
    mseYM.SetFocus;
    Abort;
  end;
//    if (Trim(TUstr.replaceStr(mseYM.Text,SEP,''))<>'') and (not TUdateTime.IsDateStr(mseYM.Text,SEP)) then => 西元年月日之呼叫方式
//    if (Trim(TUstr.replaceStr(sL_DateValue,SEP,''))<>'') and (not TUdateTime.IsYMStr(sL_DateValue,SEP)) then
//    begin
//      MessageDlg('請輸入正確的日期格式!',mtError,[mbOK],0);
//      mseYM.SetFocus;
//      abort;
//    end;

end;

function TfraChineseYM.getYM: String;
var
    sL_Year, sL_Month : String;
begin
    sL_Year := Trim(Copy(mseYM.Text,1,4));
    sL_Month := Trim(Copy(mseYM.Text,6,2));

    if sL_Year <>'' then
      sL_Year := Format('%.4d',[StrToInt(sL_Year)])
    else
      sL_Year := '    ';

    if sL_Month <>'' then
      sL_Month := Format('%.2d',[StrToInt(sL_Month)])
    else
      sL_Month := '  ';
    result := Format('%s/%s', [sL_Year, sL_Month]);
end;

end.
