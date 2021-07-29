unit fraYMU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask;
const
    SEP = '/';

type
  TfraYM = class(TFrame)
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

function TfraYM.getYM: String;
var
    sL_Year, sL_Month : String;
begin
    sL_Year := Trim(Copy(mseYM.Text,1,4));
    sL_Month := Trim(Copy(mseYM.Text,6,2));

    if sL_Year <>'' then
      sL_Year := Format('%.3d',[StrToInt(sL_Year)])
    else
      sL_Year := '    ';

    if sL_Month <>'' then
      sL_Month := Format('%.2d',[StrToInt(sL_Month)])
    else
      sL_Month := '  ';
      
    result := Format('%s/%s', [sL_Year, sL_Month]);
end;

procedure TfraYM.setYM(sI_YM: String);
begin
    mseYM.Text := sI_YM;
end;

procedure TfraYM.mseYMExit(Sender: TObject);
begin
    if (Trim(TUstr.replaceStr(mseYM.Text,SEP,''))<>'') and (not TUdateTime.IsDateStr(mseYM.Text+'/01',SEP)) then
    begin
      MessageDlg('請輸入正確的年月格式!',mtError,[mbOK],0);
      mseYM.SetFocus;
    end;

end;

end.
