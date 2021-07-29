unit frmParamU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

type
  TfrmParam = class(TForm)
    BitBtn1: TBitBtn;
    StaticText1: TStaticText;
    Edit1: TEdit;
    StaticText2: TStaticText;
    ComboBox1: TComboBox;
    StaticText3: TStaticText;
    ComboBox2: TComboBox;
    StaticText4: TStaticText;
    Edit2: TEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure Edit2Exit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function getRegionKey:String;
    function getReportbackAvailability:String;
    function getReportbackDate : String;
    function getPopulationID:String;
  end;

var
  frmParam: TfrmParam;

implementation

{$R *.dfm}

procedure TfrmParam.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmParam.Edit2Exit(Sender: TObject);
var
    sL_Text : String;
begin
    sL_Text := (Sender as TEdit).Text;
    if StrToIntDef(sL_Text,-1)=-1 then
    begin
      (Sender as TEdit).SetFocus;
      MessageDlg('½Ð¿é¤J¼Æ¦r!', mtError, [mbOK],0);
    end;
end;

function TfrmParam.getPopulationID: String;
begin
    result := Trim(Edit2.Text);
end;

function TfrmParam.getRegionKey: String;
begin
    result := Trim(Edit1.Text);
end;

function TfrmParam.getReportbackAvailability: String;
begin

    result := Copy(ComboBox1.Text,1,1);
end;

function TfrmParam.getReportbackDate: String;
begin
    result := ComboBox2.Text;
    
end;

end.
