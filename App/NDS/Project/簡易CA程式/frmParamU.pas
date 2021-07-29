unit frmParamU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons;

const
    PARAM_FILE_NAME = 'param.txt';  
type
  TfrmParam = class(TForm)
    BitBtn1: TBitBtn;
    StaticText2: TStaticText;
    ComboBox1: TComboBox;
    StaticText3: TStaticText;
    ComboBox2: TComboBox;
    StaticText4: TStaticText;
    edtPopulationID: TEdit;
    GroupBox1: TGroupBox;
    StaticText1: TStaticText;
    StaticText5: TStaticText;
    edtIP: TEdit;
    edtPort: TEdit;
    StaticText6: TStaticText;
    edtTimeout: TEdit;
    StaticText7: TStaticText;
    procedure BitBtn1Click(Sender: TObject);
    procedure edtPopulationIDExit(Sender: TObject);
    procedure edtPortExit(Sender: TObject);
    procedure edtTimeoutExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }

    function getCaIP:String;
    function getCaPort:String;
    function getCaTimeout:String;    
    function getReportbackAvailability:String;
    function getReportbackDate : String;
    function getPopulationID:String;
  end;

var
  frmParam: TfrmParam;

implementation

{$R *.dfm}

procedure TfrmParam.BitBtn1Click(Sender: TObject);
var
    L_StrList : TSTringList;
begin
    L_StrList := TSTringList.Create;
    L_StrList.Add(getCaIP);
    L_StrList.Add(getCaPort);
    L_StrList.Add(getPopulationID);
    L_StrList.Add(getCaTimeout);
    L_StrList.SaveToFile(PARAM_FILE_NAME);
    L_StrList.Free;
    Close;
end;

procedure TfrmParam.edtPopulationIDExit(Sender: TObject);
var
    sL_Text : String;
begin
    sL_Text := (Sender as TEdit).Text;
    if StrToIntDef(sL_Text,-1)=-1 then
    begin
      (Sender as TEdit).SetFocus;
      MessageDlg('請輸入數字!', mtError, [mbOK],0);
    end;
end;

function TfrmParam.getCaIP: String;
begin
    result := Trim(edtIP.Text);
end;

function TfrmParam.getCaPort: String;
begin
    result := Trim(edtPort.Text);
end;

function TfrmParam.getPopulationID: String;
begin
    result := Trim(edtPopulationID.Text);
end;


function TfrmParam.getReportbackAvailability: String;
begin

    result := Copy(ComboBox1.Text,1,1);
end;

function TfrmParam.getReportbackDate: String;
begin
    result := ComboBox2.Text;

end;

procedure TfrmParam.edtPortExit(Sender: TObject);
var
    sL_Text : String;
begin
    sL_Text := (Sender as TEdit).Text;
    if StrToIntDef(sL_Text,-1)=-1 then
    begin
      (Sender as TEdit).SetFocus;
      MessageDlg('請輸入數字!', mtError, [mbOK],0);
    end;
end;

function TfrmParam.getCaTimeout: String;
begin
    result := Trim(edtTimeout.Text);
end;

procedure TfrmParam.edtTimeoutExit(Sender: TObject);
var
    sL_Text : String;
begin
    sL_Text := (Sender as TEdit).Text;
    if StrToIntDef(sL_Text,-1)=-1 then
    begin
      (Sender as TEdit).SetFocus;
      MessageDlg('請輸入數字!', mtError, [mbOK],0);
    end;

end;

procedure TfrmParam.FormCreate(Sender: TObject);
var
    L_StrList : TSTringList;
begin
    if FileExists(PARAM_FILE_NAME) then
    begin
      L_StrList := TSTringList.Create;
      L_StrList.LoadFromFile(PARAM_FILE_NAME);
      if L_StrList.Count = 4 then
      begin
        edtIP.Text := L_StrList.Strings[0];
        edtPort.Text := L_StrList.Strings[1];
        edtPopulationID.Text := L_StrList.Strings[2];
        edtTimeout.Text := L_StrList.Strings[3];        
      end;
      L_StrList.Free;
    end;
end;

end.
