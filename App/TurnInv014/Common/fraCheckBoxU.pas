unit fraCheckBoxU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TfraCheckBox = class(TFrame)
    chk: TCheckBox;
  private
    FCheckValue: String;
    FNoneCheckValue: String;
    procedure SetCheckValue(const Value: String);
    procedure SetNoneCheckValue(const Value: String);
  private
    { Private declarations }
    property CheckValue:String read FCheckValue write SetCheckValue;
    property NoneCheckValue:String read FNoneCheckValue write SetNoneCheckValue;
  public
    { Public declarations }
    function getStatusValue:String;
    procedure setStatus(sI_Value:String);
    procedure init(sI_Caption, sI_CheckValue, sI_NoneCheckValue:String);
  end;

implementation

{$R *.dfm}

{ TfraCheckBox }

function TfraCheckBox.getStatusValue: String;
var
    sL_Result : String;
begin
    if chk.Checked then
      sL_Result := FCheckValue
    else
      sL_Result := FNoneCheckValue;

    result := sL_Result;  
end;

procedure TfraCheckBox.init(sI_Caption, sI_CheckValue,
  sI_NoneCheckValue: String);
begin
    chk.Caption := sI_Caption;
    SetCheckValue(sI_CheckValue);
    SetNoneCheckValue(sI_NoneCheckValue);
end;

procedure TfraCheckBox.SetCheckValue(const Value: String);
begin
  FCheckValue := Value;
end;

procedure TfraCheckBox.SetNoneCheckValue(const Value: String);
begin
  FNoneCheckValue := Value;
end;

procedure TfraCheckBox.setStatus(sI_Value: String);
begin
    if sI_Value=FNoneCheckValue then
      chk.Checked := false
    else
      chk.Checked := true;
end;

end.
