unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,Project1_TLB;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    Button2: TButton;
    Label2: TLabel;
    Edit2: TEdit;
    Memo1: TMemo;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    G_Intf : Iabc;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation



{$R *.dfm}

procedure TForm1.Button2Click(Sender: TObject);
begin
    try
      G_Intf := Coabc.CreateRemote(Edit2.text);
    except
      MessageDlg('無法建立遠端物件!', mtError, [mbOK],0);
    end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
    sL_Result: OleVariant;
    sL_ErrorCode: OleVariant;
    sL_ErrorMsg: OleVariant;
begin
    if Assigned(G_Intf) then
    begin
      try
        G_Intf.Method1(Edit1.Text, sL_Result, sL_ErrorCode, sL_ErrorMsg);
        Memo1.Lines.Add(sL_Result + '==' + sL_ErrorCode + '==' +  sL_ErrorMsg);
      except
        Memo1.Lines.Add('呼叫遠端程式失敗!');
      end;
    end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin

    if Assigned(G_Intf) then
    begin
      G_Intf := nil;
    end;  
end;

end.
