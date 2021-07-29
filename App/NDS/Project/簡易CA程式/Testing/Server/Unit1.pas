unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    nG_Result : Integer;    
    procedure addData(sI_Data:String);
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ TForm1 }

procedure TForm1.addData(sI_Data: String);
begin
    Memo1.Lines.Add(sI_Data);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    nG_Result := 0;
end;

end.
 