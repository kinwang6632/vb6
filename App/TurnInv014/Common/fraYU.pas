unit fraYU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask;

type
  TfraY = class(TFrame)
    mseYear: TMaskEdit;
  private
    { Private declarations }
  public
    { Public declarations }
    function getY:String;
    procedure setY(sI_Y:String);
  end;

implementation

{$R *.dfm}

{ TfraY }

function TfraY.getY: String;
begin
    result := Trim(mseYear.Text);
end;

procedure TfraY.setY(sI_Y: String);
begin
    mseYear.Text := sI_Y;
end;

end.
