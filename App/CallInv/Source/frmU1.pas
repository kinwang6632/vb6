unit frmU1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TfrmBrowINV = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmBrowINV: TfrmBrowINV;

implementation
  uses CommU;
{$R *.dfm}

procedure TfrmBrowINV.FormCreate(Sender: TObject);
begin
  AppCreate;
end;

end.
