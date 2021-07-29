unit frmCheckUserU;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls;

type
  TfrmCheckUser = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    edtID: TEdit;
    edtPassword: TEdit;
  private
    { Private declarations }
    function GetUserId: string;
    function GetPassword: string;
  public
    { Public declarations }
    property UserId: string read GetUserId;
    property Password: string read GetPassword;
  end;

var
  frmCheckUser: TfrmCheckUser;

implementation

{$R *.DFM}

function TfrmCheckUser.GetUserId: string;
begin
  Result := edtID.Text;
end;

function TfrmCheckUser.GetPassword: string;
begin
  Result := edtPassword.Text;
end;

end.
