unit frmCD078B3_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,
  cbStyleController,
  dxSkinsCore, dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit, cxTextEdit,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmCD078B3_2 = class(TForm)
    edtComment: TcxTextEdit;
    btnSave: TButton;
    Button1: TButton;
    Bevel1: TBevel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    function GetCommentText: String;
    procedure SetCommentText(const AValue: String);
  public
    { Public declarations }
    property CommentText: String read GetCommentText write SetCommentText;
  end;

var
  frmCD078B3_2: TfrmCD078B3_2;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_2.FormCreate(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_2.FormShow(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_2.FormDestroy(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3_2.GetCommentText: String;
begin
  Result := edtComment.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_2.SetCommentText(const AValue: String);
begin
  edtComment.Text := AValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_2.btnSaveClick(Sender: TObject);
begin
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

end.
