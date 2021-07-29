unit cbEmail;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, cbStyleModule, cbConfigModule, RzPanel,
  cxControls, cxContainer, cxListBox, Menus, cxLookAndFeelPainters, cxEdit,
  cxTextEdit, cxHyperLinkEdit, cxButtons, cxGraphics;

type
  TfmEmail = class(TForm)
    RzPanel1: TRzPanel;
    RzPanel2: TRzPanel;
    lstEmail: TcxListBox;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    RzPanel3: TRzPanel;
    edtEmail: TcxHyperLinkEdit;
    btnAdd: TcxButton;
    btnDelete: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure lstEmailDrawItem(AControl: TcxListBox; ACanvas: TcxCanvas;
      AIndex: Integer; ARect: TRect; AState: TOwnerDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmEmail: TfmEmail;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmEmail.FormCreate(Sender: TObject);
begin
  lstEmail.Clear;
  edtEmail.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmEmail.btnAddClick(Sender: TObject);
begin
  if ( Trim( edtEmail.Text ) <> EmptyStr ) then
    lstEmail.Items.Add( edtEmail.Text );
  edtEmail.Clear;
  if ( edtEmail.CanFocusEx ) then edtEmail.SetFocus;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmEmail.btnDeleteClick(Sender: TObject);
var
  aPos: Integer;
begin
  aPos := lstEmail.ItemIndex;
  if ( lstEmail.ItemIndex >= 0 ) then
    lstEmail.Items.Delete( lstEmail.ItemIndex );
  if ( aPos < lstEmail.Items.Count  ) then
    lstEmail.ItemIndex := aPos
  else if ( lstEmail.Items.Count > 0 ) then
    lstEmail.ItemIndex := ( lstEmail.Items.Count - 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmEmail.lstEmailDrawItem(AControl: TcxListBox;
  ACanvas: TcxCanvas; AIndex: Integer; ARect: TRect;
  AState: TOwnerDrawState);
begin
  ACanvas.FillRect( ARect );
  ARect.Left := ( ARect.Left + 3 );
  ACanvas.DrawImage( StyleModule.OptionImageList, ARect.Left, ARect.Top, 4 );
  ARect.Left := ( ARect.Left + StyleModule.OptionImageList.Width + 10 );
  ACanvas.Font.Style := ( ACanvas.Font.Style + [fsUnderline] );
  if ( odSelected in AState ) then
    ACanvas.Font.Color := clHighlightText
  else
    ACanvas.Font.Color := clBlue;
  ACanvas.DrawTexT( AControl.Items[AIndex], ARect, 0 );
end;

{ ---------------------------------------------------------------------------- }

end.
