{ ---------------------------------------------------------------------------- }
{                                                                              }
{ CableSoft Copyright (c) 2002-2003                                            }
{ Unit: Popup Date Form                                                        }
{ Author: Nick                                                                 }
{                                                                              }
{ ---------------------------------------------------------------------------- }

unit DateWindow;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls,
  { Raize CodeSite, http://www.raize.com }
  {$IFDEF DEBUG} CsIntf, {$ENDIF}
  { Developer Express http://www.devexpress.com }
  { Common Unit}
  cxEdit, cxTextEdit, cxLookAndFeelPainters, cxButtons;

type
  TfrmDateWindow = class(TForm)
    Calendar: TMonthCalendar;
    Panel1: TPanel;
    Bevel1: TBevel;
    btnClear: TcxButton;
    btnSelect: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
  private
    { Private declarations }
    FSelectedDate: String;
    FReturnControl: TcxCustomTextEdit;
    procedure SettingPopupPos;
    procedure UserSelectDate;
    procedure UserClearDate;
    procedure SetReturnControl(const Value: TcxCustomTextEdit);
  protected
    { protected declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
    property SelectedDate: String read FSelectedDate;
    property ReturnControl: TcxCustomTextEdit read FReturnControl write
      SetReturnControl;
  end;

var
  frmDateWindow: TfrmDateWindow;

implementation

{$R *.dfm}

uses cbUtilis, Math;

{ TfrmDateWindow }

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams( Params );
  Params.Style := ( WS_POPUP or WS_BORDER );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.FormCreate(Sender: TObject);
begin
  FSelectedDate := DateToStr( Date );
  Calendar.Date := Date;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.FormDeactivate(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.FormShow(Sender: TObject);
begin
  SettingPopupPos;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.SettingPopupPos;
var
  APopUpperPoint, APopLowerPoint: TPoint;
begin
  if Assigned( FReturnControl ) then
  begin
    APopUpperPoint.X := FReturnControl.ClientRect.Left;
    APopUpperPoint.Y := FReturnControl.ClientRect.Top;
    APopLowerPoint.X := FReturnControl.ClientRect.Right;
    APopLowerPoint.Y := FReturnControl.ClientRect.Bottom;
    APopUpperPoint := FReturnControl.ClientToScreen( APopUpperPoint );
    APopLowerPoint := FReturnControl.ClientToScreen( APopLowerPoint );
    Self.Left := APopUpperPoint.X;
    if ( APopUpperPoint.X + Self.Width ) > Screen.Width then
      Self.Left := ( Screen.Width - Self.Width );
    if Self.Left < 0 then Self.Left := 0;
    Self.Top := APopLowerPoint.Y;
    if ( APopLowerPoint.Y + Self.Height ) > Screen.Height then
      Self.Top := ( APopUpperPoint.Y - Self.Height );
    if Self.Top < 0 then Self.Top := 0;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.UserSelectDate;
begin
  FSelectedDate := PadDateText( DateToStr( Calendar.Date ) );
  if Assigned( FReturnControl ) then
  begin
    FReturnControl.Text := FSelectedDate;
    FReturnControl.SelectAll;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.UserClearDate;
begin
  if Assigned( FReturnControl ) then FReturnControl.Text := '';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.SetReturnControl(const Value: TcxCustomTextEdit);
begin
  FReturnControl := Value;
  FSelectedDate := DateToStr( Date );
  if Assigned( FReturnControl ) then
  begin
    if DateTextIsValid( FReturnControl.Text ) then
      FSelectedDate := Nvl( FReturnControl.Text, FSelectedDate );
  end;
  Calendar.Date := StrToDate( FSelectedDate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.btnSelectClick(Sender: TObject);
begin
  UserSelectDate;
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmDateWindow.btnClearClick(Sender: TObject);
begin
  UserClearDate;
  Close;
end;

{ ---------------------------------------------------------------------------- }

end.
