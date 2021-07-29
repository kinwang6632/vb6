unit cbClientClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, ExtCtrls, ShellAPI, Forms,
  TrayIcon,
{ Developer Express: }
  dxRibbon, dxRibbonForm;

const
  WM_REALCLOSE = WM_USER + 1;
  WM_MANUALREFRESH = WM_USER + 2;
  WM_READY = WM_USER + 3;
  WM_CSMISCLOSE = WM_USER + 4;

type
  TServerConnectState = (sctOnLine, sctOffLine, sctProcDate);
  TUserListDisplayType = (gkWorkClass, gkStatus);
  TClinetUserStatus = (usOffline, usOnline, usBusy);
  THrClientFormKind = (ftAnn, ftMsg);

const
  //Parameters for ApplyStyleConversion
  TEXT_BOLD = 1;
  TEXT_ITALIC = 2;
  TEXT_UNDERLINE = 3;
  TEXT_APPLYFONTNAME = 4;
  TEXT_APPLYFONT = 5;
  TEXT_APPLYFONTSIZE = 6;
  TEXT_COLOR = 7;
  TEXT_BACKCOLOR = 8;
  //Parameters for ApplyParaStyleConversion
  PARA_ALIGNMENT_LEFT = 1;
  PARA_ALIGNMENT_CENTER = 2;
  PARA_ALIGNMENT_RIGHT = 3;
  PARA_COLOR = 4;

type
  TcbTrayIcon = class(TTrayIcon)
  private
    FBalloonShowing: Boolean;
  protected
    procedure WindowProc(var Message: TMessage); override;
  public
    constructor Create(Owner: TComponent); override;
  public
    property Data;
  published
    property BalloonShowing: Boolean read FBalloonShowing;
  end;


  THrClientForm = class(TdxRibbonForm)
  private
    FHrClientFormKind: THrClientFormKind;
  public
    constructor Create(AOwner: TComponent; const AFormKind: THrClientFormKind); reintroduce;
    procedure ChangeColorSchema(const AColorName: String);
    property HrClientFormKind: THrClientFormKind read FHrClientFormKind;
  end;


function GetClientUserStatusText(const AUserStatus: TClinetUserStatus): String; overload;
function GetClientUserStatusText(const AUserStatus: Integer): String; overload;

implementation


{ ---------------------------------------------------------------------------- }

function GetClientUserStatusText(const AUserStatus: TClinetUserStatus): String;
begin
  Result := 'Â÷½u';
  if ( AUserStatus in [usOnline, usBusy]) then
    Result := '½u¤W';
end;

{ ---------------------------------------------------------------------------- }

function GetClientUserStatusText(const AUserStatus: Integer): String; overload;
begin
  Result := GetClientUserStatusText( TClinetUserStatus( AUserStatus ) );
end;

{ ---------------------------------------------------------------------------- }

{ TcbTrayIcon }

constructor TcbTrayIcon.Create(Owner: TComponent);
begin
  inherited Create( Owner );
  FBalloonShowing := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TcbTrayIcon.WindowProc(var Message: TMessage);
begin
  case Message.Msg of
    WM_SYSTEM_TRAY_MESSAGE:
    begin
      case Message.lParam of
        NIN_BALLOONSHOW, NIN_BALLOONUSERCLICK:
            FBalloonShowing := True;
        NIN_BALLOONHIDE, NIN_BALLOONTIMEOUT:
            FBalloonShowing := False;
      end;
    end;
  end;
  inherited WindowProc( Message );
end;

{ ---------------------------------------------------------------------------- }

{ THrClientForm }

constructor THrClientForm.Create(AOwner: TComponent; const AFormKind: THrClientFormKind);
begin
  FHrClientFormKind := AFormKind;
  inherited Create( AOwner );
end;

{ ---------------------------------------------------------------------------- }

procedure THrClientForm.ChangeColorSchema(const AColorName: String);
var
  aRibbonControl: TdxRibbon;
  aIndex: Integer;
begin
  aRibbonControl := nil;
  for aIndex := 0 to Self.ControlCount - 1 do
  begin
    if ( Self.Controls[aIndex] is TdxRibbon ) then
    begin
      aRibbonControl := TdxRibbon( Self.Controls[aIndex] );
      Break;
    end;
  end;
  if Assigned( aRibbonControl ) then
    aRibbonControl.ColorSchemeName := AColorName;
end;

{ ---------------------------------------------------------------------------- }

end.
