unit cbStyleModule;

interface

uses
  Windows, SysUtils, Classes, ImgList, Controls, cxContainer, cxEdit, cxStyles,
  cxClasses, cxTL, cbAppClass;

type
  TStyleModule = class(TDataModule)
    BarImageList: TImageList;
    SoTreeImageList: TImageList;
    CmdTreeImageList: TImageList;
    CommonStyle: TcxEditStyleController;
    MsgImageList: TImageList;
    MsgStyle: TcxEditStyleController;
    DockPanelImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    TreeListStyleSheetConsoleBlack: TcxTreeListStyleSheet;
    cxStyle1: TcxStyle;
    cxStyle2: TcxStyle;
    cxStyle3: TcxStyle;
    cxStyle4: TcxStyle;
    cxStyle5: TcxStyle;
    cxStyle6: TcxStyle;
    cxStyle7: TcxStyle;
    cxStyle8: TcxStyle;
    cxStyle9: TcxStyle;
    cxStyle10: TcxStyle;
    cxStyle11: TcxStyle;
    cxStyle12: TcxStyle;
    cxStyle13: TcxStyle;
    cxStyle14: TcxStyle;
    cxStyle15: TcxStyle;
    ConfigImageList: TImageList;
    cxEditStyle: TcxEditStyleController;
    cxStyle16: TcxStyle;
  private
    { Private declarations }
  public
    { Public declarations }
    function GetMsgImageIndex(const aFlag: Integer): Integer;
    function GetSoImageIndex(const aSo: TSo): Integer;
    function GetSoStateIndex(const aSo: TSo): Integer;
    function GetCmdImageIndex(const aCmd: TSendDvn): Integer;
    function OperatorImageIndex: Integer;
    function LowCmdImageIndex: Integer;
    function ClockSendImageIndex: Integer;
    function ClockRecvImageIndex: Integer;
    function SocketConnectImageIndex: Integer;
    function SocketDisconnectImageIndex: Integer;
    function SocketWarningImageIndex: Integer;
    function SocketNullImageIndex: Integer;
  end;


var StyleModule: TStyleModule;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

{ TStyleModule }

function TStyleModule.GetMsgImageIndex(const aFlag: Integer): Integer;
begin
  case aFlag of
    MB_ICONINFORMATION: Result := 0;
    MB_OK: Result := 1;
    MB_ICONERROR: Result := 2;
    MB_ICONWARNING: Result := 3;
  else
    Result := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetSoImageIndex(const aSo: TSo): Integer;
begin
  Result := 1;
  if ( aSo.Selected ) then
  begin
    case aSo.DbState of
      dbClose: Result := 2;
      dbOpen, dbActive: Result := 3;
      dbError: Result := 4;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetSoStateIndex(const aSo: TSo): Integer;
begin
  Result := 5;
  if ( aSo.Selected ) then
  begin
    case aSO.DbState of
      dbClose: Result := 5;
      dbOpen, dbError: Result := 6;
      dbActive: Result := 7;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetCmdImageIndex(const aCmd: TSendDvn): Integer;
begin
  if ( aCmd.CmdType = ctHigh ) then
  begin
    Result := 8;
    if ( aCmd.IsBatch ) then Result := 9;
  end else
  begin
    Result := 1;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.OperatorImageIndex: Integer;
begin
  Result := 7;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.LowCmdImageIndex: Integer;
begin
  Result := 13;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.ClockSendImageIndex: Integer;
begin
  Result := 14;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.ClockRecvImageIndex: Integer;
begin
  Result := 15;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketConnectImageIndex: Integer;
begin
  Result := 17;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketDisconnectImageIndex: Integer;
begin
  Result := 18;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketWarningImageIndex: Integer;
begin
  Result := 24;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketNullImageIndex: Integer;
begin
  Result := 26;
end;

{ ---------------------------------------------------------------------------- }

end.
