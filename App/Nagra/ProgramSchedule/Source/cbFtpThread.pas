unit cbFtpThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbAppClass, cbClass;

type
  TFtpThread = class(TMessageQueueThread)
  private
    { Private declarations }
    FMsgSubject: TMessageSubject;
    FParam: TAppParam;
    FPeriod: TAppPeriod;
    FLastExecute: TDateTime;
    procedure SetParam(const Value: TAppParam);
    procedure SetPeriod(const Value: TAppPeriod);
    function CheckPeriodTimeToReach: Boolean;
  protected
    procedure Execute; override;
  public
    constructor Create;
    destructor Destroy; override;
    property MessageSubject: TMessageSubject read FMsgSubject;
    property Param: TAppParam read FParam write SetParam;
    property Perido: TAppPeriod read FPeriod write SetPeriod;
  end;

implementation

uses DateUtils;

{ TFtpThread }

{ ---------------------------------------------------------------------------- }

constructor TFtpThread.Create;
begin
  inherited Create( True );
  FMsgSubject := TMessageSubject.Create;
  FParam := TAppParam.Create;
  FPeriod := TAppPeriod.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TFtpThread.Destroy;
begin
  FParam.Free;
  FPeriod.Free;
  FMsgSubject.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.SetParam(const Value: TAppParam);
begin
  FParam.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.SetPeriod(const Value: TAppPeriod);
begin
  FPeriod.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.Execute;
begin
  WaitForPlaySignal;
  FLastExecute := 0;
  while not Self.Terminated do
  begin
    Sleep( 100 );
    try
      if CheckPeriodTimeToReach then
      begin
        FLastExecute := Now;
//        if XmlFileDetected then
//        begin
//          if Self.Terminated then Break;
//          XmlFileMove;
//        end;
      end;
      Sleep( 100 );
      if Self.Terminated then Break;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( '【檔案下載】執行緒發生錯誤, 原因:%s。',
          [E.Message] ) );
      end;
    end;
    Sleep( 300 );
  end;
  WaitForStopSignal;
end;

{ ---------------------------------------------------------------------------- }

function TFtpThread.CheckPeriodTimeToReach: Boolean;
var
  aIndex: Integer;
  aHour: Word;
begin
  Result := False;
  if ( FPeriod.PeriodType = '1' ) then
  begin
    if ( FLastExecute = 0 ) then
    begin
      Result := True;
      Sleep( 1000 * 10 );
    end else
      Result := ( MinutesBetween( Now, FLastExecute ) >= FPeriod.PeriodMinute );
  end else
  begin
    FPeriod.CheckExpire;
    if ( DayOf( FLastExecute ) <> DayOf( Now ) ) then
      FPeriod.ResetExecute;
    aHour := HourOf( Now );
    for aIndex := 0 to FPeriod.ClockCount - 1 do
    begin
      if ( FPeriod.Clocks[aIndex].Expire ) then Continue;
      if ( not FPeriod.Clocks[aIndex].IsSet ) then Continue;
      if ( FPeriod.Clocks[aIndex].Execute ) then Continue;
      Result := ( aHour >= aIndex );
      if ( Result ) then
      begin
        FPeriod.Clocks[aIndex].Execute := True;
        Break;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
