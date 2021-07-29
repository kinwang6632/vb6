{ ---------------------------------------------------------------------------- }
{                                                                              }
{ PC HOME ONLINE Copyright (c) 2002-2003                                       }
{                                                                              }
{ Project: PC home online 網路家庭( EC2 ) ERP Program                          }
{ Unit: 自訂物件                                                               }
{ Author: Bug                                                                  }
{ Date: 2003/07/07                                                             }
{                                                                              }
{ ---------------------------------------------------------------------------- }

unit ApClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, ApUtilis;

type
  TApInfo = class(TObject)
  private
    { Private declarations }
    FUserID: String;
    FUserName: String;
    FPassWord: String;
    FUniqueID: Integer;
    FUseIME: String;
    FAdministrator: Boolean;
    FIdentity: String;
    function GetLoginIP: String;
    function GetMachineName: String;
    function GetSysDate: String;
    function GetSysTime: String;
    function GetOracleSysDate: String;
  public
    { Public declarations }
    constructor Create;
    destructor Destroy; override;
    property LoginID: String read FUserID write FUserID;
    property LoginName: String read FUserName write FUserName;
    property LoginPass: String read FPassWord write FPassWord;
    property LoginComputer: String read GetMachineName;
    property LoginIP: String read GetLoginIP;
    property LoginUniqueID: Integer read FUniqueID write FUniqueID;
    property SysDate: String read GetSysDate;
    property SysTime: String read GetSysTime;
    property UseIME: String read FUseIME write FUseIME;
    property Administrator: Boolean read FAdministrator write FAdministrator;
    property Identity: String read FIdentity write FIdentity;
    property OracleSysDate: String read GetOracleSysDate;
  end;

  { 程式權限 }

  PAuthorized = ^TAuthorized;

  TAuthorized = record
    Insert: Boolean;
    Update: Boolean;
    Delete: Boolean;
    Print: Boolean;
    Query: Boolean;
  end;  
  
implementation

{ TApInfo }

uses Sockets, Main, DM;

{ ---------------------------------------------------------------------------- }

constructor TApInfo.Create;
begin
  Inherited Create;
  FUniqueID := -1;
  FIdentity := '';
end;

{ ---------------------------------------------------------------------------- }

destructor TApInfo.Destroy;
begin
  Inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TApInfo.GetLoginIP: String;
var
  ASocket: TIpSocket;
begin
  Result := '127.0.0.1';
  ASocket := TIpSocket.Create( nil );
  try
    Result := ASocket.LocalHostAddr;
  finally
    FreeAndNil( ASocket );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TApInfo.GetMachineName: String;
var
  ALen: LongWord;
begin
  SetLength( Result, MAX_COMPUTERNAME_LENGTH + 1 );
  ALen := Length( Result );
  if Windows.GetComputerName( @Result[1], ALen ) then
    SetLength( Result, ALen );
end;

{ ---------------------------------------------------------------------------- }

function TApInfo.GetOracleSysDate: String;
begin
  frmDM.CdsOracleDate.Active := False;
  frmDM.CdsOracleDate.CommandText :=
    ' SELECT TO_CHAR( SYSDATE, ''YYYYMMDD'' ) AS ODATE FROM DUAL ';
  try
    frmDM.CdsOracleDate.Active := True;
    Result := frmDM.CdsOracleDate.FieldByName( 'ODATE' ).AsString;
  except
    Result := '';
  end;
  frmDM.CdsOracleDate.Active := False;
end;

{ ---------------------------------------------------------------------------- }

function TApInfo.GetSysDate: String;
begin
  Result := ApUtilis.GetSysDate( True );
end;

{ ---------------------------------------------------------------------------- }

function TApInfo.GetSysTime: String;
begin
  Result := ApUtilis.GetSysTime;
end;

{ ---------------------------------------------------------------------------- }

end.
