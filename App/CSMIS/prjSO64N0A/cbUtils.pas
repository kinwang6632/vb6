unit cbUtils;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms;


{ --------------------------------------------------------------------- }
{ ��ܸ߰� MessageBox                                                   }
{ --------------------------------------------------------------------- }
function ConfirmMsg(const AMsg: String): Boolean;


{ --------------------------------------------------------------------- }
{ ���ĵ�i MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure WarningMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ ��ܿ��~ MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure ErrorMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ ��ܳq�� MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure InfoMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ �p���ɮפj�p, �^�ǭȳ�� MB                                           }
{ --------------------------------------------------------------------- }
function CalcFileSizeInMB(const AFullFileName: String): Double;


{ --------------------------------------------------------------------- }
{ �p���ɮפj�p, �^�ǭȳ�� Byte                                         }
{ --------------------------------------------------------------------- }
function CalcFileSizeInByte(const AFullFileName: String): Cardinal;


{ --------------------------------------------------------------------- }
{ ������                                                              }
{ --------------------------------------------------------------------- }
procedure Delay(const AValue: Cardinal);


{ --------------------------------------------------------------------- }
{ Nvl                                                                   }
{   ��Ƕi����� InValue ���ŭȮ�, �h�^�� NullValue                     }
{ --------------------------------------------------------------------- }
function Nvl(InValue: String; NullValue: String): String; overload;


{ --------------------------------------------------------------------- }
{ Nvl                                                                   }
{   ��Ƕi����� InValue ���ŭȮ�, �h�^�� NullValue                     }
{ --------------------------------------------------------------------- }
function Nvl(InValue: String; NullValue: Integer): Integer; overload;


{ --------------------------------------------------------------------- }
{ Lpad                                                                  }
{   �Ƕi�����, �q����ɨ���Ʀr��                                      }
{ --------------------------------------------------------------------- }
function Lpad(aValue: String; const aPadLen: Integer;
  const aPadChr: Char): String;

{ --------------------------------------------------------------------- }
{ Rpad                                                                  }
{   �Ƕi�����, �q�k��ɨ���Ʀr��                                      }
{ --------------------------------------------------------------------- }
function Rpad(aValue: String; const aPadLen: Integer;
  const aPadChr: Char): String;


{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   �ѥX�æ^�ǨϥΤ��j�r�����j����, Ex: 123,456,789                     }
{ --------------------------------------------------------------------- }
function ExtractValue(var aValue: String; aSeparator: String = ','): String;


implementation

type
  TDirection = ( dtLeft, dtRight );

{ ---------------------------------------------------------------------------- }

function ConfirmMsg(const AMsg: String): Boolean;
begin
  Result := False;
  if AMsg <> '' then
    Result := ( Application.MessageBox( PChar( AMsg ), '�T�{', MB_OKCANCEL +
      MB_ICONQUESTION ) = ID_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure ErrorMsg(const AMsg: String);
begin
  if AMsg <> '' then
    Application.MessageBox( PChar( AMsg ), '���~', MB_OK +  MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure WarningMsg(const AMsg: String);
begin
  if AMsg <> '' then
    Application.MessageBox( PChar( AMsg ), 'ĵ�i', MB_OK +  MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure InfoMsg(const AMsg: String);
begin
  if AMsg <> '' then
   Application.MessageBox( PChar( AMsg ), '�T��', MB_OK + MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInMB(const AFullFileName: String): Double;
var
  ASize: Cardinal;
begin
  ASize := CalcFileSizeInByte( AFullFileName );
  Result := StrToFloat( Format( '%.2f', [ASize / 1000 / 1000] ) );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInByte(const AFullFileName: String): Cardinal;
var
  AFileWnd: Integer;
begin
  Result := 0;
  if not FileExists( AFullFileName ) then Exit;
  AFileWnd := FileOpen( AFullFileName, fmOpenRead );
  try
    if AFileWnd >= 0 then Result := GetFileSize( AFileWnd, nil );
  finally
    if AFileWnd >= 0 then FileClose( AFileWnd );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure Delay(const AValue: Cardinal);
var
  ANow: Cardinal;
begin
  ANow :=  GetTickCount;
  repeat
    Application.ProcessMessages;
  until ( GetTickCount - ANow ) >= AValue;
end;

{ ---------------------------------------------------------------------------- }

function Nvl(InValue: String; NullValue: String): String; overload;
begin
  Result := NullValue;
  if InValue <> '' then Result := InValue;
end;

{ ---------------------------------------------------------------------------- }

function Nvl(InValue: String; NullValue: Integer): Integer; overload;
begin
  Result := NullValue;
  if Trim( InValue ) <> '' then Result := StrToInt( InValue );
end;

{ ---------------------------------------------------------------------------- }

function PadChar(const aDirection: TDirection; InValue: String;
  const PadLen: Integer; const PadChr: Char): String;
var
  AIndex: Integer;
begin
  Result := InValue;
  for AIndex := 1 to PadLen - Length( InValue ) do
  begin
    case aDirection of
      dtLeft: Result := ( PadChr + Result );
      dtRight : Result := ( Result + PadChr );
    else
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function Lpad(aValue: String; const aPadLen: Integer; const aPadChr: Char): String;
begin
  Result := PadChar( dtLeft, aValue, aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

function Rpad(aValue: String; const aPadLen: Integer; const aPadChr: Char): String;
begin
  Result := PadChar( dtRight, aValue, aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

function ExtractValue(var aValue: String; aSeparator: String = ','): String;
var
  aPos: Integer;
begin
  aPos := AnsiPos( aSeparator, aValue );
  if aPos = 0 then
  begin
    Result := aValue;
    aValue := EmptyStr;
  end else
  begin
    Result := Copy( aValue, 1, aPos - 1 );
    Delete( aValue, 1, aPos );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.



