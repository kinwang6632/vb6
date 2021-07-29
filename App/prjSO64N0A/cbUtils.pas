unit cbUtils;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms;


{ --------------------------------------------------------------------- }
{ 顯示詢問 MessageBox                                                   }
{ --------------------------------------------------------------------- }
function ConfirmMsg(const AMsg: String): Boolean;


{ --------------------------------------------------------------------- }
{ 顯示警告 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure WarningMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ 顯示錯誤 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure ErrorMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ 顯示通知 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure InfoMsg(const AMsg: String);


{ --------------------------------------------------------------------- }
{ 計算檔案大小, 回傳值單位 MB                                           }
{ --------------------------------------------------------------------- }
function CalcFileSizeInMB(const AFullFileName: String): Double;


{ --------------------------------------------------------------------- }
{ 計算檔案大小, 回傳值單位 Byte                                         }
{ --------------------------------------------------------------------- }
function CalcFileSizeInByte(const AFullFileName: String): Cardinal;


{ --------------------------------------------------------------------- }
{ 延遲函數                                                              }
{ --------------------------------------------------------------------- }
procedure Delay(const AValue: Cardinal);


{ --------------------------------------------------------------------- }
{ Nvl                                                                   }
{   當傳進的資料 InValue 為空值時, 則回傳 NullValue                     }
{ --------------------------------------------------------------------- }
function Nvl(InValue: String; NullValue: String): String; overload;


{ --------------------------------------------------------------------- }
{ Nvl                                                                   }
{   當傳進的資料 InValue 為空值時, 則回傳 NullValue                     }
{ --------------------------------------------------------------------- }
function Nvl(InValue: String; NullValue: Integer): Integer; overload;


{ --------------------------------------------------------------------- }
{ Lpad                                                                  }
{   傳進的資料, 從左邊補足位數字元                                      }
{ --------------------------------------------------------------------- }
function Lpad(aValue: String; const aPadLen: Integer;
  const aPadChr: Char): String;

{ --------------------------------------------------------------------- }
{ Rpad                                                                  }
{   傳進的資料, 從右邊補足位數字元                                      }
{ --------------------------------------------------------------------- }
function Rpad(aValue: String; const aPadLen: Integer;
  const aPadChr: Char): String;


{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   解出並回傳使用分隔字元分隔的值, Ex: 123,456,789                     }
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
    Result := ( Application.MessageBox( PChar( AMsg ), '確認', MB_OKCANCEL +
      MB_ICONQUESTION ) = ID_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure ErrorMsg(const AMsg: String);
begin
  if AMsg <> '' then
    Application.MessageBox( PChar( AMsg ), '錯誤', MB_OK +  MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure WarningMsg(const AMsg: String);
begin
  if AMsg <> '' then
    Application.MessageBox( PChar( AMsg ), '警告', MB_OK +  MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure InfoMsg(const AMsg: String);
begin
  if AMsg <> '' then
   Application.MessageBox( PChar( AMsg ), '訊息', MB_OK + MB_ICONINFORMATION );
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



