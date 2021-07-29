unit CbUtils;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms, IdStackWindows;


{ --------------------------------------------------------------------- }
{ 顯示詢問 MessageBox                                                   }
{ --------------------------------------------------------------------- }
function ConfirmMsg(const aMsg: String): Boolean;


{ --------------------------------------------------------------------- }
{ 顯示警告 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure WarningMsg(const aMsg: String);


{ --------------------------------------------------------------------- }
{ 顯示錯誤 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure ErrorMsg(const aMsg: String);


{ --------------------------------------------------------------------- }
{ 顯示通知 MessageBox                                                   }
{ --------------------------------------------------------------------- }
procedure InfoMsg(const aMsg: String);


{ --------------------------------------------------------------------- }
{ 計算檔案大小, 回傳值單位 MB                                           }
{ --------------------------------------------------------------------- }
function CalcFileSizeInMB(const aFullFileName: String): Double;


{ --------------------------------------------------------------------- }
{ 計算檔案大小, 回傳值單位 Byte                                         }
{ --------------------------------------------------------------------- }
function CalcFileSizeInByte(const aFullFileName: String): Cardinal;


{ --------------------------------------------------------------------- }
{ 本機電腦名稱                                                          }
{ --------------------------------------------------------------------- }
function LocalComputerName: String;


{ --------------------------------------------------------------------- }
{ 本機IP位址                                                            }
{ --------------------------------------------------------------------- }
function LocalIpAddress: TStringList;


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
  const aPadChr: Char): String; overload;

function Lpad(aValue: Integer; const aPadLen: Integer;
  const aPadChr: Char): String; overload;


{ --------------------------------------------------------------------- }
{ Rpad                                                                  }
{   傳進的資料, 從右邊補足位數字元                                      }
{ --------------------------------------------------------------------- }
function Rpad(aValue: String; const aPadLen: Integer;
  const aPadChr: Char): String; overload;

function Rpad(aValue: Integer; const aPadLen: Integer;
  const aPadChr: Char): String; overload;


{ --------------------------------------------------------------------- }
{ StringToCharArray                                                     }
{   將字串複製一份進 CharArry                                           }
{ --------------------------------------------------------------------- }
procedure StringToCharArray(const aSource: String; const aSize: Integer;
  var aDest);

implementation

type
  TDirection = ( dtLeft, dtRight );

{ ---------------------------------------------------------------------------- }

function ConfirmMsg(const aMsg: String): Boolean;
begin
  Result := False;
  if aMsg <> '' then
    Result := ( Application.MessageBox( PChar( aMsg ), '確認', MB_OKCANCEL +
      MB_ICONQUESTION ) = ID_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure ErrorMsg(const aMsg: String);
begin
  if aMsg <> '' then
    Application.MessageBox( PChar( aMsg ), '錯誤', MB_OK +  MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure WarningMsg(const aMsg: String);
begin
  if aMsg <> '' then
    Application.MessageBox( PChar( aMsg ), '警告', MB_OK +  MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure InfoMsg(const aMsg: String);
begin
  if aMsg <> '' then
   Application.MessageBox( PChar( aMsg ), '訊息', MB_OK + MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInMB(const aFullFileName: String): Double;
var
  aSize: Cardinal;
begin
  aSize := CalcFileSizeInByte( aFullFileName );
  Result := StrToFloat( Format( '%.2f', [aSize / 1000 / 1000] ) );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInByte(const aFullFileName: String): Cardinal;
var
  aFileWnd: Integer;
begin
  Result := 0;
  if not FileExists( aFullFileName ) then Exit;
  aFileWnd := FileOpen( aFullFileName, fmOpenRead );
  try
    if aFileWnd >= 0 then Result := GetFileSize( aFileWnd, nil );
  finally
    if aFileWnd >= 0 then FileClose( aFileWnd );
  end;
end;

{ ---------------------------------------------------------------------------- }

function LocalComputerName: String;
var
  aSize: LongWord;
begin
  SetLength( Result, MAX_COMPUTERNAME_LENGTH + 1 );
  aSize := Length( Result );
  if Windows.GetComputerName( @Result[1], aSize ) then
    SetLength( Result, aSize );
end;

{ ---------------------------------------------------------------------------- }

function LocalIpAddress: TStringList;
var
  aWindStack: TIdStackWindows;
begin
  aWindStack := TIdStackWindows.Create;
  try
    Result := TStringList.Create;
    Result.Assign( aWindStack.LocalAddresses );
  finally
    aWindStack.Free;
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
  aCount: Integer;
begin
  Result := InValue;
  aCount := PadLen - Length( InValue );
  if aCount > 0 then
  begin
    if aDirection = dtLeft then
      Result := ( StringOfChar( PadChr, aCount ) + Result )
    else
      Result := ( Result + StringOfChar( PadChr, aCount ) );  
  end;
end;

{ ---------------------------------------------------------------------------- }

function Lpad(aValue: String; const aPadLen: Integer; const aPadChr: Char): String;
begin
  Result := PadChar( dtLeft, aValue, aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

function Lpad(aValue: Integer; const aPadLen: Integer;
  const aPadChr: Char): String; overload;
begin
  Result := PadChar( dtLeft, IntToStr( aValue ), aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

function Rpad(aValue: String; const aPadLen: Integer; const aPadChr: Char): String;
begin
  Result := PadChar( dtRight, aValue, aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

function Rpad(aValue: Integer; const aPadLen: Integer;
  const aPadChr: Char): String; overload;
begin
  Result := PadChar( dtRight, IntToStr( aValue ), aPadLen, aPadChr );
end;

{ ---------------------------------------------------------------------------- }

procedure StringToCharArray(const aSource: String; const aSize: Integer;
  var aDest);
var
  aIndex: Integer;
  aChr: PChar;
begin
  aChr := PChar( @aDest );
  for aIndex := 1 to Length( aSource ) do
  begin
    if aIndex <= aSize then
      aChr[aIndex-1] := aSource[aIndex]
    else
      Break;  
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
