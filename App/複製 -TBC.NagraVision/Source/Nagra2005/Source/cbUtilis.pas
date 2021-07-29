unit cbUtilis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms, IniFiles
  {$IFDEF APPDEBUG}, CodeSiteLogging {$ENDIF};



{ --------------------------------------------------------------------- }
{ 顯示詢問 MessageBox                                                   }
{ --------------------------------------------------------------------- }
function ConfirmMsg(const AMsg: String): Boolean;
function ConfirmMsgEx(const AMsg: String): Boolean;


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
{ 計算檔案大小, 回傳值單位 KB                                           }
{ --------------------------------------------------------------------- }
function CalcFileSizeInKB(const AFullFileName: String): Double;


{ --------------------------------------------------------------------- }
{ 計算檔案大小, 回傳值單位 Byte                                         }
{ --------------------------------------------------------------------- }
function CalcFileSizeInByte(const AFullFileName: String): Cardinal;



{ --------------------------------------------------------------------- }
{ 延遲函數                                                              }
{   AValue: 毫秒, Ex: 3000 = 3 秒                                       }
{ --------------------------------------------------------------------- }
procedure Delay(const AValue: Cardinal);



{ --------------------------------------------------------------------- }
{ 去掉指定的字串                                                        }
{ --------------------------------------------------------------------- }
function TrimString(ASource, ATarget: String): String;


{ --------------------------------------------------------------------- }
{ 取系統日期                                                            }
{     AType: dtROC  : 民國年                                            }
{            dtEng  : 西元年                                            }
{ --------------------------------------------------------------------- }
function GetSysDate(ADateSeparator: Boolean = False): String;


{ --------------------------------------------------------------------- }
{ 取系統時間                                                            }
{ --------------------------------------------------------------------- }
function GetSysTime(ATimeSeparator: Boolean = False): String;


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
{ DateTextIsValid                                                       }
{ DateTextIsValidEx                                                     }
{  驗證輸入的日期(西元或民國年)是否合法, Ex: 94/1/1 or 2005/1/1         }
{ --------------------------------------------------------------------- }
function DateTextIsValid(const aDateText: String): Boolean;
function DateTextIsValidEx(var aDateText: String): Boolean;


{ --------------------------------------------------------------------- }
{ PadDateText                                                           }
{ PadDateText2                                                          }
{   補足日期字串 Ex: 2003/1/2 ---> 2003/01/02                           }
{                Ex: 200312   ---> 20031200                             }
{ --------------------------------------------------------------------- }
function PadDateText(const ADateText: String): String;
function PadDateText2(const ADateText: String): String;


{ --------------------------------------------------------------------- }
{ DateConvert                                                           }
{   將西元年轉換成民國年                                                }
{ --------------------------------------------------------------------- }
function DateConvert(aDate: TDateTime; aDateSeparator: Boolean = False): String; overload;
function DateConvert(aDate: String; aDateSeparator: Boolean = False): String; overload;


{ --------------------------------------------------------------------- }
{ FormatMaskText                                                        }
{                                                                       }
{ --------------------------------------------------------------------- }
function FormatMaskText(const EditMask: String; const Value: String): String;



{ --------------------------------------------------------------------- }
{ FileNameWithoutExt                                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function FileNameWithoutExt(const FileName: String): String;


{ --------------------------------------------------------------------- }
{ CbRoundTo                                                             }
{    標準的四捨五入 , Ex: 123.45 --> 123.5                              }
{ --------------------------------------------------------------------- }
function CbRoundTo(AValue: Extended; ADigit: Word): Extended;


{ --------------------------------------------------------------------- }
{ BinToInt                                                              }
{    二進制轉數值, Ex: 100100 --> 36                                    }
{ --------------------------------------------------------------------- }
function BinToInt(const ABinary: String): Integer;


{ --------------------------------------------------------------------- }
{ 營利事業統一編號檢查                                                  }
{                                                                       }
{ --------------------------------------------------------------------- }
//function TaxtNoValidate(const ANo: String): Boolean;



{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   解出並回傳使用分隔字元分隔的值, Ex: 123,456,789                     }
{                                                                       }
{ --------------------------------------------------------------------- }
function ExtractValue(var AValue: String; ASeparator: String = ','): String; overload;

{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   解出並回傳指定位置的值, Ex: 123,456,789                             }
{   AIndex: 指定的位置, 從 1 開始                                       }
{ --------------------------------------------------------------------- }
function ExtractValue(AIndex: Integer; AValue: String; ASeparator: String = ','): String; overload;



{ --------------------------------------------------------------------- }
{ QuotedValue                                                           }
{   將每個字元加上包夾字元, Ex: 123,456,789 --> '123','456','789'       }
{                                                                       }
{ --------------------------------------------------------------------- }
function QuotedValue(AValue: String; AQuoted: String = ''''; ASeparator: String = ','): String;


{ --------------------------------------------------------------------- }
{ ValuePos                                                              }
{   取得符合該字串的位置, Ex: ACompare=456, AValue=123,456,789          }
{   回傳 2                                                              }
{ --------------------------------------------------------------------- }
function ValuePos(ACompare, AValue: String; ASeparator: String = ','): Integer;


{ --------------------------------------------------------------------- }
{ ValueExists                                                           }
{   是否有符合該字串的, Ex: ACompare=456, AValue=123,456,789            }
{   回傳 True                                                           }
{ --------------------------------------------------------------------- }
function ValueExists(ACompare, AValue: String; ASeparator: String = ','): Boolean;


{ --------------------------------------------------------------------- }
{ DeleteValue                                                           }
{   去掉符合該字串的字串, Ex: ACompare=456, AValue=123,456,789          }
{   回傳 123,789                                                        }
{ --------------------------------------------------------------------- }
function DeleteValue(ADelete, AValue: String; ASeparator: String = ','): String;


{ --------------------------------------------------------------------- }
{ AddValue                                                              }
{   增加字串的值, Ex: AAdd=456, AValue=123,789                          }
{   回傳 123,789,456                                                    }
{ --------------------------------------------------------------------- }
function AddValue(AAdd, AValue: String; ASeparator: String = ','): String;



{ --------------------------------------------------------------------- }
{ TrimChar                                                              }
{   刪除出現在來源字串中指定的字元, Ex: 2005/1_/1_                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function TrimChar(const AValue: String; const AParams: array of const): String;



{ --------------------------------------------------------------------- }
{ FormatDateTimeString                                                  }
{   格式化字串成日期分隔字元 Ex: 20051201180030 --> 2005/12/01 18:00:30 }
{                                                                       }
{ --------------------------------------------------------------------- }
function FormatDateTimeString(aSource: String): String;



{ --------------------------------------------------------------------- }
{ UTCDateTimeToLocalTime                                                }
{   將格林威治日期時間轉換成台灣時間                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function UTCDateTimeToLocalDateTime(AUTC: TDateTime): TDateTime; overload;
function UTCDateTimeToLocalDateTime(AUTC: String): TDateTime; overload;


{ --------------------------------------------------------------------- }
{ UTCDateTimeToLocalTime                                                }
{   將台灣時間轉換成格林威治日期時間                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function LocalDateTimeToUTCDateTime(ALocal: TDateTime): TDateTime; overload;
function LocalDateTimeToUTCDateTime(ALocal: String): TDateTime; overload;


{ --------------------------------------------------------------------- }
{ IIF                                                                   }
{   傳回符合條件的值                                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function IIF(ATest: Boolean; const ATrue, AFalse: String): String; overload;
function IIF(ATest: Boolean; const ATrue, AFalse: Integer): Integer; overload;
function IIF(ATest: Boolean; const ATrue, AFalse: Double): Double; overload;



{ --------------------------------------------------------------------- }
{ GetComputerName                                                       }
{   取本機的電腦名稱                                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetComputerName: String;


{ --------------------------------------------------------------------- }
{ GetComputerName                                                       }
{   取本機的IP位址                                                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetIpAddress: String;


{ --------------------------------------------------------------------- }
{ IsRunInTerminalServer                                                 }
{   是否執行在微軟 Terminal Server 上                                   }
{                                                                       }
{ --------------------------------------------------------------------- }
function IsRunInTerminalServer: Boolean;


{ --------------------------------------------------------------------- }
{ GetTerminalServerInfo                                                 }
{   取本機在微軟 Terminal Server 上的資訊                               }
{                                                                       }
{ --------------------------------------------------------------------- }
procedure GetTerminalServerInfo(var ATermId, ATermIp, ATermName, ATermPC, ATermState: String);


{ --------------------------------------------------------------------- }
{ PreviousInstanceExists                                                }
{   應用程式否已經執行中                                                }
{                                                                       }
{ --------------------------------------------------------------------- }
function PreviousInstanceExists(ATitle: String; const AShow: Boolean;
  var ASemaphore: THandle): Boolean;


{ --------------------------------------------------------------------- }
{ IsProcessExists                                                       }
{   應用程式是否已經執行中                                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function IsProcessExists(const AProcessName: String): Boolean;


{ --------------------------------------------------------------------- }
{ GetProcessId                                                          }
{   取得指定應用程式的 Process Handle                                   }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetProcessId(const AProcessName: String): Cardinal;


{ --------------------------------------------------------------------- }
{ RetrieveApplicationMemorySize                                         }
{   回收執行中的應用程式記憶體使用量                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
procedure DecreaseProcessMemorySize(AProcessId: Integer = 0);


{ --------------------------------------------------------------------- }
{ FileExists                                                            }
{   檢測指定的 IniFile 是否存在,若不存在則建立                          }
{                                                                       }
{ --------------------------------------------------------------------- }
function FileNotExistsAndCreate(AFileName: String; ACreateIfNotExists: Boolean): Boolean;


{ --------------------------------------------------------------------- }
{ FileExists                                                            }
{   檢測指定的 IniFile 是否存在,若不存在則建立                          }
{                                                                       }
{ --------------------------------------------------------------------- }


implementation

uses DateUtils, MaskUtils, IdStackWindows, WtsApi32;

type
  TDirection = ( dtLeft, dtRight );
  TFormatType = ( ftDateTime, ftString );


{ ---------------------------------------------------------------------------- }

function FormatString(aSource: String; aFormatType: TFormatType): String;
begin
  Result := aSource;
  if ( aFormatType = ftDateTime ) then
  begin
    Result :=
      Copy( aSource, 1, 4 ) + '/' +
      Copy( aSource, 5, 2 ) + '/' +
      Copy( aSource, 7, 2 ) + ' ' +
      Copy( aSource, 9, 2 ) + ':' +
      Copy( aSource, 11, 2 ) + ':' +
      Copy( aSource, 13, 2 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function ConfirmMsg(const AMsg: String): Boolean;
begin
  Result := False;
  if ( AMsg <> EmptyStr ) then
    Result := ( Application.MessageBox( PChar( AMsg ), '確認', MB_OKCANCEL +
      MB_ICONQUESTION ) = ID_OK );
end;

{ ---------------------------------------------------------------------------- }

function ConfirmMsgEx(const AMsg: String): Boolean;
begin
  Result := False;
  if ( AMsg <> EmptyStr ) then
    Result := ( Application.MessageBox( PChar( AMsg ), '確認', MB_OKCANCEL +
      MB_ICONQUESTION + MB_DEFBUTTON2	) = ID_OK );
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
  Result := StrToFloat( Format( '%.2f', [ASize / 1024 / 1024] ) );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInKB(const AFullFileName: String): Double;
var
  ASize: Cardinal;
begin
  ASize := CalcFileSizeInByte( AFullFileName );
  Result := StrToFloat( Format( '%.2f', [ASize / 1024] ) );
end;

{ ---------------------------------------------------------------------------- }

function CalcFileSizeInByte(const AFullFileName: String): Cardinal;
var
  AFileWnd: Integer;
begin
  Result := 0;
  if not FileExists( AFullFileName ) then Exit;
  AFileWnd := FileOpen( AFullFileName, fmOpenRead );
  if AFileWnd >= 0 then
  begin
    Result := GetFileSize( AFileWnd, nil );
    FileClose( AFileWnd );
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

function TrimString(ASource, ATarget: String): String;
begin
  while AnsiPos( ATarget, ASource ) > 0 do
    Delete( ASource, Pos( ATarget, ASource ), Length( ATarget ) );
  Result := ASource;
end;

{ ---------------------------------------------------------------------------- }

function GetSysDate(ADateSeparator: Boolean = False): String;
var
  AYear, AMonth, ADay: Word;
begin
  DecodeDate( Date, AYear, AMonth, ADay );
  Result := Format( '%.4d/%.2d/%.2d', [AYear, AMonth, ADay] );
  if not ADateSeparator then Result := TrimString( Result, '/' );
end;

{ ---------------------------------------------------------------------------- }

function GetSysTime(ATimeSeparator: Boolean = False): String;
begin
  DateTimeToString( Result, 'hh:mm:ss', Now );
  if not ATimeSeparator then Result := TrimString( Result, ':' );
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

function DateTextIsValid(const aDateText: String): Boolean;
var
  AYear, AMonth, ADay: Word;
  ACheckText: String;
begin
  Result := ( Trim( aDateText ) = '' );
  if not Result then
  begin
    ACheckText := TrimString( aDateText, '/' );
    if ( Length( ACheckText ) = 8 ) then
    begin
      AYear := StrToInt( Copy( ACheckText, 1, 4 ) );
      AMonth := StrToInt( Copy( ACheckText, 5, 2 ) );
      ADay := StrToInt( Copy( ACheckText, 7, 2 ) );
      Result := IsValidDate( AYear, AMonth, ADay );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function DateTextIsValidEx(var aDateText: String): Boolean;
var
  IsRocYear: Boolean;
  aTemp, aYear, aMonth, aDay, aFullDateStr: String;
  aConY, aConM, aConD: Integer;
  aConverDest: TDateTime;
begin
  aTemp := aDateText;
  aYear := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  aMonth := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  aDay := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  Result := (
     ( Trim( aYear ) = EmptyStr ) and
     ( Trim( aMonth ) = EmptyStr ) and
     ( Trim( aDay ) = EmptyStr ) );
  if Result then Exit;
  Result := (
    TryStrToInt( aYear, aConY ) and
    TryStrToInt( aMonth, aConM ) and
    TryStrToInt( aDay, aConD ) );
  if not Result then Exit;
  IsRocYear := ( aConY < 1000 );
  if IsRocYear then Inc( aConY, 1911 );
  aFullDateStr := Format( '%.4d/%.2d/%.2d', [aConY, aConM, aConD] );
  Result := TryStrToDate( aFullDateStr, aConverDest );
  if Result then
  begin
    if IsRocYear then
      aDateText := Format( '%.2d/%.2d/%.2d', [aConY-1911, aConM, aConD] )
    else
      aDateText := aFullDateStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

function DateTimeValidate(var aValue: String): Boolean;
var
  IsRocYear: Boolean;
  aTemp, aYear, aMonth, aDay, aFullDateStr: String;
  aConY, aConM, aConD: Integer;
  aConverDest: TDateTime;
begin
  aTemp := aValue;
  aYear := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  aMonth := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  aDay := TrimChar( ExtractValue( aTemp, '/' ),  ['_', #32] );
  Result := (
     ( Trim( aYear ) = EmptyStr ) and
     ( Trim( aMonth ) = EmptyStr ) and
     ( Trim( aDay ) = EmptyStr ) );
  if Result then Exit;
  Result := (
    TryStrToInt( aYear, aConY ) and
    TryStrToInt( aMonth, aConM ) and
    TryStrToInt( aDay, aConD ) );
  if not Result then Exit;
  IsRocYear := ( aConY < 1000 );
  if IsRocYear then Inc( aConY, 1911 );
  aFullDateStr := Format( '%.4d/%.2d/%.2d', [aConY, aConM, aConD] );
  Result := TryStrToDate( aFullDateStr, aConverDest );
  if Result then
  begin
    if IsRocYear then
      aValue := Format( '%.2d/%.2d/%.2d', [aConY-1911, aConM, aConD] )
    else
      aValue := aFullDateStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

function PadDateText(const ADateText: String): String;
var
  AText: String;
begin
  Result := '';
  if Trim( ADateText ) = '' then Exit;
  AText := FormatMaskText( '!9999/99/99;1;', ADateText );
  Result := StringReplace( AText, ' ', '0', [rfReplaceAll] );
end;


{ ---------------------------------------------------------------------------- }

function PadDateText2(const ADateText: String): String;
var
  AText: String;
begin
  Result := '';
  if Trim( ADateText ) = '' then Exit;
  AText := Lpad( Copy( ADateText, 1, 4 ), 4, '0' ) +
           Lpad( Copy( ADateText, 5, 2 ), 2, '0' ) +
           Lpad( Copy( ADateText, 7, 2 ), 2, '0' );
  Result := AText;
end;

{ ---------------------------------------------------------------------------- }

function DateConvert(aDate: TDateTime; aDateSeparator: Boolean = False): String;
var
  aYear, aMonth, aDay: Word;
begin
  DecodeDate( aDate, aYear, aMonth, aDay );
  Result := Format( '%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
  if not ADateSeparator then Result := TrimString( Result, '/' );
end;

{ ---------------------------------------------------------------------------- }

function DateConvert(aDate: String; aDateSeparator: Boolean = False): String;
var
  aDateTime: TDateTime;
begin
  Result := aDate;
  if TryStrToDate( aDate, aDateTime ) then
    Result := DateConvert( aDateTime, aDateSeparator );
end;

{ ---------------------------------------------------------------------------- }

function FormatMaskText(const EditMask: String; const Value: String): String;
begin
  Result := Value;
  if Trim( Value ) = EmptyStr then Exit;
  Result := MaskUtils.FormatMaskText( EditMask, Value );
end;

{ ---------------------------------------------------------------------------- }

function FileNameWithoutExt(const FileName: String): String;
var
  AIndex: Integer;
begin
  Result := '';
  AIndex := LastDelimiter( '.', FileName );
  if AIndex > 0 then Result := Copy( FileName, 1, AIndex - 1 );
end;

{ ---------------------------------------------------------------------------- }

function CbRoundTo(AValue: Extended; ADigit: Word): Extended;
begin

  Result:= StrToFloat( Format( '%1.' + IntToStr( ADigit ) + 'f', [AValue] ) );
end;

{ ---------------------------------------------------------------------------- }

function BinToInt(const ABinary: String): Integer;
var
  aIndex: Integer;
begin
  Result := 0;
  for aIndex := 1 to Length( ABinary ) do
    Result := Result shl 1 or ( Byte( ABinary[aIndex] ) and 1 );
end;

{ ---------------------------------------------------------------------------- }

(*
function TaxtNoValidate(const ANo: String): Boolean;
const
  ACalc: array [1..8] of Integer = ( 1, 2, 1, 2, 1, 2, 4, 1 );
var
  ANum: array [1..8] of Integer;
  AIndex, AIndex2, ASum: Integer;
  AChkStr: String;
  IsSeven: Boolean;
begin
  Result := False;
  IsSeven := False;
  if Length( Trim( ANo ) ) <> 8 then Exit;
  for AIndex := Low( ANum ) to High( ANum ) do ANum[AIndex] := 0;
  for AIndex := 1 to Length( ANo ) do
  begin
    if not ( ANo[AIndex] in ['0'..'9'] ) then Exit;
    if ( AIndex = 7 ) then IsSeven := ( ANo[AIndex] = '7' );
    ANum[AIndex] := StrToInt( ANo[AIndex] ) * ACalc[AIndex];
  end;
  for AIndex := Low( ANum ) to High( ANum ) do
  begin
    AChkStr := IntToStr( ANum[AIndex] );
    ANum[AIndex] := 0;
    for AIndex2 := 1 to Length( AChkStr ) do
      ANum[AIndex] := ( ANum[AIndex] + StrToInt( AChkStr[AIndex2] ) );
  end;
  ASum := 0;
  for AIndex := Low( ANum ) to High( ANum ) do
  begin
    if ( AIndex = 7 ) and IsSeven then Continue;
    ASum := ( ASum + ANum[AIndex] );
  end;
  Result := ( ( ASum mod 10 ) = 0 );
end;
*)

{ ---------------------------------------------------------------------------- }

function ExtractValue(var AValue: String; ASeparator: String = ','): String;
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

function ExtractValue(AIndex: Integer; AValue: String; ASeparator: String = ','): String; overload;
begin
  repeat
    Result := ExtractValue( AValue, ASeparator );
    Dec( AIndex );
  until ( AIndex <= 0 );
end;

{ ---------------------------------------------------------------------------- }

function QuotedValue(AValue: String; AQuoted: String = ''''; ASeparator: String = ','): String;
begin
  Result := EmptyStr;
  if ( AValue <> EmptyStr ) then
  begin
    repeat
       Result := ( Result + AQuoted +
         ExtractValue( AValue, ASeparator ) + AQuoted + ASeparator );
    until ( aValue = EmptyStr );
  end;
  if IsDelimiter( ASeparator, Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function ValuePos(ACompare, AValue: String; ASeparator: String = ','): Integer;
var
  aExt: String;
  aIndex: Integer;
begin
  Result := 0;
  aIndex := 0;
  repeat
    aExt := ExtractValue( AValue, ASeparator );
    Inc( aIndex );
    if AnsiSameText( aExt, ACompare ) then
      Result := aIndex;
  until ( Result > 0 ) or ( AValue = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function ValueExists(ACompare, AValue: String; ASeparator: String = ','): Boolean;
begin
  Result := ( ValuePos( ACompare, AValue, ASeparator ) > 0 );
end;

{ ---------------------------------------------------------------------------- }

function DeleteValue(ADelete, AValue: String; ASeparator: String = ','): String;
var
  aExt: String;
begin
  Result := EmptyStr;
  repeat
    aExt := ExtractValue( AValue, ASeparator );
    if not AnsiSameText( aExt, ADelete ) then
      Result := ( Result + aExt + ASeparator );
  until ( AValue = EmptyStr );
  if IsDelimiter( ASeparator, Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }
                                           
function AddValue(AAdd, AValue: String; ASeparator: String = ','): String;
begin
  Result := ( IIF( AValue = EmptyStr, EmptyStr, AValue + ASeparator ) + AAdd );
end;

{ ---------------------------------------------------------------------------- }

function TrimChar(const AValue: String; const AParams: array of const): String;

    { ----------------------------------------- }

    function InChar(AStr: String): Boolean;
    var
      aIndex: Integer;
    begin
      Result := False;
      for aIndex := Low( AParams ) to High( AParams ) do
      begin
        Result := ( CompareText( AStr, TVarRec( AParams[aIndex] ).VChar ) = 0 );
        if Result then Break;
      end;
    end;

    { ------------------------------------------}

var
  aIndex: Integer;
begin
  Result := AValue;
  if ( Result = EmptyStr ) then Exit;
  if Length( AParams ) <= 0 then Exit;
  for aIndex := Length( Result ) downto 1 do
  begin
    if not ( ByteType( Result, aIndex ) = mbSingleByte ) then Continue;
    if InChar( Result[aIndex] ) then Delete( Result, aIndex, 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function FormatDateTimeString(aSource: String): String;
begin
  Result := FormatString( aSource, ftDateTime );
end;

{ ---------------------------------------------------------------------------- }

function UTCDateTimeToLocalDateTime(AUTC: TDateTime): TDateTime;
begin
  Result := ( AUTC + ( 8 / 24 ) );
end;

{ ---------------------------------------------------------------------------- }

function UTCDateTimeToLocalDateTime(AUTC: String): TDateTime;
begin
  Result := UTCDateTimeToLocalDateTime( StrToDateTime( AUTC ) );
end;

{ ---------------------------------------------------------------------------- }

function LocalDateTimeToUTCDateTime(ALocal: TDateTime): TDateTime; overload;
begin
  Result := ( ALocal - ( 8 / 24 ) );
end;

{ ---------------------------------------------------------------------------- }

function LocalDateTimeToUTCDateTime(ALocal: String): TDateTime; overload;
begin
  Result := LocalDateTimeToUTCDateTime( StrToDateTime( ALocal ) );
end;

{ ---------------------------------------------------------------------------- }

function IIF(ATest: Boolean; const ATrue, AFalse: String): String;
begin
  Result := AFalse;
  if ATest then
    Result := ATrue;
end;

{ ---------------------------------------------------------------------------- }

function IIF(ATest: Boolean; const ATrue, AFalse: Integer): Integer;
begin
  Result := AFalse;
  if ATest then
    Result := ATrue;
end;

{ ---------------------------------------------------------------------------- }

function IIF(ATest: Boolean; const ATrue, AFalse: Double): Double; overload;
begin
  Result := AFalse;
  if ATest then
    Result := ATrue;
end;

{ ---------------------------------------------------------------------------- }

function GetComputerName: String;
var
  aBuffer: array [0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  aSize: Cardinal;
begin
  aSize := SizeOf( aBuffer );
  ZeroMemory( @aBuffer, aSize );
  Windows.GetComputerName( aBuffer, aSize );
  Result := aBuffer;
end;

{ ---------------------------------------------------------------------------- }

function GetIpAddress: String;
var
  aStack: TIdStackWindows;
begin
  aStack := TIdStackWindows.Create;
  try
    Result := aStack.LocalAddress;
  finally
    FreeAndNil( aStack );
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsRunInTerminalServer: Boolean;
begin
  Result := ( GetSystemMetrics( SM_REMOTESESSION  ) > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure GetTerminalServerInfo(var ATermId, ATermIp, ATermName, ATermPC, ATermState: String);
var
  aBuffer: Pointer;
  aBytes: Cardinal;

  { ----------------------------------------------------- }

  function WTSStateToString(const AState: TWtsConnectStateClass): String;
  begin
    case AState of
      WTSActive: Result := 'Active';
      WTSConnected: Result := 'Connected';
      WTSConnectQuery: Result := 'Connecting';
      WTSShadow: Result := 'Shadow';
      WTSDisconnected: Result := 'Disconnected';
      WTSIdle: Result := 'Idle';
      WTSListen: Result := 'Listen';
      WTSReset: Result := 'Reset';
      WTSDown: Result := 'Down';
      WTSInit: Result := 'Init';
    end;
  end;

  { ----------------------------------------------------- }

begin
  { Session Id }
  WTSQuerySessionInformation( WTS_CURRENT_SERVER, WTS_CURRENT_SESSION, WTSSessionId,
    aBuffer, aBytes );
  ATermId := IntToStr( Integer( aBuffer^ ) );
  WTSFreeMemory( aBuffer );
  { Terminal Session Client IP Address }
  WTSQuerySessionInformation( WTS_CURRENT_SERVER, WTS_CURRENT_SESSION, WTSClientAddress,
    aBuffer, aBytes );
  ATermIp := Format( '%d.%d.%d.%d', [
    PWTS_CLIENT_ADDRESS( aBuffer ).Address[2],
    PWTS_CLIENT_ADDRESS( aBuffer ).Address[3],
    PWTS_CLIENT_ADDRESS( aBuffer ).Address[4],
    PWTS_CLIENT_ADDRESS( aBuffer ).Address[5]] );
  WTSFreeMemory( aBuffer );
  { Terminal Session Name }
  WTSQuerySessionInformation( WTS_CURRENT_SERVER, WTS_CURRENT_SESSION, WTSWinStationName,
    aBuffer, aBytes );
  ATermName := PChar( aBuffer );
  WTSFreeMemory( aBuffer );
  { Terminal Session Client Host Name }
  WTSQuerySessionInformation( WTS_CURRENT_SERVER, WTS_CURRENT_SESSION, WTSClientName,
    aBuffer, aBytes );
  ATermPC := PChar( aBuffer );
  WTSFreeMemory( aBuffer );
  { Terminal Session State }
  WTSQuerySessionInformation( WTS_CURRENT_SERVER, WTS_CURRENT_SESSION, WTSConnectState,
    aBuffer, aBytes );
  ATermState := WTSStateToString( TWtsConnectStateClass( aBuffer^ ) );
  WTSFreeMemory( aBuffer );
end;

{ ---------------------------------------------------------------------------- }

function PreviousInstanceExists(ATitle: String; const AShow: Boolean;
  var ASemaphore: THandle): Boolean;
var
  AName, AClass: array [0..MAX_PATH] of Char;
  AWnd: THandle;
begin
  Result := False;
  ZeroMemory( @AName, SizeOf( ATitle ) );
  ZeroMemory( @AClass, SizeOf( 'TApplication' ) );
  StrPCopy( AName, ATitle );
  StrPCopy( AClass, 'TApplication' );
  ASemaphore := CreateSemaphore( nil, 0, 1, AName );
  if GetLastError = ERROR_ALREADY_EXISTS then
  begin
    if ASemaphore <> 0 then CloseHandle( ASemaphore );
    { 找出第一個 Application Instance }
    AWnd := FindWindow( AClass, AName );
    { 改 Application Title, 隨便填 }
    SetWindowText( AWnd, 'z,z,z' );
    { 找第二次 }
    AWnd := FindWindow( AClass, AName );
    Result := ( AWnd <> 0 );
    { 若仍找到, 帶出原本的 Application }
    if Result and AShow then
    begin
      if IsIconic( AWnd ) then ShowWindow( AWnd, SW_SHOWNORMAL );
      SetForegroundWindow( AWnd );
      SetActiveWindow( AWnd );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsProcessExists(const AProcessName: String): Boolean;
var
  aWtsInfo: PWTS_PROCESS_INFO;
  aIndex: Integer;
  aSessionId: ^Cardinal;
  aCount: Cardinal;
begin
  Result := False;
  WTSQuerySessionInformation( WTS_CURRENT_SERVER_HANDLE, WTS_CURRENT_SESSION,
    WTSSessionId, Pointer( aSessionId ), aCount );
  WTSEnumerateProcesses( WTS_CURRENT_SERVER_HANDLE, 0, 1, aWtsInfo, aCount );
  try
    for aIndex := 0 to aCount - 1 do
    begin
      {$IFDEF APPDEBUG} CodeSite.SendFmtMsg( 'PID=%d, PName=%s',
        [aWtsInfo.ProcessId, aWtsInfo.pProcessName] );
      {$ENDIF}
      Result :=
        ( aWtsInfo.SessionId = aSessionId^ ) and
        ( AnsiSameText( aWtsInfo.pProcessName, AProcessName ) and
        ( GetCurrentProcessId <> aWtsInfo.ProcessId ) );
      if Result then Break;
      Inc( aWtsInfo );
    end;
  finally
    WTSFreeMemory( Pointer( aSessionId ) );
    WTSFreeMemory( aWtsInfo );
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetProcessId(const AProcessName: String): Cardinal;
var
  aWtsInfo: PWTS_PROCESS_INFO;
  aIndex: Integer;
  aSessionId: ^Cardinal;
  aCount: Cardinal;
  aFound: Boolean;
begin
  Result := 0;
  WTSQuerySessionInformation( WTS_CURRENT_SERVER_HANDLE, WTS_CURRENT_SESSION,
    WTSSessionId, Pointer( aSessionId ), aCount );
  WTSEnumerateProcesses( WTS_CURRENT_SERVER_HANDLE, 0, 1, aWtsInfo, aCount );
  try
    for aIndex := 0 to aCount - 1 do
    begin
      aFound :=
        ( aWtsInfo.SessionId = aSessionId^ ) and
        ( AnsiSameText( aWtsInfo.pProcessName, AProcessName ) and
        ( GetCurrentProcessId <> aWtsInfo.ProcessId ) );
      if aFound then
      begin
        Result := aWtsInfo.ProcessId;
        Break;
      end;
      Inc( aWtsInfo );
    end;
  finally
    WTSFreeMemory( Pointer( aSessionId ) );
    WTSFreeMemory( aWtsInfo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure DecreaseProcessMemorySize(AProcessId: Integer = 0);
var
  aHnd: THandle;
begin
  if ( AProcessId = 0 ) then AProcessId := GetCurrentProcessId;
  aHnd := OpenProcess( PROCESS_ALL_ACCESS, False, AProcessId );
  try
    SetProcessWorkingSetSize( aHnd, DWORD( -1 ), DWORD( -1 ) );
  finally
    CloseHandle( aHnd );
  end;
end;

{ ---------------------------------------------------------------------------- }

function FileNotExistsAndCreate(AFileName: String; ACreateIfNotExists: Boolean): Boolean;
var
  aFileHnd: Integer;
begin
  Result := FileExists( AFileName );
  if ( not Result ) and ( ACreateIfNotExists ) then
  begin
    aFileHnd := FileCreate( AFileName );
    FileClose( aFileHnd );
    Result := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.

