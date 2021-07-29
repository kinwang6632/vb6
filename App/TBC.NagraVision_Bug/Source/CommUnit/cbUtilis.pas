unit cbUtilis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms, IniFiles
  {$IFDEF APPDEBUG}, CodeSiteLogging {$ENDIF};



{ --------------------------------------------------------------------- }
{ ��ܸ߰� MessageBox                                                   }
{ --------------------------------------------------------------------- }
function ConfirmMsg(const AMsg: String): Boolean;
function ConfirmMsgEx(const AMsg: String): Boolean;


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
{ �p���ɮפj�p, �^�ǭȳ�� KB                                           }
{ --------------------------------------------------------------------- }
function CalcFileSizeInKB(const AFullFileName: String): Double;


{ --------------------------------------------------------------------- }
{ �p���ɮפj�p, �^�ǭȳ�� Byte                                         }
{ --------------------------------------------------------------------- }
function CalcFileSizeInByte(const AFullFileName: String): Cardinal;



{ --------------------------------------------------------------------- }
{ ������                                                              }
{   AValue: �@��, Ex: 3000 = 3 ��                                       }
{ --------------------------------------------------------------------- }
procedure Delay(const AValue: Cardinal);



{ --------------------------------------------------------------------- }
{ �h�����w���r��                                                        }
{ --------------------------------------------------------------------- }
function TrimString(ASource, ATarget: String): String;


{ --------------------------------------------------------------------- }
{ ���t�Τ��                                                            }
{     AType: dtROC  : ����~                                            }
{            dtEng  : �褸�~                                            }
{ --------------------------------------------------------------------- }
function GetSysDate(ADateSeparator: Boolean = False): String;


{ --------------------------------------------------------------------- }
{ ���t�ήɶ�                                                            }
{ --------------------------------------------------------------------- }
function GetSysTime(ATimeSeparator: Boolean = False): String;


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
{ DateTextIsValid                                                       }
{ DateTextIsValidEx                                                     }
{  ���ҿ�J�����(�褸�Υ���~)�O�_�X�k, Ex: 94/1/1 or 2005/1/1         }
{ --------------------------------------------------------------------- }
function DateTextIsValid(const aDateText: String): Boolean;
function DateTextIsValidEx(var aDateText: String): Boolean;


{ --------------------------------------------------------------------- }
{ PadDateText                                                           }
{ PadDateText2                                                          }
{   �ɨ�����r�� Ex: 2003/1/2 ---> 2003/01/02                           }
{                Ex: 200312   ---> 20031200                             }
{ --------------------------------------------------------------------- }
function PadDateText(const ADateText: String): String;
function PadDateText2(const ADateText: String): String;


{ --------------------------------------------------------------------- }
{ DateConvert                                                           }
{   �N�褸�~�ഫ������~                                                }
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
{    �зǪ��|�ˤ��J , Ex: 123.45 --> 123.5                              }
{ --------------------------------------------------------------------- }
function CbRoundTo(AValue: Extended; ADigit: Word): Extended;


{ --------------------------------------------------------------------- }
{ BinToInt                                                              }
{    �G�i����ƭ�, Ex: 100100 --> 36                                    }
{ --------------------------------------------------------------------- }
function BinToInt(const ABinary: String): Integer;


{ --------------------------------------------------------------------- }
{ ��Q�Ʒ~�Τ@�s���ˬd                                                  }
{                                                                       }
{ --------------------------------------------------------------------- }
//function TaxtNoValidate(const ANo: String): Boolean;



{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   �ѥX�æ^�ǨϥΤ��j�r�����j����, Ex: 123,456,789                     }
{                                                                       }
{ --------------------------------------------------------------------- }
function ExtractValue(var AValue: String; ASeparator: String = ','): String; overload;

{ --------------------------------------------------------------------- }
{ ExtractValue                                                          }
{   �ѥX�æ^�ǫ��w��m����, Ex: 123,456,789                             }
{   AIndex: ���w����m, �q 1 �}�l                                       }
{ --------------------------------------------------------------------- }
function ExtractValue(AIndex: Integer; AValue: String; ASeparator: String = ','): String; overload;



{ --------------------------------------------------------------------- }
{ QuotedValue                                                           }
{   �N�C�Ӧr���[�W�]���r��, Ex: 123,456,789 --> '123','456','789'       }
{                                                                       }
{ --------------------------------------------------------------------- }
function QuotedValue(AValue: String; AQuoted: String = ''''; ASeparator: String = ','): String;


{ --------------------------------------------------------------------- }
{ ValuePos                                                              }
{   ���o�ŦX�Ӧr�ꪺ��m, Ex: ACompare=456, AValue=123,456,789          }
{   �^�� 2                                                              }
{ --------------------------------------------------------------------- }
function ValuePos(ACompare, AValue: String; ASeparator: String = ','): Integer;


{ --------------------------------------------------------------------- }
{ ValueExists                                                           }
{   �O�_���ŦX�Ӧr�ꪺ, Ex: ACompare=456, AValue=123,456,789            }
{   �^�� True                                                           }
{ --------------------------------------------------------------------- }
function ValueExists(ACompare, AValue: String; ASeparator: String = ','): Boolean;


{ --------------------------------------------------------------------- }
{ DeleteValue                                                           }
{   �h���ŦX�Ӧr�ꪺ�r��, Ex: ACompare=456, AValue=123,456,789          }
{   �^�� 123,789                                                        }
{ --------------------------------------------------------------------- }
function DeleteValue(ADelete, AValue: String; ASeparator: String = ','): String;


{ --------------------------------------------------------------------- }
{ AddValue                                                              }
{   �W�[�r�ꪺ��, Ex: AAdd=456, AValue=123,789                          }
{   �^�� 123,789,456                                                    }
{ --------------------------------------------------------------------- }
function AddValue(AAdd, AValue: String; ASeparator: String = ','): String;



{ --------------------------------------------------------------------- }
{ TrimChar                                                              }
{   �R���X�{�b�ӷ��r�ꤤ���w���r��, Ex: 2005/1_/1_                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function TrimChar(const AValue: String; const AParams: array of const): String;



{ --------------------------------------------------------------------- }
{ FormatDateTimeString                                                  }
{   �榡�Ʀr�ꦨ������j�r�� Ex: 20051201180030 --> 2005/12/01 18:00:30 }
{                                                                       }
{ --------------------------------------------------------------------- }
function FormatDateTimeString(aSource: String): String;



{ --------------------------------------------------------------------- }
{ UTCDateTimeToLocalTime                                                }
{   �N��L�ªv����ɶ��ഫ���x�W�ɶ�                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function UTCDateTimeToLocalDateTime(AUTC: TDateTime): TDateTime; overload;
function UTCDateTimeToLocalDateTime(AUTC: String): TDateTime; overload;


{ --------------------------------------------------------------------- }
{ UTCDateTimeToLocalTime                                                }
{   �N�x�W�ɶ��ഫ����L�ªv����ɶ�                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function LocalDateTimeToUTCDateTime(ALocal: TDateTime): TDateTime; overload;
function LocalDateTimeToUTCDateTime(ALocal: String): TDateTime; overload;


{ --------------------------------------------------------------------- }
{ IIF                                                                   }
{   �Ǧ^�ŦX���󪺭�                                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function IIF(ATest: Boolean; const ATrue, AFalse: String): String; overload;
function IIF(ATest: Boolean; const ATrue, AFalse: Integer): Integer; overload;
function IIF(ATest: Boolean; const ATrue, AFalse: Double): Double; overload;



{ --------------------------------------------------------------------- }
{ GetComputerName                                                       }
{   ���������q���W��                                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetComputerName: String;


{ --------------------------------------------------------------------- }
{ GetComputerName                                                       }
{   ��������IP��}                                                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetIpAddress: String;


{ --------------------------------------------------------------------- }
{ IsRunInTerminalServer                                                 }
{   �O�_����b�L�n Terminal Server �W                                   }
{                                                                       }
{ --------------------------------------------------------------------- }
function IsRunInTerminalServer: Boolean;


{ --------------------------------------------------------------------- }
{ GetTerminalServerInfo                                                 }
{   �������b�L�n Terminal Server �W����T                               }
{                                                                       }
{ --------------------------------------------------------------------- }
procedure GetTerminalServerInfo(var ATermId, ATermIp, ATermName, ATermPC, ATermState: String);


{ --------------------------------------------------------------------- }
{ PreviousInstanceExists                                                }
{   ���ε{���_�w�g���椤                                                }
{                                                                       }
{ --------------------------------------------------------------------- }
function PreviousInstanceExists(ATitle: String; const AShow: Boolean;
  var ASemaphore: THandle): Boolean;


{ --------------------------------------------------------------------- }
{ IsProcessExists                                                       }
{   ���ε{���O�_�w�g���椤                                      }
{                                                                       }
{ --------------------------------------------------------------------- }
function IsProcessExists(const AProcessName: String): Boolean;


{ --------------------------------------------------------------------- }
{ GetProcessId                                                          }
{   ���o���w���ε{���� Process Handle                                   }
{                                                                       }
{ --------------------------------------------------------------------- }
function GetProcessId(const AProcessName: String): Cardinal;


{ --------------------------------------------------------------------- }
{ RetrieveApplicationMemorySize                                         }
{   �^�����椤�����ε{���O����ϥζq                                    }
{                                                                       }
{ --------------------------------------------------------------------- }
procedure DecreaseProcessMemorySize(AProcessId: Integer = 0);


{ --------------------------------------------------------------------- }
{ FileExists                                                            }
{   �˴����w�� IniFile �O�_�s�b,�Y���s�b�h�إ�                          }
{                                                                       }
{ --------------------------------------------------------------------- }
function FileNotExistsAndCreate(AFileName: String; ACreateIfNotExists: Boolean): Boolean;


{ --------------------------------------------------------------------- }
{ FileExists                                                            }
{   �˴����w�� IniFile �O�_�s�b,�Y���s�b�h�إ�                          }
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
    Result := ( Application.MessageBox( PChar( AMsg ), '�T�{', MB_OKCANCEL +
      MB_ICONQUESTION ) = ID_OK );
end;

{ ---------------------------------------------------------------------------- }

function ConfirmMsgEx(const AMsg: String): Boolean;
begin
  Result := False;
  if ( AMsg <> EmptyStr ) then
    Result := ( Application.MessageBox( PChar( AMsg ), '�T�{', MB_OKCANCEL +
      MB_ICONQUESTION + MB_DEFBUTTON2	) = ID_OK );
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
    { ��X�Ĥ@�� Application Instance }
    AWnd := FindWindow( AClass, AName );
    { �� Application Title, �H�K�� }
    SetWindowText( AWnd, 'z,z,z' );
    { ��ĤG�� }
    AWnd := FindWindow( AClass, AName );
    Result := ( AWnd <> 0 );
    { �Y�����, �a�X�쥻�� Application }
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

