unit CommU;



interface
uses
  SysUtils,Dialogs,Windows,Classes,IniFiles, Forms;
const
  INV_SYS_INFO_INI_FILE1 = 'CallInv_0.ini';
  INV_SYS_INFO_INI_FILE2 = 'CallInv_1.ini';
  TMP_INV_SYS_INFO_INI_FILE = 'Tmp.ini';
  INV_Def_Field_INI_FILE = 'DefInvFld.ini';
  function AppCreate:Boolean;
  procedure TransTmpIniFile(const aSourceFileName, aTempFileName: String);
  function BuildDataArear(const aFileName: String) : Boolean;


implementation
uses Encryption_TLB,dtmCommU;


function AppCreate:Boolean;
var
  aSourceFileName, aTempFileName: String;
  aTempDir: array [0..MAX_PATH-1] of Char;
begin
  //#5922 耞琌ㄏノ穦INI郎 By Kin 2011/02/17
  if ParamStr(5) = '0' then
  begin
    aSourceFileName := IncludeTrailingPathDelimiter(
      ExtractFilePath( Application.ExeName ) ) + INV_SYS_INFO_INI_FILE1;
  end else
  begin
    aSourceFileName := IncludeTrailingPathDelimiter(
      ExtractFilePath( Application.ExeName ) ) + INV_SYS_INFO_INI_FILE2;
  end;
  ZeroMemory( @aTempDir, SizeOf( aTempDir ) );
  GetTempPath( SizeOf( aTempDir ), aTempDir );
  aTempFileName := IncludeTrailingPathDelimiter( String( aTempDir ) ) +
    TMP_INV_SYS_INFO_INI_FILE;
  dtmComm.FInvDefFile := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_Def_Field_INI_FILE;
  try
    if FileExists( aSourceFileName ) then
    begin
      TransTmpIniFile( aSourceFileName, aTempFileName );
      try
        if BuildDataArear( aTempFileName ) then
          Result := True
        else
          Result := False;
      finally
        DeleteFile(PChar(aTempFileName));
      end;
    end else
    begin
      MessageDlg('тぃ硈挡祇布╰参把计郎',mtWarning,[mbOK],0);
      Result := False;
    end;
  except
    Result := False;
  end;
end;

procedure TransTmpIniFile(const aSourceFileName, aTempFileName: String);
var
  aStrList, aTmpStrList : TStringList;
  aIndex: Integer;
  aIntf: _Password;
  aEncKey: WideString;
begin
  aEncKey := 'CS';
  aStrList := TStringList.Create;
  try
    aTmpStrList := TStringList.Create;
    try
      aIntf := CoPassword.Create;
      try
        aStrList.LoadFromFile( aSourceFileName );
        for AIndex := 0 to aStrList.Count - 1 do
        begin
          if ( Copy( aStrList.Strings[AIndex], 1, 2 ) <> '//' ) then
            aTmpStrList.Add( aIntf.Decrypt( aStrList.Strings[aIndex],
              aEncKey ) )
          else
            aTmpStrList.Add( aStrList.Strings[AIndex] );
        end;
        aTmpStrList.SaveToFile( aTempFileName );
      finally
        aIntf := nil;
      end;
    finally
      aTmpStrList.Free;
    end;
  finally
     aStrList.Free;
  end;
end;
function BuildDataArear(const aFileName: String) : Boolean;
begin
  if  dtmComm.GetSMSDb(aFileName) then
    Result := True
  else
    Result := False;
end;



end.
