unit frmCD078BAU;

interface

uses
   Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, ADODB, Menus, ComObj,

  cbDBController, cbStyleController,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cxCurrencyEdit, cxCheckBox, cxRadioGroup,
  cxLookAndFeelPainters, cxButtons, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdFTP;

type TvarAry = array of String;
type
  TfrmCD078BA = class(TForm)
    pnl1: TPanel;
    btnCompCodeStr: TButton;
    cxButton2: TcxButton;
    edtCompCodeStr: TcxTextEdit;
    pnl2: TPanel;
    bvl1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    conDataConnection: TADOConnection;
    adoDataReader: TADOQuery;
    adoDataWriter: TADOQuery;
    IdFTP1: TIdFTP;
    procedure FormCreate(Sender: TObject);
    procedure btnCompCodeStrClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    FCloneCodeNo: String;
    FCloneDescription: String;
    BpCodeStr: String;
    BpNameStr: string;
    CompCodeStr: String;
    CompCodeStrCanSelect: String;
    CompNameStr: String;
    FReader: TADOQuery;
    FWriter: TADOQuery;
    FSourceLinkKeyList: TStringList;
    FCloneLinkKeyList: TStringList;
    FExists:Boolean;
    function Split(const aSourceStr, aSplitStr : String) : TvarAry;
    function UploadFileToFtp(var errMessage : string) :Boolean;
    procedure WriteStreamStr(Stream : TStream; Str : string);
    procedure WriteStreamInt(Stream : TStream; Num : integer);
  public
    { Public declarations }
  end;

var
  frmCD078BA: TfrmCD078BA;
implementation
uses cbUtilis, frmMultiSelectU;
{$R *.dfm}

{ TfrmCD078BA }

procedure TfrmCD078BA.FormCreate(Sender: TObject);
var
  aIndex: Integer;
begin
  FSourceLinkKeyList := TStringList.Create;
  FCloneLinkKeyList := TStringList.Create;
  FReader := DBController.DataReader;
  //#6644 ???P?_?v??,???H?h?[1=1 By Kin 2013/10/25
  if ( DBController.LoginInfo.UserId = 'A' ) or ( 1=1 ) then
    begin
      CompCodeStrCanSelect := '';
      for aIndex :=1 to 15 do
        CompCodeStrCanSelect := CompCodeStrCanSelect + IntToStr(aIndex)+',';
      CompCodeStrCanSelect := Copy (CompCodeStrCanSelect, 1, Length(CompCodeStrCanSelect)-1)
    end
  else
    begin
      FReader.Close;
      FReader.SQL.Text := Format (
        'SELECT COMPSTR FROM %s.SO026 WHERE USERID=''%s'' ',
        [DBController.LoginInfo.DbAccount, DBController.LoginInfo.UserId] );
      FReader.Open;
      CompCodeStrCanSelect := FReader.FieldByName('COMPSTR').AsString;
    end;
  edtCompCodeStr.Text := '(??????)';
end;

procedure TfrmCD078BA.btnCompCodeStrClick(Sender: TObject);
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := DBController.DataConnection;
    frmMultiSelect.KeyFields := 'CODENO';
    frmMultiSelect.KeyValues := CompCodeStr;
    frmMultiSelect.DisplayFields := 'CODENO, ?N?X,DESCRIPTION,?W??';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION FROM %s.CD052 WHERE 1=1 ',
       [DBController.LoginInfo.DbAccount,
       DBController.LoginInfo.CompCode]);
    {???n?u???w1-15???q?O By Kin }
    if ( CompCodeStrCanSelect <> EmptyStr )  then
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        ' AND CODENO IN ( ' + CompCodeStrCanSelect +' ) OR 1 = 1'

    else
      frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
        ' And 1=0 ';
    frmMultiSelect.SQL.Text := frmMultiSelect.SQL.Text +
      '  ORDER BY CODENO ';
    if frmMultiSelect.ShowModal = mrOk then
      begin
        CompCodeStr := frmMultiSelect.SelectedValue;
        CompNameStr := frmMultiSelect.SelectedDisplay;
        edtCompCodeStr.Text := frmMultiSelect.SelectedDisplay;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TfrmCD078BA.btnSaveClick(Sender: TObject);
var
  aIndex: Integer;
  aCompCodeStr, aCompCode,aCompNameStr,aCompName: String;
  aData, aOwner, aPass: String;
  aErrMessage: String;
  aList: TStringList;
  EnCryptX: OleVariant;
begin
  if  (CompCodeStr = EmptyStr) then
  begin
    WarningMsg( '?|???????i???q?O?j?C' );
  end else
  begin
    {}
    EnCryptX := CreateOleObject( 'DevPower.Encrypt' );
    aList := TStringList.Create;
    Screen.Cursor := crHourGlass;
    try
      aList.LoadFromFile( ExtractFilePath(Application.ExeName) + 'GICMIS2.INI' );
      aCompCodeStr := CompCodeStr;
      aCompNameStr := CompNameStr;
      //???q?O?j??
      while ( aCompCodeStr <> EmptyStr ) do
      begin
        aCompCode := ExtractValue( aCompCodeStr );
        aCompCode := StringReplace( aCompCode, '''', EmptyStr, [rfReplaceAll] );
        aCompName := ExtractValue( aCompNameStr );

        //?d ConnectString ?? DB USER
        if DBController.LoginInfo.CompCode = aCompCode then
        begin
          aData := DBController.LoginInfo.DbAliase;
          aOwner := DBController.logininfo.DbAccount;
          aPass := DBController.Logininfo.DbPassword;
        end else
        begin
          for aIndex := 0 to aList.Count -1 do
          begin
            if aList.Strings[aIndex] = '[' + aCompCode + ']' then
            begin
              aData := Copy( aList.Strings[aIndex + 1],3,Length(aList.Strings[aIndex + 1] ) - 2 );
              aData := EnCryptX.DeCrypt( aData );
              aOwner := Copy(aList.Strings[aIndex + 2],3,Length(aList.Strings[aIndex + 2] ) - 2 );
              aOwner := EnCryptX.DeCrypt(aOwner);
              aPass := Copy( aList.Strings[aIndex + 3],3,Length(aList.Strings[aIndex + 3] ) - 2 );
              aPass := EnCryptX.DeCrypt( aPass );
            end;
          end;
        end;
        conDataConnection.ConnectionString := Format(
          'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
          [aPass, aOwner, aData] );
        try
          try
            conDataConnection.Open;
            aErrMessage := EmptyStr;
            UploadFileToFtp( aErrMessage );
            if aErrMessage <> EmptyStr then
            begin
              WarningMsg( Format( '???q?O?G[ %s ] ?????~' + #13+ #10+ '???~???]?G%s'
                                   ,[aCompName, aErrMessage ])) ;
            end;
            {
            if not UploadFileToFtp( aErrMessage) then
            begin
              WarningMsg( aErrMessage );
            end;
            }
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '???q?O?G[ %s ] ???????L?k?}?? ' + #13+#10 +'???]?G%s?C', [aCompName,E.Message] ) );
            end;
          end;
        finally
          conDataConnection.Close;
        end;
      end;
    finally
      aList.Free;
      Screen.Cursor := crDefault;
      MessageDlg( '?????????I',mtInformation,[mbOK],0);
    end;
    {
    if ( aTotalErrCount <= 0 ) then
      InfoMsg( aMessage )
    else
      WarningMsg( aMessage );
    }
  end;
end;

function TfrmCD078BA.Split(const aSourceStr, aSplitStr: String): TvarAry;
var
  aStr: string;
  aSource: String;
  aIndex : Integer;
  aArry : TvarAry;
begin
  if Length(aSourceStr) = 0 then
  begin
    Result := nil;
  end else
  begin
    aStr := EmptyStr;
    aSource := aSourceStr;
    if Pos(aSplitStr,aSource) > 0 then
    begin
      for aIndex := 1 to Length( aSource ) do
      begin
        if ( aSource[aIndex] <> aSplitStr ) then
        begin
          aStr := ( aStr + aSource[aIndex] );
        end
        else begin
          SetLength( aArry, Length( aArry ) + 1 );
          aArry[ Length( aArry ) - 1 ] := aStr;
          aStr := '';
        end;
      end;
      if ( aStr <> '' ) then
      begin
        SetLength( aArry, Length( aArry ) + 1 );
        aArry[ Length( aArry ) - 1 ] := aStr;
        aStr := '';
      end;
    end else
    begin
      SetLength(aArry,1);
      aArry[ Length( aArry ) - 1 ] := aSource;
    end;
  end;
  Result := aArry;
end;

function TfrmCD078BA.UploadFileToFtp(var errMessage: string): Boolean;
  var
    aFtpFileName : String;
    aStringSplit,aUserInfo,aHostInfo : TvarAry;
    aIndex,aIndex2 :Integer;
    aList : TStringList;
    aStm : TStream;
    aHost,aPassword,aUserName,aFileName,aDirName : String;
begin
  try
    try
      aList := TStringList.Create;
      aStm := TMemoryStream.Create;
      adoDataReader.Close;
      adoDataReader.SQL.CommaText := 'SELECT * FROM v_CNSBPCodeToCSV ';
      adoDataReader.Open;
      if adoDataReader.RecordCount > 0 then
      begin
        aFtpFileName := adoDataReader.FieldByName( 'FTPFILENAME' ).AsString;
        aFtpFileName := StringReplace( aFtpFileName ,'ftp://',EmptyStr,[rfReplaceAll] );
        aStringSplit := Split(aFtpFileName, '@');
        for aIndex := Low(aStringSplit) to High(aStringSplit) do
        begin
          if aIndex = 0 then
          begin
            aUserInfo := Split(aStringSplit[0],':');
            aUserName := aUserInfo[0];
            for aIndex2 := 1 to High( aUserInfo ) do
            begin
              if aIndex2 = 1 then
              begin
                aPassword := aUserInfo[aIndex2];
              end else
              begin
                aPassword := aPassword + ':' + aUserInfo[aIndex2];
              end;
            end;
          end else
          begin
            aHostInfo := Split(aStringSplit[1],'/');
            aHost := aHostInfo[0];
            aDirName := '/';
            for aIndex2 := 1 to High(aHostInfo)-1 do
            begin
              if aIndex2 = 1 then
                aDirName := '/' + aHostInfo[aIndex2]
              else
                aDirName := aDirName + '/' + aHostInfo[aIndex2];
            end;
            aFileName := aHostInfo[High(aHostInfo)];
          end;
        end;
        IdFTP1.Host := aHost;
        IdFTP1.Username := aUserName;
        IdFTP1.Password := aPassword;
        IdFTP1.Connect;
        IdFTP1.ChangeDir( aDirName );
//        aList.Add( '?}??????');
//        aList.Add( '??????ABCD1122');
        adoDataReader.First;
        while not adoDataReader.Eof do
        begin
//          aList.Count
          aList.Add( adoDataReader.FieldByName( 'TXTLINE' ).AsString );
          adoDataReader.Next;
        end;
        aList.SaveToStream(aStm);
        IdFTP1.Put( aStm , ExtractFileName( aFileName ));
        adoDataReader.First;
        {
        while not adoDataReader.Eof do
        begin

        end;
        }
        Result := True;
      end else
      begin
        errMessage := '?L????FTP???T?i?}??!';
        Result := False;
      end;
    except
      on E: Exception do
      begin
        errMessage := e.Message;
        Result:= False;
  //      ErrorMsg( Format( '???q?O:[ %s ] ???????L?k?}??, ???]:%s?C', [aCompName,E.Message] ) );
      end;
    end;
  finally
    if IdFTP1.Connected then
      IdFTP1.Abort;
    IdFTP1.Quit;
    aList.Free;
    aStm.Free;
    adoDataReader.Close;
  end;
end;

procedure TfrmCD078BA.WriteStreamStr(Stream: TStream; Str: string);
var
 StrLen : integer;
begin
  {get length of string}
  StrLen := Length(Str);
  {write length of string}
  WriteStreamInt(Stream, StrLen);
  if StrLen > 0 then
  {write characters}
  Stream.Write(Str[1], StrLen);
end;

procedure TfrmCD078BA.WriteStreamInt(Stream: TStream; Num: integer);
begin
  Stream.WriteBuffer(Num, SizeOf(Integer));
end;

procedure TfrmCD078BA.btnCancelClick(Sender: TObject);
begin
  Self.Close;
end;

end.
