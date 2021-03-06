unit CallbackService_Impl;

{----------------------------------------------------------------------------}
{ This unit was automatically generated by the RemObjects SDK after reading  }
{ the RODL file associated with this project .                               }
{                                                                            }
{ This is where you are supposed to code the implementation of your objects. }
{----------------------------------------------------------------------------}

{$I Remobjects.inc}

interface

uses
  Classes, SysUtils, Variants, DB,
{ App: }
  cbSrvClass, cbDesignPattern, cbLanguage, cbHrHelper,
{ ODAC: }
   Ora, MemDS, VirtualTable,
{ RemObjects: }
  uROXMLIntf, uROClientIntf, uROTypes, uROServer, uROServerIntf, uROSessions,
  uRORemoteDataModule, uROClient, CsHorse2Library_Intf, uROBinMessage,
  uROEventRepository;

type
  { TCallbackService }
  TCallbackService = class(TRORemoteDataModule, ICallbackService)
    procedure RORemoteDataModuleCreate(Sender: TObject);
    procedure RORemoteDataModuleDestroy(Sender: TObject);
  private
    FReader: TOraQuery;
    FWriter: TOraSQL;
    FBuffer: TVirtualTable;
    procedure MsgCallback(const AMsgInfo: TMsgInfo; const ARecver: TVirtualTable; var AErrMsg: String);
    function IsMsgRecver(const AInfo: TLoginInfo; const AMsgRecvCompCode, AMsgRecvUserId: String): Boolean;
    function UpdateFromSession(const AInfo: TLoginInfo): TLoginInfo;
    function GroupListToBuffer(const AInfo: TLoginInfo): Binary;
    function UserListToBuffer(const AInfo: TLoginInfo): Binary;
    function SaveMsg(const AInfo: TLoginInfo; const ARecver, AMsg: Binary;
      const AMsgInfo: TMsgInfo; var AErrMsg: String): Boolean;
    function MsgListToBuffer(const AInfo: TLoginInfo): Binary;
    function MsgToBuffer(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo): Binary;
    procedure UpdateReadMsg(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo);
  protected
    { ICallbackService methods }
    function GetGroupList(const AInfo: TLoginInfo): Binary;
    function GetUserList(const AInfo: TLoginInfo): Binary;
    function SendMsg(const AInfo: TLoginInfo; const ARecver: Binary; const AMsg: Binary; const AMsgInfo: TMsgInfo; var AErrMsg: String): Boolean;
    function GetMsgList(const AInfo: TLoginInfo): Binary;
    function GetMsg(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo): Binary;
    function GetOraSysDate(const ACompCode: String): String;
    procedure SetMsgRead(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo);
  end;

implementation

{$R *.dfm}

uses cbUtilis, cbROServerModule;

const MSG_LIST = 'GetMsgList.sql';


{ CallbackService }

{ ---------------------------------------------------------------------------- }

procedure TCallbackService.RORemoteDataModuleCreate(Sender: TObject);
begin
  FReader := TOraQuery.Create( nil );
  FReader.FetchAll := True;
  FWriter := TOraSQL.Create( nil );
  FBuffer := TVirtualTable.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackService.RORemoteDataModuleDestroy(Sender: TObject);
begin
  FReader.Session := nil;
  FWriter.Session := nil;
  FReader.Free;
  FWriter.Free;
  FBuffer.Free;
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GetGroupList(const AInfo: TLoginInfo): Binary;
var
  aInfo2: TLoginInfo;
begin
  Result := nil;
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  if ( aInfo2.CompStr <> EmptyStr ) then Result := GroupListToBuffer( aInfo2 );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GetUserList(const AInfo: TLoginInfo): Binary;
var
  aInfo2: TLoginInfo;
begin
  Result := nil;
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  if ( aInfo2.CompStr <> EmptyStr ) then Result := UserListToBuffer( aInfo2 );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.SendMsg(const AInfo: TLoginInfo; const ARecver: Binary;
  const AMsg: Binary; const AMsgInfo: TMsgInfo; var AErrMsg: String): Boolean;
var
  aInfo2: TLoginInfo;
begin
  Result := False;
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  if ( aInfo2.CompStr <> EmptyStr ) then
    Result := SaveMsg( aInfo2, ARecver, AMsg, AMsgInfo, AErrMsg );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.UpdateFromSession(const AInfo: TLoginInfo): TLoginInfo;
begin
  Result := AInfo;
  if ( not Assigned( Result ) ) and ( not VarIsNull( Session.Values['AInfo'] ) ) then
    Result := TLoginInfo( Integer( Session['AInfo'] ) );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GroupListToBuffer(const AInfo: TLoginInfo): Binary;
var
  aTemp, aCurrComp, aCompName, aSql: String;
  aDbSession: TDbSession;
begin
  Result := nil;
  TBufferHelper.CreateFieldDefs( biGroupList, FBuffer );
  TBufferHelper.CreateSqlText( biGroupList, aSql );
  {}
  aTemp := AInfo.CompStr;
  while ( aTemp <> EmptyStr ) do
  begin
    aCurrComp := ExtractValue( aTemp );
    aCompName := ROServerModule.GetCompName( aCurrComp );
    aDbSession := ROServerModule.GetDbSession( aCurrComp );
    if Assigned( aDbSession ) then
    begin
      FReader.Session := aDbSession.Connection;
      try
        FReader.SQL.Text := Format( aSql, [aCurrComp, aCompName] );
        FReader.Open;
        FReader.First;
        while not FReader.Eof do
        begin
         FBuffer.Append;
         FBuffer.FieldByName( 'CompCode' ).AsString := FReader.FieldByName( 'CompCode' ).AsString;
         FBuffer.FieldByName( 'CompName' ).AsString := FReader.FieldByName( 'CompName' ).AsString;
         FBuffer.FieldByName( 'GroupId' ).AsString := FReader.FieldByName( 'GroupId' ).AsString;
         FBuffer.FieldByName( 'GroupName' ).AsString := FReader.FieldByName( 'GroupName' ).AsString;
         FBuffer.FieldByName( 'RGroupId' ).AsString := FReader.FieldByName( 'RGroupId' ).AsString;
         FBuffer.Post;
         FReader.Next;
        end;
        Sleep( 10 );
      finally
        FReader.Close;
        FReader.Session := nil;
        ROServerModule.ReleaseDbSession( aCurrComp, aDbSession );
      end;
    end;
  end;
  if ( not FBuffer.IsEmpty ) then
  begin
    Result := TROBinaryMemoryStream.Create;
    FBuffer.SaveToStream( Result, True );
    FBuffer.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.IsMsgRecver(const AInfo: TLoginInfo;
  const AMsgRecvCompCode, AMsgRecvUserId: String): Boolean;
begin
  Result := False;
  if Assigned( AInfo ) then
  begin
    Result := ( AInfo.CompCode = AMsgRecvCompCode ) and ( AInfo.UserId = AMsgRecvUserId );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.UserListToBuffer(const AInfo: TLoginInfo): Binary;
var
  aTemp, aCurrComp, aCompName, aSql: String;
  aDbSession: TDbSession;
begin
  Result := nil;
  TBufferHelper.CreateFieldDefs( biUserList, FBuffer );
  TBufferHelper.CreateSqlText( biUserList, aSql );
  {}
  aTemp := AInfo.CompStr;
  while ( aTemp <> EmptyStr ) do
  begin
    aCurrComp := ExtractValue( aTemp );
    aCompName := ROServerModule.GetCompName( aCurrComp );
    aDbSession := ROServerModule.GetDbSession( aCurrComp );
    if Assigned( aDbSession ) then
    begin
      FReader.Session := aDbSession.Connection;
      try
        FReader.Close;
        FReader.SQL.Text := Format( aSql, [aCurrComp, aCompName] );
        FReader.Open;
        FReader.First;
        while not FReader.Eof do
        begin
         FBuffer.Append;
         FBuffer.FieldByName( 'CompCode' ).AsString := FReader.FieldByName( 'CompCode' ).AsString;
         FBuffer.FieldByName( 'CompName' ).AsString := FReader.FieldByName( 'CompName' ).AsString;
         FBuffer.FieldByName( 'UserId' ).AsString := FReader.FieldByName( 'UserId' ).AsString;
         FBuffer.FieldByName( 'UserName' ).AsString := FReader.FieldByName( 'UserName' ).AsString;
         FBuffer.FieldByName( 'WorkClass' ).AsString := FReader.FieldByName( 'WorkClass' ).AsString;
         FBuffer.FieldByName( 'GroupName' ).AsString := FReader.FieldByName( 'GroupName' ).AsString;
         FBuffer.FieldByName( 'RGroupId' ).AsString := FReader.FieldByName( 'WorkClass' ).AsString;
         FBuffer.FieldByName( 'Status' ).AsString := FReader.FieldByName( 'Status' ).AsString;
         FBuffer.FieldByName( 'RStatus' ).AsString := FReader.FieldByName( 'Status' ).AsString;
         FBuffer.FieldByName( 'SessionId' ).AsString := FReader.FieldByName( 'SessionId' ).AsString;
         FBuffer.FieldByName( 'DisplayType' ).AsInteger := 2;
         FBuffer.Post;
         FReader.Next;
        end;
        Sleep( 10 );
      finally
        FReader.Close;
        FReader.Session := nil;
        ROServerModule.ReleaseDbSession( aCurrComp, aDbSession );
      end;
    end;
  end;
  if ( not FBuffer.IsEmpty ) then
  begin
    Result := TROBinaryMemoryStream.Create;
    FBuffer.SaveToStream( Result, True );
    FBuffer.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackService.MsgCallback(const AMsgInfo: TMsgInfo;
  const ARecver: TVirtualTable; var AErrMsg: String);
var
  aEvent: ISrvCallbackEvent_Writer;
  aIndex: Integer;
  aInfo2: TLoginInfo;
begin
  AErrMsg := EmptyStr;
  aEvent := Self.EventRepository as ISrvCallbackEvent_Writer;
  aEvent.ExcludeSender := True;
  ARecver.First;
  while not ARecver.Eof do
  begin
    aIndex := 0;
    UserList.BeginWriteLock;
    try
      while aIndex < UserList.Count do
      begin
        aInfo2 := TLoginInfo( UserList.Objects[aIndex] );
        if IsMsgRecver( aInfo2,
          ARecver.FieldByName( 'CompCode' ).AsString,
          ARecver.FieldByName( 'UserId' ).AsString ) then
        begin
          aEvent.SessionList.Add( aInfo2.SessionId );
          Break;
        end;
        Inc( aIndex );
      end;
    finally
      UserList.EndWriteLock;
    end;
    ARecver.Next;
  end;
  try
    aEvent.MsgChange( ClientID, AMsgInfo );
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  if ( aErrMsg <> EmptyStr ) then
    ROServerModule.LoginMsgSubject.Error( LanguageManager.GetFmt( 'SServiceCallbackError2', [aErrMsg] ) );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.SaveMsg(const AInfo: TLoginInfo; const ARecver, AMsg: Binary;
  const AMsgInfo: TMsgInfo; var AErrMsg: String): Boolean;

  { ------------------------------------------- }

  procedure PrepareCH006SqlParam;
  begin
    FWriter.ParamByName( 'CompCode' ).DataType := ftInteger;
    FWriter.ParamByName( 'CompCode' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgId' ).DataType := ftString;
    FWriter.ParamByName( 'MsgId' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgType' ).DataType := ftInteger;
    FWriter.ParamByName( 'MsgType' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgPriority' ).DataType := ftInteger;
    FWriter.ParamByName( 'MsgPriority' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgContent' ).DataType := ftOraClob;
    FWriter.ParamByName( 'MsgContent' ).ParamType := ptInput;
    FWriter.ParamByName( 'IsMsgReply' ).DataType := ftInteger;
    FWriter.ParamByName( 'IsMsgReply' ).ParamType := ptInput;
    FWriter.ParamByName( 'UserId' ).DataType := ftString;
    FWriter.ParamByName( 'UserId' ).ParamType := ptInput;
    FWriter.ParamByName( 'UserName' ).DataType := ftString;
    FWriter.ParamByName( 'UserName' ).ParamType := ptInput;
    FWriter.ParamByName( 'HostName' ).DataType := ftString;
    FWriter.ParamByName( 'HostName' ).ParamType := ptInput;
    FWriter.ParamByName( 'Ip' ).DataType := ftString;
    FWriter.ParamByName( 'Ip' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgTime' ).DataType := ftString;
    FWriter.ParamByName( 'MsgTime' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgSubject' ).DataType := ftString;
    FWriter.ParamByName( 'MsgSubject' ).ParamType := ptInput;
  end;

  { ------------------------------------------- }

  procedure PrepareCH006ASqlParam;
  begin
    FWriter.ParamByName( 'CompCode' ).DataType := ftInteger;
    FWriter.ParamByName( 'CompCode' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgId' ).DataType := ftString;
    FWriter.ParamByName( 'MsgId' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgRecvCompCode' ).DataType := ftInteger;
    FWriter.ParamByName( 'MsgRecvCompCode' ).ParamType := ptInput;
    FWriter.ParamByName( 'MsgRecver' ).DataType := ftString;
    FWriter.ParamByName( 'MsgRecver' ).ParamType := ptInput;
  end;

  { ------------------------------------------- }

  procedure ExecuteCH006Sql;
  begin
    FWriter.ParamByName( 'CompCode' ).AsString := AInfo.CompCode;
    FWriter.ParamByName( 'MsgId' ).AsString := AMsgInfo.MsgId;
    FWriter.ParamByName( 'MsgType' ).AsInteger := 0;
    FWriter.ParamByName( 'MsgPriority' ).AsString := AMsgInfo.MsgPriority;
    FWriter.ParamByName( 'MsgContent' ).LoadFromStream( AMsg, ftOraClob );
    FWriter.ParamByName( 'IsMsgReply' ).AsInteger := 0;
    FWriter.ParamByName( 'UserId' ).AsString := AInfo.UserId;
    FWriter.ParamByName( 'UserName' ).AsString := AInfo.UserName;
    if ( AInfo.TermSIP <> EmptyStr ) then
    begin
      FWriter.ParamByName( 'HostName' ).AsString := AInfo.TermSPC;
      FWriter.ParamByName( 'Ip' ).AsString := AInfo.TermSIP;
    end else
    begin
      FWriter.ParamByName( 'HostName' ).AsString := AInfo.HostName;
      FWriter.ParamByName( 'Ip' ).AsString := AInfo.IP;
    end;
    FWriter.ParamByName( 'MsgTime' ).AsString := AMsgInfo.MsgTime;
    FWriter.ParamByName( 'MsgSubject' ).AsString := AMsgInfo.MsgSubject;
    FWriter.Execute;
  end;

  { ------------------------------------------- }

  procedure ExecuteCH006ASql;
  begin
    FBuffer.First;
    while not FBuffer.Eof do
    begin
      FWriter.ParamByName( 'CompCode' ).AsString := AInfo.CompCode;
      FWriter.ParamByName( 'MsgId' ).AsString := AMsgInfo.MsgId;
      FWriter.ParamByName( 'MsgRecvCompCode' ).AsString := FBuffer.FieldByName( 'CompCode' ).AsString;
      FWriter.ParamByName( 'MsgRecver' ).AsString := FBuffer.FieldByName( 'UserId' ).AsString;
      FWriter.Execute;
      FBuffer.Next;
    end;
  end;

  { ------------------------------------------- }

var
  aSql: String;
  aDbSession: TDbSession;
begin
  AErrMsg := EmptyStr;
  aDbSession := ROServerModule.GetDbSession( AInfo.CompCode );
  if Assigned( aDbSession ) then
  begin
    FWriter.Session := aDbSession.Connection;
    try
      FBuffer.Close;
      FBuffer.Open;
      FBuffer.LoadFromStream( ARecver );
      FBuffer.First;
      TBufferHelper.CreateSqlText( biInsertCH006, aSql );
      FWriter.Text := aSql;
      PrepareCH006SqlParam;
      try
        AMsgInfo.MsgId := ROServerModule.MsgIdGenerator.NextValue;
        AMsgInfo.MsgTime := GetOraSysDate( AInfo.CompCode );
        if ( AMsgInfo.MsgTime = EmptyStr ) then
          AMsgInfo.MsgTime := GetSysDate( True ) + ' ' + GetSysTime( True );
        ExecuteCH006Sql;
        TBufferHelper.CreateSqlText( biInsertCH006A, aSql );
        FWriter.Text := aSql;
        PrepareCH006ASqlParam;
        ExecuteCH006ASql;
        FWriter.Session.Commit;
      except
        on E: Exception do
        begin
          AErrMsg := E.Message;
        end;
      end;
    finally
      FWriter.Session := nil;
      ROServerModule.ReleaseDbSession( AInfo.CompCode, aDbSession );
    end;
    if ( AErrMsg = EmptyStr ) then
    begin
      try
        MsgCallback( AMsgInfo, FBuffer, AErrMsg );
      except
        on E: Exception do AErrMsg := E.Message;
      end;
    end;
  end;
  Result := ( AErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GetMsgList(const AInfo: TLoginInfo): Binary;
var
  aInfo2: TLoginInfo;
begin
  Result := nil;
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  Result := MsgListToBuffer( aInfo2 );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.MsgListToBuffer(const AInfo: TLoginInfo): Binary;

  { ------------------------------------------- }

  procedure PrepareSqlParam;
  begin
    FWriter.ParamByName( 'AGetType' ).DataType := ftInteger;
    FWriter.ParamByName( 'AGetType' ).ParamType := ptInput;

    FWriter.ParamByName( 'ACount' ).DataType := ftInteger;
    FWriter.ParamByName( 'ACount' ).ParamType := ptOutput;

    FWriter.ParamByName( 'ACompCode' ).DataType := ftString;
    FWriter.ParamByName( 'ACompCode' ).ParamType := ptInput;

    FWriter.ParamByName( 'ARecvCompCode' ).DataType := ftString;
    FWriter.ParamByName( 'ARecvCompCode' ).ParamType := ptInput;

    FWriter.ParamByName( 'ARecvUserId' ).DataType := ftString;
    FWriter.ParamByName( 'ARecvUserId' ).ParamType := ptInput;

    FWriter.ParamByName( 'AMsgPriority' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgPriority' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgPriority' ).Table := True;
    FWriter.ParamByName( 'AMsgPriority' ).Length := 1;

    FWriter.ParamByName( 'AMsgId' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgId' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgId' ).Table := True;
    FWriter.ParamByName( 'AMsgId' ).Length := 1;

    FWriter.ParamByName( 'AMsgSubject' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgSubject' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgSubject' ).Table := True;
    FWriter.ParamByName( 'AMsgSubject' ).Length := 1;

    FWriter.ParamByName( 'AMsgTime' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgTime' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgTime' ).Table := True;
    FWriter.ParamByName( 'AMsgTime' ).Length := 1;

    FWriter.ParamByName( 'AMsgReply' ).DataType := ftInteger;
    FWriter.ParamByName( 'AMsgReply' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgReply' ).Table := True;
    FWriter.ParamByName( 'AMsgReply' ).Length := 1;

    FWriter.ParamByName( 'AMsgSenderId' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgSenderId' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgSenderId' ).Table := True;
    FWriter.ParamByName( 'AMsgSenderId' ).Length := 1;

    FWriter.ParamByName( 'AMsgSenderName' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgSenderName' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgSenderName' ).Table := True;
    FWriter.ParamByName( 'AMsgSenderName' ).Length := 1;

    FWriter.ParamByName( 'AMsgSenderWorkClass' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgSenderWorkClass' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgSenderWorkClass' ).Table := True;
    FWriter.ParamByName( 'AMsgSenderWorkClass' ).Length := 1;

    FWriter.ParamByName( 'AMsgSenderWorkName' ).DataType := ftString;
    FWriter.ParamByName( 'AMsgSenderWorkName' ).ParamType := ptOutput;
    FWriter.ParamByName( 'AMsgSenderWorkName' ).Table := True;
    FWriter.ParamByName( 'AMsgSenderWorkName' ).Length := 1;
  end;

  { ------------------------------------------- }

var
  aCompIndex, aIndex, aRecords: Integer;
  aPath: String;
  aDbSession: TDbSession;
begin
  Result := nil;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    Format( '%s\%s', [SQLFOLDER, MSG_LIST] );
  if not FileExists( aPath ) then Exit;
  TBufferHelper.CreateFieldDefs( biMsgList, FBuffer );
  FWriter.SQL.Clear;
  FWriter.SQL.LoadFromFile( aPath );
  PrepareSqlParam;
  for aCompIndex := 0 to ROServerModule.SoList.Count - 1 do
  begin
    if not ROServerModule.SoList[aCompIndex].Selected then Continue;
    aDbSession := ROServerModule.GetDbSession( ROServerModule.SoList[aCompIndex].CompCode );
    if Assigned( aDbSession ) then
    begin
      try
        try
          FWriter.Session := aDbSession.Connection;
          FWriter.ParamByName( 'AGetType' ).AsInteger := 1;
          FWriter.ParamByName( 'ACompCode' ).AsString := ROServerModule.SoList[aCompIndex].CompCode;
          FWriter.ParamByName( 'ARecvCompCode' ).AsString := AInfo.CompCode;
          FWriter.ParamByName( 'ARecvUserId' ).AsString := AInfo.UserId;
          FWriter.Execute;
          aRecords := FWriter.ParamByName( 'ACount' ).AsInteger;
          for aIndex := 0 to FWriter.ParamCount - 1 do
          begin
            if ( FWriter.Params[aIndex].Table ) then
              FWriter.Params[aIndex].Length := aRecords;
          end;
          FWriter.ParamByName( 'AGetType' ).AsInteger := 2;
          FWriter.Execute;
          for aIndex := 1 to aRecords do
          begin
            FBuffer.Append;
            FBuffer.FieldByName( 'CompCode' ).AsString := ROServerModule.SoList[aCompIndex].CompCode;
            FBuffer.FieldByName( 'CompName' ).AsString := ROServerModule.GetCompName( ROServerModule.SoList[aCompIndex].CompCode );
            FBuffer.FieldByName( 'MsgPriority' ).AsString := FWriter.ParamByName( 'AMsgPriority' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgId' ).AsString := FWriter.ParamByName( 'AMsgId' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgSubject' ).AsString := FWriter.ParamByName( 'AMsgSubject' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgTime' ).AsString := FWriter.ParamByName( 'AMsgTime' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgReply' ).AsString := FWriter.ParamByName( 'AMsgReply' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgSenderId' ).AsString := FWriter.ParamByName( 'AMsgSenderId' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgSenderName' ).AsString := FWriter.ParamByName( 'AMsgSenderName' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgSenderWorkClass' ).AsString := FWriter.ParamByName( 'AMsgSenderWorkClass' ).ItemAsString[aIndex];
            FBuffer.FieldByName( 'MsgSenderWorkName' ).AsString := FWriter.ParamByName( 'AMsgSenderWorkName' ).ItemAsString[aIndex];
            FBuffer.Post;
          end;
        except
          on E: Exception do
          begin
          end;
        end;
      finally
        FWriter.Session := nil;
        ROServerModule.ReleaseDbSession( ROServerModule.SoList[aCompIndex].CompCode, aDbSession );
      end;
    end;
  end;
  if ( not FBuffer.IsEmpty ) then
  begin
    Result := TROBinaryMemoryStream.Create;
    FBuffer.SaveToStream( Result, True );
    FBuffer.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GetMsg(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo): Binary;
var
  aInfo2: TLoginInfo;
begin
  Result := nil;
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  Result := MsgToBuffer( aInfo2, AMsgInfo );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.MsgToBuffer(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo): Binary;
var
  aSql: String;
  aDbSession: TDbSession;
begin
  Result := nil;
  aDbSession := ROServerModule.GetDbSession( AMsgInfo.CompCode );
  if Assigned( aDbSession ) then
  begin
    FReader.Session := aDbSession.Connection;
    FReader.Close;
    TBufferHelper.CreateSqlText( biMsg, aSql );
    try
      FReader.SQL.Text := Format( aSql, [AMsgInfo.CompCode, AMsgInfo.MsgSenderId, AMsgInfo.MsgId] );
      FReader.Open;
      if ( not FReader.FieldByName( 'MsgContent' ).IsNull ) then
      begin
        Result := Binary.Create;
        TBlobField( FReader.FieldByName( 'MsgContent' ) ).SaveToStream( Result );
        FReader.Close;
      end;
    finally
      FReader.Session := nil;
      ROServerModule.ReleaseDbSession( AMsgInfo.CompCode, aDbSession );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCallbackService.GetOraSysDate(const ACompCode: String): String;
var
  aDbSession: TDbSession;
begin
  Result := EmptyStr;
  aDbSession := ROServerModule.GetDbSession( ACompCode );
  if Assigned( aDbSession ) then
  begin
    FReader.Session := aDbSession.Connection;
    try
      FReader.SQL.Text := ' select to_char( sysdate, ''yyyy/mm/dd hh24:mi:ss'' ) from dual ';
      FReader.Open;
      Result := FReader.Fields[0].AsString;
    finally
      FReader.Close;
      FReader.Session := nil;
      ROServerModule.ReleaseDbSession( ACompCode, aDbSession );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackService.UpdateReadMsg(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo);
var
  aDbSession: TDbSession;
  aSql: String; 
begin
  aDbSession := ROServerModule.GetDbSession( AMsgInfo.CompCode );
  if Assigned( aDbSession ) then
  begin
    FWriter.Session := aDbSession.Connection;
    TBufferHelper.CreateSqlText( biSetReadMsg, aSql );
    try
      FWriter.Text := Format( aSql, [AMsgInfo.CompCode, AMsgInfo.MsgId, AInfo.CompCode, AInfo.UserId] );
      FWriter.Execute;
      FWriter.Session.Commit;
    finally
      FReader.Session := nil;
      ROServerModule.ReleaseDbSession( AMsgInfo.CompCode, aDbSession );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackService.SetMsgRead(const AInfo: TLoginInfo; const AMsgInfo: TMsgInfo);
var
  aInfo2: TLoginInfo;
begin
  aInfo2 := UpdateFromSession( AInfo );
  if not Assigned( aInfo2 ) then Exit;
  UpdateReadMsg( aInfo2, AMsgInfo );
end;

{ ---------------------------------------------------------------------------- }

end.
