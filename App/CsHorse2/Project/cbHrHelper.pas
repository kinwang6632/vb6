unit cbHrHelper;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DB,
{ App: }
  cbUtilis,  
{ ODAC: }
  MemDS, VirtualTable;

type

  { TCreateBufferId }

  TCreateBufferId = ( biLoginInfo, biClientParam, biCH005, biSO021, biCD042,
    biSO022, biSO023, biCH009, biCD071, biGroupList, biUserList, biUserNode,
    biRecvList, biInsertCH006, biInsertCH006A, biMsgList, biMsg, biSetReadMsg );

  { TBufferHelper }

  TBufferHelper = class(TObject)
  public
    class procedure CreateFieldDefs(const AId: TCreateBufferId; ABuffer: TVirtualTable);
    class procedure CreateSqlText(const AId: TCreateBufferId; var ASql: String);
  end;

  function IsCompCodeFound(ACompCode, ACompStr: String): Boolean;

implementation

{ TBufferHelper }

class procedure TBufferHelper.CreateFieldDefs(const AId: TCreateBufferId;
  ABuffer: TVirtualTable);
begin
  ABuffer.Close;
  ABuffer.DeleteFields;
  case AId of
    biClientParam:
      begin
        { 手動或自動更新資料 }
        ABuffer.AddField( 'AutoRefresh', ftBoolean );
        { 可切換公司別更新頻率 }
        ABuffer.AddField( 'AuthorizeRefreshRate', ftInteger );
        { 公佈欄自動更新頻率, 單位秒 }
        ABuffer.AddField( 'AnnRefreshRate', ftInteger );
        { 線上使用者清單 (包含群組) 更新頻率, 單位秒 }
        ABuffer.AddField( 'UserRefreshRate' , ftInteger );
        { 當斷線後每多少秒自動重試連線 }
        ABuffer.AddField( 'TryReconnectRate', ftInteger );
      end;
    biCH005:
      begin
        ABuffer.AddField( 'SessionId', ftString, 50 );
        ABuffer.AddField( 'CompCode', ftString, 3 );
        ABuffer.AddField( 'CompName' , ftString, 100 );
        ABuffer.AddField( 'UserId', ftString, 10 );
        ABuffer.AddField( 'UserName', ftString, 20 );
        ABuffer.AddField( 'CompStr' , ftString, 200 );
        ABuffer.AddField( 'HostName', ftString, 20 );
        ABuffer.AddField( 'IP', ftString, 15 );
        ABuffer.AddField( 'WorkClassCode', ftString, 3 );
        ABuffer.AddField( 'WorkClassName', ftString, 20 );
        ABuffer.AddField( 'Status', ftString, 1 );
        ABuffer.AddField( 'TermSId', ftString, 5 );
        ABuffer.AddField( 'TermSName', ftString, 32 );
        ABuffer.AddField( 'TermSPC', ftString, 63 );
        ABuffer.AddField( 'TermSIP', ftString, 15 );
        ABuffer.AddField( 'TermState', ftString, 20 );
        ABuffer.AddField( 'OnTime', ftDateTime );
        ABuffer.AddField( 'OffTime',ftDateTime );
      end;
    biCD042:
      begin
        ABuffer.AddField( 'CompCode', ftString, 3 );
        ABuffer.AddField( 'CompName', ftString, 100 );
        ABuffer.AddField( 'CodeNo', ftString, 10 );
        ABuffer.AddField( 'ActDateStToEd', ftString, 50 );
        ABuffer.AddField( 'Description', ftString, 200 );
        ABuffer.AddField( 'Note', ftMemo );

      end;
    biSO021:
      begin
        ABuffer.AddField( 'CompCode', ftString, 3 );
        ABuffer.AddField( 'CompName', ftString, 100 );
        ABuffer.AddField( 'BoardTime', ftString, 50 );
        ABuffer.AddField( 'BoardEn', ftString, 20 );
        ABuffer.AddField( 'Subject', ftString, 100 );
        ABuffer.AddField( 'Content', ftMemo );
      end;
    biSO022:
      begin
        ABuffer.AddField( 'CompCode', ftString, 3 );
        ABuffer.AddField( 'CompName', ftString, 100 );
        ABuffer.AddField( 'SNo', ftString, 8 );
        ABuffer.AddField( 'ErrorTime', ftString, 50 );
        ABuffer.AddField( 'EndTime', ftString, 50 );
        ABuffer.AddField( 'MfCode', ftString, 3 );
        ABuffer.AddField( 'MfName', ftString, 50 );
        ABuffer.AddField( 'Description', ftString, 80 );
        ABuffer.AddField( 'AcceptName', ftString, 50 );
      end;
    biSO023:
      begin
        ABuffer.AddField( 'CompCode', ftString, 3 );
        ABuffer.AddField( 'CompName', ftString, 100 );
        ABuffer.AddField( 'SNo', ftString, 8 );
        ABuffer.AddField( 'Address', ftMemo );
      end;
    biCH009:
      begin
        ABuffer.AddField( 'CompCode',ftInteger );
        ABuffer.AddField( 'CompName',ftString, 100 );
        ABuffer.AddField( 'GroupId', ftString, 3 );
        ABuffer.AddField( 'GroupName', ftString, 20 );
        ABuffer.AddField( 'RGroupId', ftString, 3 );
        ABuffer.AddField( 'StopFlag', ftInteger );
        ABuffer.AddField( 'ImageIndex', ftInteger );
      end;
    biCD071:
      begin
        ABuffer.AddField( 'CompCode',ftInteger );
        ABuffer.AddField( 'CompName',ftString, 100 );
        ABuffer.AddField( 'CodeNo', ftString, 3 );
        ABuffer.AddField( 'Description', ftString, 20 );
        ABuffer.AddField( 'StopFlag', ftInteger );
        ABuffer.AddField( 'ImageIndex', ftInteger );
      end;
    biGroupList:
      begin
        ABuffer.AddField( 'CompCode',ftInteger );
        ABuffer.AddField( 'CompName',ftString, 100 );
        ABuffer.AddField( 'GroupId', ftString, 3 );
        ABuffer.AddField( 'GroupName', ftString, 20 );
        ABuffer.AddField( 'RGroupId', ftString, 3 );
      end;
    biUserList:
      begin
        ABuffer.AddField( 'CompCode',ftInteger );
        ABuffer.AddField( 'CompName',ftString, 100 );
        ABuffer.AddField( 'UserId',ftString, 10 );
        ABuffer.AddField( 'UserName',ftString, 20 );
        ABuffer.AddField( 'CompStr',ftString, 200);
        ABuffer.AddField( 'WorkClass',ftString, 3 );
        ABuffer.AddField( 'GroupName',ftString, 20 );
        ABuffer.AddField( 'RGroupId',ftString, 3 );
        ABuffer.AddField( 'Status',ftString, 1 );
        ABuffer.AddField( 'RStatus',ftString, 1 );
        ABuffer.AddField( 'SessionId', ftString, 50 );
        ABuffer.AddField( 'ImageIndex', ftInteger );
        ABuffer.AddField( 'DisplayType', ftInteger );
        ABuffer.AddField( 'DisplayText1', ftString, 100 );
        ABuffer.AddField( 'DisplayText2', ftString, 100 );
      end;
    biUserNode:
      begin
        ABuffer.AddField( 'CompCode',ftInteger );
        ABuffer.AddField( 'UserId',ftString, 10 );
      end;
    biRecvList:
      begin
        ABuffer.AddField( 'CompCode', ftInteger );
        ABuffer.AddField( 'UserId', ftString, 10 );
        ABuffer.AddField( 'Status',ftString, 1 );
        ABuffer.AddField( 'DisplayText', ftString, 100 );
      end;
    biMsgList:
      begin
        ABuffer.AddField( 'CompCode', ftInteger );
        ABuffer.AddField( 'CompName', ftString, 10 );
        ABuffer.AddField( 'MsgPriority', ftInteger );
        ABuffer.AddField( 'MsgId', ftString, 12 );
        ABuffer.AddField( 'MsgSubject', ftString, 500 );
        ABuffer.AddField( 'MsgTime', ftString, 20 );
        ABuffer.AddField( 'MsgReply', ftInteger, 0 );
        ABuffer.AddField( 'MsgSenderId', ftString, 10 );
        ABuffer.AddField( 'MsgSenderName', ftString, 20 );
        ABuffer.AddField( 'MsgSenderWorkClass', ftString, 3 );
        ABuffer.AddField( 'MsgSenderWorkName', ftString, 20 );
      end;  
  end;
  ABuffer.Open;
end;

{ ---------------------------------------------------------------------------- }

class procedure TBufferHelper.CreateSqlText(const AId: TCreateBufferId; var ASql: String);
begin
  ASql := EmptyStr;
  case AId of
    biLoginInfo:
      begin
        ASql :=
          ' select a.compstr,                   ' +
          '        a.username,                  ' +
          '        b.description,               ' +
          '        a.workclass,                ' +
          '        c.groupname                  ' +
          '   from ch005 a, cd039 b, ch009 c    ' +
          '  where a.compcode = b.codeno(+)     ' +
          '    and a.compcode = ''%s''          ' +
          '    and a.compcode = c.compcode(+)   ' +
          '    and a.workclass = c.groupid(+)   ' +
          '    and a.userid = ''%s''            ' +
          '    and a.stopflag = 0               ';
      end;
    biCH005:
      begin
        ASql :=
         '  SELECT                              ' +
         '      a.compcode,                     ' +
         '      b.description as compname,      ' +
         '      a.sessionid,                    ' +
         '      a.userid,                       ' +
         '      a.username,                     ' +
         '      a.compstr,                      ' +
         '      a.hostname,                     ' +
         '      a.ip,                           ' +
         '      a.workclass,                    ' +
         '      c.description as workclassname, ' +
         '      a.stopflag,                     ' +
         '      a.status,                       ' +
         '      a.termsid,                      ' +
         '      a.termsname,                    ' +
         '      a.termspc,                      ' +
         '      a.termsip,                      ' +
         '      a.termstate,                    ' +
         '      a.ontime,                       ' +
         '      a.offtime                       ' +
         '  FROM CH005 A, CD039 B, CD071 C      ' +
         ' WHERE a.compcode = b.codeno          ' +
         '   AND a.workclass = c.codeno         ' +
         ' ORDER BY a.compcode, a.userid        ';
      end;
    biCD042:
      begin
        ASql :=
          ' select ''%s'' as compcode,          ' +
          '        ''%s'' as compname,          ' +
          '        a.codeno,                    ' +
          '        a.actstartdate,              ' +
          '        a.actstopdate,               ' +
          '        a.description,               ' +
          '        a.note                       ' +
          '  from cd042 a                       ' +
          ' where stopflag = 0                  ' +
          '   and sysdate between actstartdate and actstopdate ' +
          ' order by actstartdate desc  ';
      end;
    biSO021:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.boardtime,            ' +
          '        a.boarden,              ' +
          '        a.subject,              ' +
          '        a.content               ' +
          '  from so021 a                  ' +
          ' where stopflag = 0             ' +
          '   and sysdate between boarddate1 and boarddate2 ' +
          ' order by boardtime desc        ';
      end;
    biSO022:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.sno,                  ' +
          '        a.errortime,            ' +
          '        a.endtime,              ' +
          '        a.mfcode,               ' +
          '        a.mfname,               ' +
          '        a.description,          ' +
          '        a.acceptname            ' +
          '  from so022 a                  ' +
          ' where fintime is null          ' +
          ' order by errortime desc        ';
      end;
    biSO023:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.sno,                  ' +
          '        a.address               ' +
          '  from so023 a, so022 b         ' +
          ' where a.sno = b.sno            ' +
          '   and b.fintime is null        ' +
          ' order by a.sno desc, a.ord     ';
      end;
    biCH009:
      begin
        ASql :=
          ' select ''%s'' as compcode,      ' +
          '        ''%s'' as compname,      ' +
          '        a.groupid,               ' +
          '        a.groupname,             ' +
          '        a.rgroupid,              ' +
          '        nvl( b.stopflag, 0 ) as stopflag    ' +
          '   from ch009 a, cd071 b         ' +
          '  where a.groupid = b.codeno(+)  ' +
          '  order by a.groupid             ';
      end;
    biCD071:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.codeno,               ' +
          '        a.description,          ' +
          '        a.stopflag              ' +
          '   from cd071 a                 ' +
          ' order by a.codeno              ';
      end;
    biGroupList:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.groupid,              ' +
          '        a.groupname,            ' +
          '        a.rgroupid              ' +
          '   from ch009 a, cd071 b        ' +
          '  where a.groupid = b.codeno    ' +
          '    and b.stopflag = 0          ' +
          '  order by a.groupid,           ' +
          '           a.rgroupid           ';
      end;
    biUserList:
      begin
        ASql :=
          ' select ''%s'' as compcode,     ' +
          '        ''%s'' as compname,     ' +
          '        a.userid,               ' +
          '        a.username,             ' +
          '        a.compstr,              ' +
          '        a.workclass,            ' +
          '        b.groupname,            ' +
          '        b.rgroupid,             ' +
          '        a.status,               ' +
          '        a.sessionid             ' +
          '   from ch005 a, ch009 b        ' +
          '  where a.workclass = b.groupid ' +
          '    and a.stopflag = 0          ';
      end;
    biInsertCH006 :
      begin
        ASql :=
          ' INSERT INTO CH006 (            ' +
          '    CompCode,                   ' +
          '    MsgId,                      ' +
          '    MsgType,                    ' +
          '    MsgPriority,                ' +
          '    MsgContent,                 ' +
          '    IsMsgReply,                 ' +
          '    UserId,                     ' +
          '    UserName,                   ' +
          '    HostName,                   ' +
          '    Ip,                         ' +
          '    MsgTime,                    ' +
          '    BoardDate1,                 ' +
          '    BoardDate2,                 ' +
          '    MsgSubject )                ' +
          ' VALUES (                       ' +
          '    :CompCode,                  ' +
          '    :MsgId,                     ' +
          '    :MsgType,                   ' +
          '    :MsgPriority,               ' +
          '    EMPTY_CLOB(),               ' +
          '    :IsMsgReply,                ' +
          '    :UserId,                    ' +
          '    :UserName,                  ' +
          '    :HostName,                  ' +
          '    :Ip,                        ' +
          '    TO_DATE( :MsgTime, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
          '    NULL,                       ' +
          '    NULL,                       ' +
          '    :MsgSubject )               ' +
          '  RETURNING MsgContent          ' +
          '  INTO :MsgContent              ';
      end;
    biInsertCH006A:
      begin
        ASql :=
          ' INSERT INTO CH006A (           ' +
          '    CompCode,                   ' +
          '    MsgId,                      ' +
          '    MsgRecvCompCode,            ' +
          '    MsgRecver,                  ' +
          '    MsgRecvTime,                ' +
          '    MsgReply          )         ' +
          ' VALUES (                       ' +
          '    :CompCode,                  ' +
          '    :MsgId,                     ' +
          '    :MsgRecvCompCode,           ' +
          '    :MsgRecver,                 ' +
          '    NULL,                       ' +
          '    NULL          )             ';
      end;
    biMsg:
      begin
        ASql :=
          ' SELECT A.MsgContent            ' +
          '   FROM CH006 A                 ' +
          '  WHERE CompCode = ''%s''       ' +
          '    AND UserId = ''%s''         ' +
          '    AND MsgId = ''%s''          ';
      end;
    biSetReadMsg:
      begin
        ASql :=
          ' UPDATE CH006A A                   ' +
          '    SET A.MsgRecvTime = SYSDATE    ' +
          '  WHERE A.CompCode = ''%s''        ' +
          '    AND A.MsgId = ''%s''           ' +
          '    AND A.MsgRecvCompCode = ''%s'' ' +
          '    AND A.MsgRecver = ''%s''       ';
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsCompCodeFound(ACompCode, ACompStr: String): Boolean;
begin
  repeat
    Result := ( ACompCode = ExtractValue( ACompStr ) );
  until ( Result or ( ACompStr = EmptyStr ) );
end;

{ ---------------------------------------------------------------------------- }

end.
