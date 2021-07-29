unit cbConfigModule;

interface

uses
  SysUtils, Classes, Windows, Variants, IniFiles, Forms, DB, ADODB, Controls,
  DBClient, cxTL, cbAppClass, CbClass;

type
  TConfigModule = class(TDataModule)
    AccessConnection: TADOConnection;
    DataReader: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FConfigFileName: String;
    FSoList: TAppSoList;
    FMsgSubject: TMessageSubject;
    function ConnectToAccess(const aFileName: String): Boolean;
    procedure DisconnectFromAccess;
    procedure ReadSettingValue;
    function GetSoList: Boolean;
    function SetSoList: Boolean;
  public
    { Public declarations }
    function ShowConfigForm: TModalResult;
    function LoadFromFile(const aFileName: String): Boolean;
    function SaveToFile(const aFileName: String): Boolean;
    property SoList: TAppSoList read FSoList;
    property MsgSubject: TMessageSubject read FMsgSubject;
  end;

var
  ConfigModule: TConfigModule;

implementation

uses cbUtilis, cbMain, cbConfig;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FConfigFileName := EmptyStr;
  FSoList := TAppSoList.Create;
  FMsgSubject := TMessageSubject.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
begin
  FSoList.Free;
  FMsgSubject.Free;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.LoadFromFile(const aFileName: String): Boolean;
begin
  Result := ConnectToAccess( aFileName );
  if Result then
  begin
    FConfigFileName := aFileName;
    FSoList.Clear;
    Result := GetSoList;
  end;  
  DisconnectFromAccess;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SaveToFile(const aFileName: String): Boolean;
begin
  Result := ConnectToAccess( aFileName );
  if Result then
    Result := SetSoList;
  DisconnectFromAccess;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ConnectToAccess(const aFileName: String): Boolean;
const
  aDbPassword = 'cyc84177282';
begin
  AccessConnection.Connected := False;
  AccessConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aDbPassword] );
  try
    AccessConnection.Connected := True;
  except
    on E: Exception do
      FMsgSubject.Error ( Format( '設定檔讀取錯誤, 訊息:%s。', [E.Message] ) );
  end;
  Result := AccessConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DisconnectFromAccess;
begin
  if AccessConnection.Connected then
    AccessConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetSoList: Boolean;
var
  aSo: TAppSo;
begin
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.SO_SELECTED,               ' +
    '        A.SO_POS,                    ' +
    '        B.SO_POS_NAME,               ' +
    '        A.SO_COMP_CODE,              ' +
    '        A.SO_COMP_NAME,              ' +
    '        A.SO_LOGINUSER,              ' +
    '        A.SO_LOGINPASS,              ' +
    '        A.SO_DBALIASE,               ' +
    '        A.CA_LOGINUSER,              ' +
    '        A.CA_LOGINPASS,              ' +
    '        A.CA_DBALIASE                ' +
    '   FROM SO_COMP A, SOPOS_ENV B       ' +
    '  WHERE A.SO_POS = B.SO_POS          ' +
    '  ORDER BY CINT( A.SO_COMP_CODE )    ';
  try
    DataReader.Open;
    DataReader.First;
    while not DataReader.Eof do
    begin
      aSo := TAppSo.Create;
      aSo.Selected := ( DataReader.FieldByName( 'SO_SELECTED' ).AsString = 'Y' );
      aSo.Pos := DataReader.FieldByName( 'SO_POS' ).AsInteger;
      aSo.PosName := DataReader.FieldByName( 'SO_POS_NAME' ).AsString;
      aSo.CompCode := DataReader.FieldByName( 'SO_COMP_CODE' ).AsString;
      aSo.CompName := DataReader.FieldByName( 'SO_COMP_NAME' ).AsString;
      aSo.DbLoginUser := DataReader.FieldByName( 'SO_LOGINUSER' ).AsString;
      aSo.DbLoginPass := DataReader.FieldByName( 'SO_LOGINPASS' ).AsString;
      aSo.DbAliase := DataReader.FieldByName( 'SO_DBALIASE' ).AsString;
      aSo.DbConnectStatus := dbNoSelect;
      if aSo.Selected then aSo.DbConnectStatus := dbNone;
      aSo.ComDbLoginUser := DataReader.FieldByName( 'CA_LOGINUSER' ).AsString;
      aSo.ComDbLoginPass := DataReader.FieldByName( 'CA_LOGINPASS' ).AsString;
      aSo.ComDbAliase := DataReader.FieldByName( 'CA_DBALIASE' ).AsString;
      aSo.ComDbConnectStatus := dbNone;
      FSoList.Add( aSo );
      FMsgSubject.Normal( Format( '讀取系統台[%s]設定中。', [aSo.CompName] ) );
      Delay( 100 );
      DataReader.Next;
    end;
    Result := True;
  except
    on E: Exception do
    begin
      FSoList.Clear;
      FMsgSubject.Error( Format( '設定檔讀取錯誤, 訊息:%s。', [E.Message] ) );
      Result := False;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetSoList: Boolean;
var
  aIndex: Integer;
  aSo: TAppSo;
begin
  Result := True;
  DataReader.SQL.Text := ' DELETE FROM SO_COMP ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO SO_COMP (         ' +
    '    SO_SELECTED,               ' +
    '    SO_POS,                    ' +
    '    SO_COMP_CODE,              ' +
    '    SO_COMP_NAME,              ' +
    '    SO_LOGINUSER,              ' +
    '    SO_LOGINPASS,              ' +
    '    SO_DBALIASE,               ' +
    '    CA_LOGINUSER,              ' +
    '    CA_LOGINPASS,              ' +
    '    CA_DBALIASE   )            ' +
    ' VALUES (                      ' +
    '    :1,                        ' +
    '    :2,                        ' +
    '    :3,                        ' +
    '    :4,                        ' +
    '    :5,                        ' +
    '    :6,                        ' +
    '    :7,                        ' +
    '    :8,                        ' +
    '    :9,                        ' +
    '    :10   )                    ';
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSo := FSoList[aIndex];
    DataReader.Parameters.ParamByName( '1' ).Value := 'N';
    if aSo.Selected then
      DataReader.Parameters.ParamByName( '1' ).Value := 'Y';
    DataReader.Parameters.ParamByName( '2' ).Value := aSo.Pos;
    DataReader.Parameters.ParamByName( '3' ).Value := aSo.CompCode;
    DataReader.Parameters.ParamByName( '4' ).Value := aSo.CompName;
    DataReader.Parameters.ParamByName( '5' ).Value := aSo.DbLoginUser;
    DataReader.Parameters.ParamByName( '6' ).Value := aSo.DbLoginPass;
    DataReader.Parameters.ParamByName( '7' ).Value := aSo.DbAliase;
    DataReader.Parameters.ParamByName( '8' ).Value := aSo.ComDbLoginUser;
    DataReader.Parameters.ParamByName( '9' ).Value := aSo.ComDbLoginPass;
    DataReader.Parameters.ParamByName( '10' ).Value := aSo.ComDbAliase;
    try
      DataReader.ExecSQL;
      FMsgSubject.OK( Format( '系統台[%s]設定存檔完成。', [aSo.CompName] ) );
      Delay( 100 );
    except
      on E: Exception do
      begin
        Result := False;
        FMsgSubject.Error( Format( '系統台設定存檔失敗, 訊息:%s。', [E.Message] ) );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ShowConfigForm: TModalResult;
begin
  fmConfig := TfmConfig.Create( nil );
  try
    Result := fmConfig.ShowModal;
    if ( Result = mrOK ) then
    begin
      Screen.Cursor := crHourGlass;
      try
        ReadSettingValue;
        SaveToFile( FConfigFileName );
      finally
        Screen.Cursor := crDefault;
      end;
    end;
  finally
    fmConfig.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ReadSettingValue;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aSo: TAppSo;
begin
  FSoList.Clear;
  for aIndex := 0 to fmConfig.SoTree.Nodes.Count - 1 do
  begin
    if ( fmConfig.SoTree.Nodes.Items[aIndex].Level = 0 ) then Continue;
    aNode := fmConfig.SoTree.Nodes[aIndex];
    aSo := TAppSo.Create;
    aSo.Selected := ( aNode.Values[fmConfig.SoTreeColumn1.ItemIndex] = 'Y' );
    aSo.Pos := StrToIntDef( aNode.Texts[fmConfig.SoTreeColumn9.ItemIndex], 1 );
    aSo.PosName := aNode.Parent.Texts[fmConfig.SoTreeColumn3.ItemIndex];
    aSo.CompCode := aNode.Texts[fmConfig.SoTreeColumn2.ItemIndex];
    aSo.CompName := aNode.Texts[fmConfig.SoTreeColumn3.ItemIndex];
    aSo.DbLoginUser := aNode.Texts[fmConfig.SoTreeColumn5.ItemIndex];
    aSo.DbLoginPass := aNode.Texts[fmConfig.SoTreeColumn6.ItemIndex];
    aSo.DbAliase := aNode.Texts[fmConfig.SoTreeColumn7.ItemIndex];
    aSo.DbConnectStatus := dbNone;
    if not aSo.Selected then aSo.DbConnectStatus := dbNoSelect;
    aSo.ComDbLoginUser := aNode.Texts[fmConfig.SoTreeColumn10.ItemIndex];
    aSo.ComDbLoginPass := aNode.Texts[fmConfig.SoTreeColumn11.ItemIndex];
    aSo.ComDbAliase := aNode.Texts[fmConfig.SoTreeColumn12.ItemIndex];
    aSo.ComDbConnectStatus := dbNone;
    FSoList.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
