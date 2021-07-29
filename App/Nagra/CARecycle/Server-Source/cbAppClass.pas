unit cbAppClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, SyncObjs, ComObj, Contnrs,
  DB, ADODB, cbClass;

type

  TDbConnectStatus = ( dbNoSelect, dbNone, dbOK, dbError, dbActive );

  TAppSo = class(TSo)
  private
    FPos: Integer;
    FPosName: String;
    FDbConnectStatus: TDbConnectStatus;
    FComDbConnectStatus: TDbConnectStatus;
    FDisplayNode: Pointer;
    FDataModule: TDataModule;
    FComDataModule: TDataModule;
    FLastDbErrTime: TDateTime;
    FLastComDbErrTime: TDateTime;
    FFirstQuery: Boolean;
    FComDbLoginUser: String;
    FComDbLoginPass: String;
    FComDbAliase: String;
    procedure SetDbConnectStatus(const aValue: TDbConnectStatus);
    procedure SetComDbConnectStatus(const aValue: TDbConnectStatus);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    property Pos: Integer read FPos write FPos;
    property PosName: String read FPosName write FPosName;
    property DbConnectStatus: TDbConnectStatus read FDbConnectStatus write SetDbConnectStatus;
    property DisplayNode: Pointer read FDisplayNode write FDisplayNode;
    property DataModule: TDataModule read FDataModule write FDataModule;
    property LastDbErrTime: TDateTime read FLastDbErrTime;
    property LastComDbErrTime: TDateTime read FLastComDbErrTime;
    property FirstQuery: Boolean read FFirstQuery write FFirstQuery;
    property ComDbLoginUser: String read FComDbLoginUser write FComDbLoginUser;
    property ComDbLoginPass: String read FComDbLoginPass write FComDbLoginPass;
    property ComDbAliase: String read FComDbAliase write FComDbAliase;
    property ComDbConnectStatus: TDbConnectStatus read FComDbConnectStatus write SetComDbConnectStatus;
    property ComDataModule: TDataModule read FComDataModule write FComDataModule;
  end;


  TAppSoList = class(TSoList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TAppSo;
    procedure SetItem(Index: Integer; aObj: TAppSo);
  public
    property Items[Index: Integer]: TAppSo read GetItem write SetItem; default;
  end;


  TRecycleCommand = class
  private
    FSeqNo: String;
    FStbAutoFlag: Integer;
    FCompCode: Integer;
    FCompName: String;
    FStbFlag: Integer;
    FCMDFlag: Integer;
    FIccNo: String;
    FConfirm: String;
    FTransNum: String;
    FRecycleText: String;
    FErrMsg: String;
    FOperator: String;
    FStbNo: String;
    FUpdTime: TDateTime;
    FRCStartDate: String;
    FRCEndDate: String;
    FLastDownloadTime: String;
    FData: Pointer;
  public
    property SeqNo: String read FSeqNo write FSeqNo;
    property IccNo: String read FIccNo write FIccNo;
    property StbNo: String read FStbNo write FStbNo;
    property StbFlag: Integer read FStbFlag write FStbFlag;
    property StbAutoFlag: Integer read FStbAutoFlag write FStbAutoFlag;
    property CompCode: Integer read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property RCStartDate: String read FRCStartDate write FRCStartDate;
    property RCEndDate: String read FRCEndDate write FRCEndDate;
    property Operator: String read FOperator write FOperator;
    property UpdTime: TDateTime read FUpdTime write FUpdTime;
    property LastDownloadTime: String read FLastDownloadTime write FLastDownloadTime;
    property CMDFlag: Integer read FCMDFlag write FCMDFlag;
    property RecycleText: String read FRecycleText write FRecycleText;
    property TransNum: String read FTransNum write FTransNum;
    property Confirm: String read FConfirm write FConfirm;
    property ErrMsg: String read FErrMsg write FErrMsg;
    property Data: Pointer read FData write FData;
  end;


  TRecycleCommandList = class(TList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TRecycleCommand;
    procedure SetItem(Index: Integer; aCmd: TRecycleCommand);
  public
    procedure Insert(Index: Integer; aCmd: TRecycleCommand);
    function Add(aCmd: TRecycleCommand): Integer;
    function Extract(aCmd: TRecycleCommand): TRecycleCommand;
    function Remove(aCmd: TRecycleCommand): Integer;
    function IndexOf(aCmd: TRecycleCommand): Integer; overload;
    function IndexOf(aSeqNo, aIccNo: String): Integer; overload;
    function FindBySeq(aSeq: String): Integer;
    function FindByIccNo(aIccNo: String): Integer;
    function First: TRecycleCommand;
    function Last: TRecycleCommand;
    property Items[Index: Integer]: TRecycleCommand read GetItem write SetItem; default;
  end;



  TCommandHelper = class
  private
    FCmdMapping: TStringList;
    FTransMapping: TStringList;
    FCommand: TRecycleCommand;
    procedure ChangeCommandMapping;
    function GetCmd(Index: Integer): Char;
    procedure SetCmd(Index: Integer; const Value: Char);
    function GetTrans(Index: Integer): string;
    procedure SetTrans(Index: Integer; const Value: string);
    procedure SetRecycleCommand(const Value: TRecycleCommand);
  public
    constructor Create;
    destructor Destory;
    procedure CmdOk(const aCmd: String);
    procedure CmdErr(const aCmd, aErrMsg: String);
    procedure CmdProc(const aCmd, aTrans: String);
    function FindTextByCmd(const aCmd: String): String;
    function FindTransByCmd(const aCmd: String): String;
    function FindIndexByCmd(const aCmd: String): Integer;
    function NextExecuteCmd: String;
    function CurrentExecuteCmd: String;
    property Cmd[Index: Integer]: Char read GetCmd write SetCmd;
    property Trans[Index: Integer]: string read GetTrans write SetTrans;
    property RecycleCommand: TRecycleCommand read FCommand write SetRecycleCommand;
  end;


implementation

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

{ TAppSo }

constructor TAppSo.Create;
begin
  inherited Create;
  FPos := -1;
  FPosName := EmptyStr;
  FDbConnectStatus := dbNoSelect;
  FComDbConnectStatus := dbNone;
  FDisplayNode := nil;
  FDataModule := nil;
  FLastDbErrTime := 0;
  FLastComDbErrTime := 0;
  FFirstQuery := True;
end;

{ ---------------------------------------------------------------------------- }

destructor TAppSo.Destroy;
begin
  FDisplayNode := nil;
  FDataModule := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSo.SetDbConnectStatus(const aValue: TDbConnectStatus);
begin
  FDbConnectStatus := aValue;
  FLastDbErrTime := 0;
  if FDbConnectStatus = dbError then FLastDbErrTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSo.SetComDbConnectStatus(const aValue: TDbConnectStatus);
begin
  FComDbConnectStatus := aValue;
  FLastComDbErrTime := 0;
  if FComDbConnectStatus = dbError then FLastComDbErrTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSo.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TSo ) then
  begin
    TAppSo( Dest ).Pos := FPos;
    TAppSo( Dest ).PosName := FPosName;
    TAppSo( Dest ).FDbConnectStatus := FDbConnectStatus;
    TAppSo( Dest ).DisplayNode := FDisplayNode;
    TAppSo( Dest ).DataModule := FDataModule;
    TAppSo( Dest ).FLastDbErrTime := FLastDbErrTime;
    TAppSo( Dest ).FirstQuery := FFirstQuery;
    TAppSo( Dest ).ComDbLoginUser := FComDbLoginUser;
    TAppSo( Dest ).ComDbLoginPass := FComDbLoginPass;
    TAppSo( Dest ).ComDbAliase := FComDbAliase;
    TAppSo( Dest ).ComDataModule := FComDataModule;
  end;
  inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TAppSoList }

procedure TAppSoList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
  begin
    TAppSo( Ptr ).Free;
    Ptr := nil;
  end;  
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TAppSoList.GetItem(Index: Integer): TAppSo;
begin
  Result := TAppSo( inherited GetItem( Index ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSoList.SetItem(Index: Integer; aObj: TAppSo);
begin
  inherited SetItem( Index, aObj );
end;

{ ---------------------------------------------------------------------------- }

{ TRecycleCommandList }

procedure TRecycleCommandList.Notify(Ptr: Pointer;
  Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
    TRecycleCommand( Ptr ).Free;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.Add(aCmd: TRecycleCommand): Integer;
begin
  Result := inherited Add( aCmd );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.Extract(
  aCmd: TRecycleCommand): TRecycleCommand;
begin
  Result := TRecycleCommand( inherited Extract( aCmd ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommandList.Insert(Index: Integer;
  aCmd: TRecycleCommand);
begin
  inherited Insert( Index, aCmd );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.Remove(aCmd: TRecycleCommand): Integer;
begin
  Result := inherited Remove( aCmd );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.First: TRecycleCommand;
begin
  Result := TRecycleCommand( inherited First );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.Last: TRecycleCommand;
begin
  Result := TRecycleCommand( inherited Last );
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommandList.SetItem(Index: Integer;
  aCmd: TRecycleCommand);
begin
  inherited Items[Index] := aCmd;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.GetItem(Index: Integer): TRecycleCommand;
begin
  Result := TRecycleCommand( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.IndexOf(aCmd: TRecycleCommand): Integer;
begin
  Result := inherited IndexOf( aCmd );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.IndexOf(aSeqNo, aIccNo: String): Integer;
begin
  Result := 0;
  while ( Result < Count ) and
        ( ( TRecycleCommand( List^[Result] ).SeqNo <> aSeqNo ) or
          ( TRecycleCommand( List^[Result] ).IccNo <> aIccNo ) )  do
    Inc(Result);
  if Result = Count then
    Result := -1;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.FindByIccNo(aIccNo: String): Integer;
begin
  Result := 0;
  while ( Result < Count ) and
        ( TRecycleCommand( List^[Result] ).IccNo <> aIccNo ) do
  begin
    Inc( Result );
  end;
  if Result = Count then
    Result := -1;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommandList.FindBySeq(aSeq: String): Integer;
var
  aIndex: Integer;
begin
  Result := -1;
  for aIndex := Low( List^ ) to High( List^ ) do
  begin
    if Pos( aSeq, TRecycleCommand( List^[aIndex] ).TransNum ) > 0 then
    begin
      Result := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TCommandHelper }

constructor TCommandHelper.Create;
begin
  FCmdMapping := TStringList.Create;
  FTransMapping := TStringList.Create;
  FCommand := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TCommandHelper.Destory;
begin
  FTransMapping.Free;
  FCmdMapping.Free;
  FCommand := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.GetCmd(Index: Integer): Char;
begin
  if ( Index < 1 ) or ( Index > Length( FCommand.RecycleText ) ) then
    Exception.Create( 'RecycleText index out of bounds.' );
  Result := FCommand.RecycleText[Index];
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.SetCmd(Index: Integer; const Value: Char);
var
  aText: String;
begin
  if ( Index < 1 ) or ( Index > Length( FCommand.RecycleText ) ) then
    Exception.Create( 'RecycleText index out of bounds.' );
  aText := FCommand.RecycleText;
  aText[Index] := Value;
  FCommand.RecycleText := aText;
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.GetTrans(Index: Integer): string;
begin
  FTransMapping.DelimitedText := FCommand.TransNum;
  Result := FTransMapping[Index];
  FTransMapping.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.SetTrans(Index: Integer; const Value: string);
begin
  FTransMapping.DelimitedText := FCommand.TransNum;
  FTransMapping[Index] := Value;
  FCommand.TransNum := FTransMapping.DelimitedText;
  FTransMapping.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.CmdErr(const aCmd, aErrMsg: String);
var
  aIndex: Integer;
  aText: String;
begin
  aIndex := FCmdMapping.IndexOf( aCmd );
  if ( aIndex < 0 ) then
    raise Exception.Create( 'CmdMapping out of bound¡C' );
  aText := FCommand.RecycleText;
  aText[aIndex+1] := '2';
  FCommand.RecycleText := aText;
  FCommand.ErrMsg := aErrMsg;
  FCommand.CMDFlag := 4;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.CmdOk(const aCmd: String);
var
  aIndex: Integer;
  aText: String;
begin
  aIndex := FCmdMapping.IndexOf( aCmd );
  if ( aIndex < 0 ) then
    raise Exception.Create( 'CmdMapping out of bound¡C' );
  aText := FCommand.RecycleText;
  aText[aIndex+1] := '1';
  FCommand.RecycleText := aText;
  if ( FCommand.RecycleText = Lpad( EmptyStr, FCmdMapping.Count - 1, '1' ) + '0' ) then
    FCommand.CMDFlag := 2
  else if ( FCommand.RecycleText = Lpad( EmptyStr, FCmdMapping.Count, '1' ) ) then
    FCommand.CMDFlag := 3; 
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.CmdProc(const aCmd, aTrans: String);
var
  aIndex: Integer;
  aText: String;
begin
  aIndex := FCmdMapping.IndexOf( aCmd );
  if ( aIndex < 0 ) then
    raise Exception.Create( 'CmdMapping out of bound¡C' );
  aText := FCommand.RecycleText;
  aText[aIndex+1] := '3';
  FCommand.RecycleText := aText;
  SetTrans( aIndex, aTrans );
  if ( FCommand.CMDFlag in [0,1] ) then
    FCommand.CMDFlag := 1;
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.FindTextByCmd(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FCmdMapping.IndexOf( aCmd );
  if ( aIndex < 0 ) then
     raise Exception.Create( 'CmdMapping out of bound¡C' );
  Result := FCommand.RecycleText[aIndex+1];
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.FindTransByCmd(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FCmdMapping.IndexOf( aCmd );
  if ( aIndex < 0 ) then
     raise Exception.Create( 'CmdMapping out of bound¡C' );
  FTransMapping.DelimitedText := FCommand.TransNum;
  Result := FTransMapping[aIndex];
  FTransMapping.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.FindIndexByCmd(const aCmd: String): Integer;
begin
  Result := FCmdMapping.IndexOf( aCmd );
  if Result >= 0 then Inc( Result );
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.NextExecuteCmd: String;
var
  aIndex: Integer;
  aPreCmd: String;
begin
  Result := EmptyStr;
  aPreCmd := '1';
  for aIndex := 1 to Length( FCommand.RecycleText ) do
  begin
    if ( aIndex > 1 ) then aPreCmd := FCommand.RecycleText[aIndex-1];
    if ( FCommand.RecycleText[aIndex] = '0' ) and ( aPreCmd = '1' ) then
    begin
      Result := FCmdMapping[aIndex-1];
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandHelper.CurrentExecuteCmd: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  if Pos( '2', FCommand.RecycleText ) > 0 then Exit;
  for aIndex := 1 to Length( FCommand.RecycleText ) do
  begin
    if ( FCommand.RecycleText[aIndex] = '3' ) then
    begin
      Result := FCmdMapping[aIndex-1];
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.ChangeCommandMapping;
begin
  FCmdMapping.Clear;
  if ( FCommand.StbAutoFlag = 0 ) then
  begin
    FCmdMapping.Add( '52_1' );  // Paire
    FCmdMapping.Add( '8' );     // Set Credit = 0
    FCmdMapping.Add( '97_96' ); // Set IPPV Reported and Purge
    FCmdMapping.Add( '7' );     // Cancel All Product
    FCmdMapping.Add( '48' );    // Set ZipCode
    FCmdMapping.Add( '52_2' );  // UnPaire
  end else
  begin
    FCmdMapping.Add( '52_1' );  // Paire
    FCmdMapping.Add( '62' );    // disable auto callback
    FCmdMapping.Add( '8' );     // Set Credit = 0
    FCmdMapping.Add( '60' );    // immediate callback
    FCmdMapping.Add( '7' );     // Cancel All Product
    FCmdMapping.Add( '48' );    // Set ZipCode
    FCmdMapping.Add( '99' );    // wait for callback
    FCmdMapping.Add( '52_2' );  // UnPaire
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandHelper.SetRecycleCommand(const Value: TRecycleCommand);
begin
  FCommand := Value;
  ChangeCommandMapping;
end;

{ ---------------------------------------------------------------------------- }

end.
