unit cbAppClass;

interface

uses
  SysUtils, Classes;

type

  TSo = class
  private
    FCompCode: String;
    FPassword: String;
    FCompName: String;
    FAliase: String;
    FUserId: String;
    FCanSwitch: Boolean;
    FActive: Boolean;
    procedure SetActive(const Value: Boolean);
  public
    constructor Create;
    destructor Destroy; override;
    property UserId: String read FUserId write FUserId;
    property Password: String read FPassword write FPassword;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property Aliase: String read FAliase write FAliase;
    property Active: Boolean read FActive write SetActive;
    property CanSwitch: Boolean read FCanSwitch write FCanSwitch; 
  end;



  TUser = class(TPersistent)
  private
    FUserId: String;
    FUserName: String;
    FPassword: String;
    FCompCode: String;
    FCompName: String;
    FComputerName: String;
    FAdmin: Boolean;
    FIsWareHouse: Boolean;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    property UserId: String read FUserId write FUserId;
    property UserName: String read FUserName write FUserName;
    property Password: String read FPassword write FPassword;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property ComputerName: String read FComputerName write FComputerName;
    property Admin: Boolean read FAdmin write FAdmin;
    property IsWareHouse: Boolean read FIsWareHouse write FIsWareHouse;
  end;


  TRecycleCommand = class
  private
    FList: TStringList;
    FText: String;
    FTrans: String;
    FAutoCallBack: Integer;
    function GetDefaultText: String;
    function GetDefaultTrans: String;
    procedure SetAutoCallBack(const Value: Integer);
    procedure ChangeCmdList;
  public
    constructor Create(const aAutoCallBack: Integer);
    destructor Destroy; override;
    function CmdOkText(const aCmd: String): String;
    function CmdErrText(const aCmd: String): String;
    function CmdResetText(const aCmd: String): String;
    function FindTextByCmd(const aCmd: String): String;
    function SetTransTextByCmd(const aCmd, aTrans: String): String;
    function FindTransByCmd(const aCmd: String): String;
    procedure ResetCmdText;
    procedure ResetTransText;
    property AutoCallBack: Integer read FAutoCallBack write SetAutoCallBack;
    property CmdText: String read FText write FText;
    property TransText: String read FTrans write FTrans;
  end;


  TCARecycle = class(TPersistent)
  private
    FStbFlag: Integer;
    FCompCode: Integer;
    FStbAutoFlag: Integer;
    FCMDFlag: Integer;
    FSeqNo: String;
    FIccNo: String;
    FConfirm: String;
    FStbNo: String;
    FOperator: String;
    FRCEndDate: TDateTime;
    FRCStattDate: TDateTime;
    FLastDownloadTime: TDateTime;
    FUpdTime: TDateTime;
    FErrMsg: String;
    FCmd: TRecycleCommand;
    function GetRecycleText: String;
    function GetTransText: String;
    procedure SetRecycleText(const Value: String);
    procedure SetTransText(const Value: String);
    procedure SetStbAutoFlag(const Value: Integer);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    procedure ResetRecycleText;
    procedure ResetTransNum;
    property SeqNo: String read FSeqno write FSeqno; 
    property IccNo: String read FIccNo write FIccNo;
    property StbNo: String read FStbNo write FStbNo;
    property StbFlag: Integer read FStbFlag write FStbFlag;
    property StbAutoFlag: Integer read FStbAutoFlag write SetStbAutoFlag;
    property CompCode: Integer read FCompCode write FCompCode;
    property RCStattDate: TDateTime read FRCStattDate write FRCStattDate;
    property RCEndDate: TDateTime read FRCEndDate write FRCEndDate;
    property Operator: String read FOperator write FOperator;
    property UpdTime: TDateTime read FUpdTime write FUpdTime;
    property LastDownloadTime: TDateTime read FLastDownloadTime write FLastDownloadTime;
    property CMDFlag: Integer read FCMDFlag write FCMDFlag;
    property RecycleText: String read GetRecycleText write SetRecycleText;
    property TransNum: String read GetTransText write SetTransText;
    property Confirm: String read FConfirm write FConfirm;
    property ErrMsg: String read FErrMsg write FErrMsg;
  end;


implementation

uses cbUtilis;

{ TUser }

{ ---------------------------------------------------------------------------- }

constructor TUser.Create;
begin
  FUserId := EmptyStr;
  FUserName := EmptyStr;
  FPassword := EmptyStr;
  FCompCode := EmptyStr;
  FCompName := EmptyStr;
  FComputerName := EmptyStr;
  FAdmin := False;
  FIsWareHouse := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TUser.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TUser ) then 
  begin
    TUser( Dest ).UserId := Self.UserId;
    TUser( Dest ).UserName := Self.UserName;
    TUser( Dest ).Password := Self.Password;
    TUser( Dest ).CompCode := Self.CompCode;
    TUser( Dest ).CompName := Self.CompName;
    TUser( Dest ).ComputerName := Self.ComputerName;
    TUser( Dest ).Admin := Self.Admin;
    TUser( Dest ).IsWareHouse := Self.IsWareHouse;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

destructor TUser.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

{ TCARecycle }

constructor TCARecycle.Create;
begin
  inherited Create;
  FConfirm := 'N';
  FStbFlag := 0;
  FCompCode := -1;
  FStbAutoFlag := 0;
  FCMDFlag := 0;
  FIccNo := EmptyStr;
  FStbNo := EmptyStr;
  FOperator := EmptyStr;
  FRCStattDate := 0;
  FRCEndDate := 0;
  FLastDownloadTime := 0;
  FUpdTime := 0;
  FErrMsg := EmptyStr;
  FCmd := TRecycleCommand.Create( FStbAutoFlag );
end;

{ ---------------------------------------------------------------------------- }

destructor TCARecycle.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TCARecycle.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TCARecycle ) then
  begin
    TCARecycle( Dest ).StbFlag := Self.StbFlag;
    TCARecycle( Dest ).CompCode := Self.CompCode;
    TCARecycle( Dest ).StbAutoFlag := Self.FStbAutoFlag;
    TCARecycle( Dest ).CMDFlag := Self.CMDFlag;
    TCARecycle( Dest ).RecycleText := Self.RecycleText;
    TCARecycle( Dest ).SeqNo := Self.SeqNo;
    TCARecycle( Dest ).IccNo := Self.IccNo;
    TCARecycle( Dest ).Confirm := Self.Confirm;
    TCARecycle( Dest ).StbNo := Self.StbNo;
    TCARecycle( Dest ).Operator := Self.Operator;
    TCARecycle( Dest ).TransNum := Self.TransNum;
    TCARecycle( Dest ).RCEndDate := Self.RCEndDate;
    TCARecycle( Dest ).RCStattDate := Self.RCStattDate;
    TCARecycle( Dest ).LastDownloadTime := Self.LastDownloadTime;
    TCARecycle( Dest ).UpdTime := Self.UpdTime;
    TCARecycle( Dest ).ErrMsg := Self.ErrMsg;
  end else
    inherited AssignTo( Dest );

end;
{ ---------------------------------------------------------------------------- }

procedure TCARecycle.ResetRecycleText;
begin
  FCmd.ResetCmdText;
end;

{ ---------------------------------------------------------------------------- }

procedure TCARecycle.ResetTransNum;
begin
  FCmd.ResetTransText;
end;

{ ---------------------------------------------------------------------------- }

function TCARecycle.GetRecycleText: String;
begin
  Result := FCmd.CmdText;
end;

{ ---------------------------------------------------------------------------- }

function TCARecycle.GetTransText: String;
begin
  Result := FCmd.TransText;
end;

{ ---------------------------------------------------------------------------- }

procedure TCARecycle.SetRecycleText(const Value: String);
begin
  FCmd.CmdText := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TCARecycle.SetTransText(const Value: String);
begin
  FCmd.TransText := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TCARecycle.SetStbAutoFlag(const Value: Integer);
begin
  FStbAutoFlag := Value;
  if ( FStbAutoFlag <> FCmd.AutoCallBack ) then
    FCmd.AutoCallBack := FStbAutoFlag;
end;

{ ---------------------------------------------------------------------------- }

{ TRecycleCommand }

constructor TRecycleCommand.Create(const aAutoCallBack: Integer);
begin
  inherited Create;
  FList := TStringList.Create;
  FAutoCallBack := aAutoCallBack;
  ChangeCmdList;
  FText := GetDefaultText;
  FTrans := GetDefaultTrans;
end;

{ ---------------------------------------------------------------------------- }

destructor TRecycleCommand.Destroy;
begin
  FList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.CmdErrText(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
  begin
    FText[aIndex+1] := '2';
  end else
  begin
    raise Exception.Create( 'CmdErrText Index -1¡C' );
  end;
  Result := FText;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.CmdOkText(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
  begin
    FText[aIndex+1] := '1';
  end else
  begin
    raise Exception.Create( 'CmdOkText Index -1¡C' );
  end;
  Result := FText;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.CmdResetText(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
  begin
    FText[aIndex+1] := '0';
  end else
  begin
    raise Exception.Create( 'CmdResetText Index -1¡C' );
  end;
  Result := FText;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.FindTextByCmd(const aCmd: String): String;
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
    Result := FText[aIndex+1]
  else begin
    raise Exception.Create( 'FindTextByCmd Index -1¡C' );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.SetTransTextByCmd(const aCmd, aTrans: String): String;
var
  aIndex: Integer;
  aTempList: TStringList;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
  begin
    aTempList := TStringList.Create;
    try
      aTempList.Delimiter := ',';
      aTempList.DelimitedText := FTrans;
      aTempList.Strings[aIndex] := aTrans;
      FTrans := aTempList.DelimitedText;
      Result := FTrans;
    finally
      aTempList.Free;
    end;
  end else
  begin
    raise Exception.Create( 'SetTransTextByCmd Index -1¡C' );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.FindTransByCmd(const aCmd: String): String;
var
  aIndex: Integer;
  aTempList: TStringList;
begin
  aIndex := FList.IndexOf( aCmd );
  if ( aIndex >= 0 ) then
  begin
    aTempList := TStringList.Create;
    try
      aTempList.Delimiter := ',';
      aTempList.DelimitedText := FTrans;
      Result := aTempList.Strings[aIndex];
    finally
      aTempList.Free;
    end;
  end else
  begin
    raise Exception.Create( 'FindTransByCmd Index -1¡C' );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.GetDefaultText: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 0 to FList.Count - 1 do
    Result := ( Result + '0' );
  if ( Result = EmptyStr ) then
    raise Exception.Create( 'Default Text Is Empty¡C' );
end;

{ ---------------------------------------------------------------------------- }

function TRecycleCommand.GetDefaultTrans: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 0 to FList.Count - 1 do
  begin
    Result := ( Result + '000000000' );
    if ( aIndex < ( FList.Count - 1 ) ) then Result := ( Result + ',' ); 
  end;
  if ( Result = EmptyStr ) then
    raise Exception.Create( 'Default Trans Is Empty¡C' );
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommand.ResetCmdText;
begin
  FText := GetDefaultText;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommand.ResetTransText;
begin
  FTrans := GetDefaultTrans;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommand.SetAutoCallBack(const Value: Integer);
begin
  if ( FAutoCallBack <> Value ) then
  begin
    FAutoCallBack := Value;
    ChangeCmdList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecycleCommand.ChangeCmdList;
begin
  FList.Clear;
  if ( FAutoCallBack = 0 ) then
  begin
    FList.Add( '52_1' );  // Paire
    FList.Add( '8' );     // Set Credit = 0
    FList.Add( '97_96' ); // Set IPPV Reported and Purge
    FList.Add( '7' );     // Cancel All Product
    FList.Add( '48' );    // Set ZipCode
    FList.Add( '52_2' );  // UnPaire
  end else
  begin
    FList.Add( '52_1' );  // Paire
    FList.Add( '62' );    // disable auto callback
    FList.Add( '8' );     // Set Credit = 0
    FList.Add( '60' );    // immediate callback
    FList.Add( '7' );     // Cancel All Product
    FList.Add( '48' );    // Set ZipCode
    FList.Add( '99' );    // wait for callback
    FList.Add( '52_2' );  // UnPaire
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSo }

constructor TSo.Create;
begin
  FUserId := EmptyStr;
  FCompCode := EmptyStr;
  FPassword := EmptyStr;
  FCompName := EmptyStr;
  FAliase := EmptyStr;
  FActive := False;
  FCanSwitch := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TSo.Destroy;
begin
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.SetActive(const Value: Boolean);
begin
  FActive := Value;
  if ( FActive ) and ( not FCanSwitch ) then FCanSwitch := True;
end;

{ ---------------------------------------------------------------------------- }

end.
