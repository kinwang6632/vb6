unit utlUCU;

interface

uses
  Classes, SysUtils, ComCtrls, Forms, IniFiles;

type

  TDataFile = class(TObject)
  public
    Id: Integer;
    FileName: string;
    Desc: string;
    AutoReg: Boolean;
    function GetString: string;
  end;

  TFunc = class(TObject)
  private
    FDataFiles: TStrings;
    function GetDataFile(Idx: Integer): TDataFile;
    function GetDataFileCount: Integer;
  public
    Id: Integer;
    Desc: string;
    ExeName: string;
    Param: string;
    IconName: string;
    constructor Create;
    destructor Destroy; override;
    procedure AddDataFile(DataFile: TDataFile);
    procedure DeleteDataFile(DataFile: TDataFile);
    property DataFileCount: Integer read GetDataFileCount;
    property DataFiles[Idx: Integer]: TDataFile read GetDataFile;
    function HasDataFile(DataFile: TDataFile): Boolean;
    procedure DeleteAllDataFile;
    function GetString: string;
  end;

  TApp = class(TObject)
  private
    FFuncs: TStrings;
    function GetFunc(Idx: Integer): TFunc;
    function GetFuncCount: Integer;
  public
    Id: Integer;
    Desc: string;
    ParentId: Integer;
    constructor Create;
    destructor Destroy; override;
    property FuncCount: Integer read GetFuncCount;
    property Funcs[Idx: Integer]: TFunc read GetFunc;
    procedure AddFunc(Func: TFunc);
    procedure DeleteFunc(Func: TFunc);
    function HasFunc(Func: TFunc): Boolean;
    function GetString: string;
  end;

  TUser = class(TObject)
  public
    Id: Integer;
    Code: string;
    Desc: string;
    Password: string;
    function GetString: string;
  end;

  TUserGrp = class(TObject)
  private
    FUsers: TStrings;
    FFuncs: TStrings;
    function GetUser(Idx: Integer): TUser;
    function GetUserCount: Integer;
    function GetFunc(Idx: Integer): TFunc;
    function GetFuncCount: Integer;
  public
    Id: Integer;
    Desc: string;
    ParentId: Integer;
    constructor Create;
    destructor Destroy; override;
    property UserCount: Integer read GetUserCount;
    property Users[Idx: Integer]: TUser read GetUser;
    procedure AddUser(User: TUser);
    procedure DeleteUser(User: TUser);
    property FuncCount: Integer read GetFuncCount;
    property Funcs[Idx: Integer]: TFunc read GetFunc;
    procedure AddFunc(Func: TFunc);
    procedure DeleteFunc(Func: TFunc);
    procedure DeleteAllFunc;
    function HasFunc(Func: TFunc): Boolean;
    function HasUser(User: TUser): BOolean;
    function GetString: string;
  end;

  TUCMgr = class(TObject)
  private
    FIniFile: TiniFile;
    FMaxAppId, FMaxFuncId, FMaxDataFileId: Integer;
    FMaxUserGrpId, FMaxUserId: Integer;
    FAllApps, FAllUserGrps: TTreeNodes;
    FAllFuncs, FAllUsers, FAllDataFiles: TStrings;

    function GetMaxAppId: Integer;
    function GetMaxDataFileId: Integer;
    function GetMaxFuncId: Integer;
    function GetMaxUserGrpId: Integer;
    function GetMaxUserId: Integer;

    function GetAppCount: Integer;
    function GetDataFileCount: Integer;
    function GetFuncCount: Integer;
    function GetUserCount: Integer;
    function GetUserGrpCount: Integer;

    function GetFunc(Idx: Integer): TFunc;
    function GetApp(Idx: Integer): TApp;
    function GetDataFile(Idx: Integer): TDataFile;
    function GetUser(Idx: Integer): TUser;
    function GetUserGrp(Idx: Integer): TUserGrp;

    function ExtractApp(Str: string): TApp;
    function ExtractFunc(Str: string): TFunc;
    function ExtractDataFile(Str: string): TDataFile;
    function ExtractUserGrp(Str: string): TUserGrp;
    function ExtractUser(Str: string): TUser;

    procedure GetAllUserGrps;
    procedure GetAllUsers;
    procedure GetAllDataFiles;
    procedure GetAllFuncs;
    procedure GetAllApps;

    procedure GetFuncsOfUserGrp(UserGrp: TUserGrp; UserGrpIdx: Integer);
    procedure GetFuncsOfApp(App: TApp; AppIdx: Integer);
    procedure GetDataFilesOfFunc(Func: TFunc; FuncIdx: Integer);
    procedure GetUsersOfUserGrp(UserGrp: TUserGrp; UserGrpIdx: Integer);

    function GetParentAppNode(ParentId: Integer): TTreeNode;
    function GetParentUserGrpNode(ParentId: Integer): TTreeNode;

    procedure ClearFile;

  public
    property MaxAppId: Integer read GetMaxAppId;
    property MaxFuncId: Integer read GetMaxFuncId;
    property MaxUserGrpId: Integer read GetMaxUserGrpId;
    property MaxUserId: Integer read GetMaxUserId;
    property MaxDataFileId: Integer read GetMaxDataFileId;

    constructor Create(AllApps, AllUserGrps: TTreeNodes);
    destructor Destroy; override;

    property Apps[Idx: Integer]: TApp read GetApp;
    property AppCount: Integer read GetAppCount;
    property Funcs[Idx: Integer]: TFunc read GetFunc;
    property FuncCount: Integer read GetFuncCount;
    property UserGrps[Idx: Integer]: TUserGrp read GetUserGrp;
    property UserGrpCount: Integer read GetUserGrpCount;
    property Users[Idx: Integer]: TUser read GetUser;
    property UserCount: Integer read GetUserCount;
    property DataFiles[Idx: Integer]: TDataFile read GetDataFile;
    property DataFileCount: Integer read GetDataFileCount;

    procedure DisplayApp(AppNode: TTreeNode);
    procedure DisplayUserGrp(UserGrpNode: TTreeNode);
    procedure DisplayFunc(FuncItem: TListItem);
    procedure DisplayUser(UserItem: TListItem);
    procedure DisplayDataFile(DataFileItem: TListItem);

    procedure NewApp(App: TApp; ParentNode: TTreeNode);
    procedure NewFunc(Func: TFunc; ListItems: TListItems);
    function NewDataFile(DataFile: TDataFile; ListItems: TListItems): TListItem;
    procedure NewUserGrp(UserGrp: TUserGrp; ParentNode: TTreeNode);
    procedure NewUser(User: TUser; ListItems: TListItems);

    procedure DeleteApp(AppNode: TTreeNode);
    procedure DeleteFunc(FuncItem: TListItem);
    procedure DeleteDataFile(DataFileItem: TListItem);
    procedure DeleteUserGrp(UserGrpNode: TTreeNode);
    procedure DeleteUser(UserItem: TListItem);

    procedure DeleteFuncFromApp(AppNode: TTreeNode; FuncItem: TListItem);
    procedure DeleteUserFromUserGrp(UserGrpNode: TTreeNode; UserItem: TListItem);
    procedure DeleteFuncFromUserGrp(UserGrpNode: TTreeNode; FuncItem: TListItem);
    procedure DeleteDataFileFromFunc(FuncItem, DataFileItem: TListItem);

    procedure AddFuncToApp(AppNode: TTreeNode; Func: TFunc;
      ListItems: TListItems);
    procedure AddUserToUserGrp(UserGrpNode: TTreeNode; User: TUser;
      ListItems: TListItems);
    procedure AddFuncToUserGrp(UserGrpNode: TTreeNode; FuncItem: TListItem);

    function HasUser(UserCode, Password: string): TUser;    
    function HasUserFunc(User: TUser; Func: TFunc): Boolean;

    procedure SaveToFile;
  end;

  TUserMgr = class(TObject)
  private
    FAllUsers: TStrings;
    function ExtractUser(Str: string): TUser;
  public
    constructor Create;
    destructor Destroy; override;
    function HasUser(UserCode, Password: string): TUser;
  end;

var
  UCMgr: TUCMgr;

implementation

procedure ParseStrings(Strs:string; Sep:Char; RStrs: TStrings);
var
  ii, jj : Integer ;
begin
  RStrs.Clear;
  if Strs = '' then Exit;
  jj := 1;
  for ii:= 1 to Length(Strs) do
  begin
    if Strs[ii] = Sep then
    begin
      RStrs.Add(Trim(Copy(Strs, jj, ii-jj)));
      jj := ii + 1;
    end;
  end;
  RStrs.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1)));
end;

function Max(Val1, Val2: Integer): Integer;
begin
  Result := Val1;
  if Val2 > Val1 then Result := Val2;
end;

function GetConfigFileName: string;
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'UC.ini');
  Result := IniFile.ReadString('System', 'UCPath', '') + '\UC.dat';
  IniFile.Free;
end;

{ TFunc }
constructor TFunc.Create;
begin
  FDataFiles := TStringList.Create;
  IconName := 'Default.ico';
end;

destructor TFunc.Destroy;
begin
  FDataFiles.Free;
end;

procedure TFunc.AddDataFile(DataFile: TDataFile);
begin
  if FDataFiles.IndexOfObject(DataFile) < 0 then
    FDataFiles.AddObject('', DataFile);
end;

procedure TFunc.DeleteDataFile(DataFile: TDataFile);
var
  Idx: Integer;
begin
  Idx := FDataFiles.IndexOfObject(DataFile);
  if Idx >= 0 then FDataFiles.Delete(Idx);
end;

function TFunc.GetDataFile(Idx: Integer): TDataFile;
begin
  Result := FDataFiles.Objects[Idx] As TDataFile;
end;

function TFunc.GetDataFileCount: Integer;
begin
  Result := FDataFiles.Count;
end;

function TFunc.HasDataFile(DataFile: TDataFile): Boolean;
begin
  Result := FDataFiles.IndexOfObject(DataFile) >= 0;
end;

procedure TFunc.DeleteAllDataFile;
begin
  FDataFiles.Clear;
end;

function TFunc.GetString: string;
begin
  Result := IntToStr(Id) + ',' +  Desc + ',' + ExeName + ',' + 
    Param + ',' + IconName;
end;

{ TApp }

constructor TApp.Create;
begin
  FFuncs := TStringList.Create;
end;

destructor TApp.Destroy;
begin
  FFuncs.Free;
end;

procedure TApp.AddFunc(Func: TFunc);
begin
  if FFuncs.IndexOfObject(Func) < 0 then
    FFuncs.AddObject('', Func);
end;

procedure TApp.DeleteFunc(Func: TFunc);
var
  Idx: Integer;
begin
  Idx := FFuncs.IndexOfObject(Func);
  if Idx >= 0 then FFuncs.Delete(Idx);
end;


function TApp.GetFunc(Idx: Integer): TFunc;
begin
  Result := FFuncs.Objects[Idx] As TFunc;
end;

function TApp.GetFuncCount: Integer;
begin
  Result := FFuncs.Count;
end;

function TApp.HasFunc(Func: TFunc): Boolean;
begin
  Result := FFuncs.IndexOfObject(Func) >= 0;
end;

function TApp.GetString: string;
begin
  Result := IntToStr(Id) + ',' + Desc + ',' + IntToStr(ParentId);
end;

{ TUserGrp }

constructor TUserGrp.Create;
begin
  FUsers := TStringList.Create;
  FFuncs := TStringList.Create;
end;

destructor TUserGrp.Destroy;
begin
  FUsers.Free;
  FFuncs.Free;
end;

procedure TUserGrp.AddUser(User: TUser);
begin
  if FUsers.IndexOfObject(User) < 0 then
    FUsers.AddObject('', User);
end;

procedure TUserGrp.DeleteUser(User: TUser);
var
  Idx: Integer;
begin
  Idx := FUsers.IndexOfObject(User);
  if Idx >= 0 then FUsers.Delete(Idx);
end;

function TUserGrp.GetUser(Idx: Integer): TUser;
begin
  Result := FUsers.Objects[Idx] As TUser;
end;

function TUserGrp.GetUserCount: Integer;
begin
  Result := FUsers.Count;
end;

procedure TUserGrp.AddFunc(Func: TFunc);
begin
  if FFuncs.IndexOfObject(Func) < 0 then
    FFuncs.AddObject('', Func);
end;

procedure TUserGrp.DeleteFunc(Func: TFunc);
var
  Idx: Integer;
begin
  Idx := FFuncs.IndexOfObject(Func);
  if Idx >= 0 then FFuncs.Delete(Idx);
end;

function TUserGrp.GetFunc(Idx: Integer): TFunc;
begin
  Result := FFuncs.Objects[Idx] As TFunc;
end;

function TUserGrp.GetFuncCount: Integer;
begin
  Result := FFuncs.Count;
end;

function TUserGrp.HasFunc(Func: TFunc): Boolean;
begin
  Result := FFuncs.IndexOfObject(Func) >= 0;
end;

function TUserGrp.HasUser(User: TUser): BOolean;
begin
  Result := FUsers.IndexOfObject(User) >= 0;
end;

procedure TUserGrp.DeleteAllFunc;
begin
  FFuncs.Clear;
end;

function TUserGrp.GetString: string;
begin
  Result := IntToStr(Id) + ',' + Desc + ',' + IntToStr(ParentId);
end;

{ TDataFile }

function TDataFile.GetString: string;
begin
  Result := IntToStr(Id) + ',' + FileName + ',' + Desc + ',';
  if AutoReg then Result := Result + '1'
  else Result := Result + '0';
end;

{ TUser }

function TUser.GetString: string;
begin
  Result := IntToStr(Id) + ',' + Code + ',' + Desc + ',' + Password;
end;

{ TUCMgr }

constructor TUCMgr.Create(AllApps, AllUserGrps: TTreeNodes);
begin
  FIniFile := TIniFile.Create(GetConfigFileName);

  FAllApps := AllApps;
  FAllUserGrps := AllUserGrps;
  FAllUsers := TStringList.Create;
  FAllFuncs := TStringList.Create;
  FAllDataFiles := TStringList.Create;

  GetAllDataFiles;
  GetAllFuncs;
  GetAllApps;
  GetAllUsers;
  GetAllUserGrps;

  FIniFile.Free;
end;

destructor TUCMgr.Destroy;
begin
  FAllUsers.Free;
  FAllFuncs.Free;
  FAllDataFiles.Free;
end;

procedure TUCMgr.DeleteApp(AppNode: TTreeNode);
begin
  TApp(AppNode.Data).Free;
  AppNode.Delete;
end;

procedure TUCMgr.DeleteDataFile(DataFileItem: TListItem);
var
  Idx: Integer;
begin
  Idx := FAllDataFiles.IndexOfObject(TObject(DataFileItem.Data));
  if Idx >= 0 then FAllDataFiles.Delete(Idx);
  for Idx:=0 to FuncCount-1 do
    Funcs[Idx].DeleteDataFile(TDataFile(DataFileItem.Data));
  TDataFile(DataFileItem.Data).Free;
  DataFileItem.Delete;
end;

procedure TUCMgr.DeleteFunc(FuncItem: TListItem);
var
  Idx: Integer;
begin
  Idx := FAllFuncs.IndexOfObject(TObject(FuncItem.Data));
  if Idx >= 0 then FAllFuncs.Delete(Idx);
  for Idx:=0 to AppCount-1 do
    Apps[Idx].DeleteFunc(TFunc(FuncItem.Data));
  for Idx:=0 to UserGrpCount-1 do
    UserGrps[Idx].DeleteFunc(TFunc(FuncItem.Data));
  TFunc(FuncItem.Data).Free;
  FuncItem.Delete;
end;

procedure TUCMgr.DeleteUser(UserItem: TListItem);
var
  Idx: Integer;
begin
  Idx := FAllUsers.IndexOfObject(TObject(UserItem.Data));
  if Idx >= 0 then FAllUsers.Delete(Idx);
  for Idx:=0 to UserGrpCount-1 do
    UserGrps[Idx].DeleteUser(TUser(UserItem.Data));
  TUser(UserItem.Data).Free;
  UserItem.Delete;
end;

procedure TUCMgr.DeleteUserGrp(UserGrpNode: TTreeNode);
begin
  TUserGrp(UserGrpNode.Data).Free;
  UserGrpNode.Delete;
end;

function TUCMgr.GetApp(Idx: Integer): TApp;
begin
  Result := TApp(FAllApps.Item[Idx].Data);
end;

function TUCMgr.GetAppCount: Integer;
begin
  Result := FAllApps.Count;
end;

function TUCMgr.GetDataFile(Idx: Integer): TDataFile;
begin
  Result := FAllDataFiles.Objects[Idx] As TDataFile;
end;

function TUCMgr.GetDataFileCount: Integer;
begin
  Result := FAllDataFiles.Count;
end;

function TUCMgr.GetFunc(Idx: Integer): TFunc;
begin
  Result := FAllFuncs.Objects[Idx] As TFunc;
end;

function TUCMgr.GetFuncCount: Integer;
begin
  Result := FAllFuncs.Count;
end;

function TUCMgr.GetMaxAppId: Integer;
begin
  Result := FMaxAppId;
end;

function TUCMgr.GetMaxDataFileId: Integer;
begin
  Result := FMaxDataFileId;
end;

function TUCMgr.GetMaxFuncId: Integer;
begin
  Result := FMaxFuncId;
end;

function TUCMgr.GetMaxUserGrpId: Integer;
begin
  Result := FMaxUserGrpId;
end;

function TUCMgr.GetMaxUserId: Integer;
begin
  Result := FMaxUserId;
end;

function TUCMgr.GetUser(Idx: Integer): TUser;
begin
  Result := FAllUsers.Objects[Idx] As TUser;
end;

function TUCMgr.GetUserCount: Integer;
begin
  Result := FAllUsers.Count;
end;

function TUCMgr.GetUserGrp(Idx: Integer): TUserGrp;
begin
  Result := TUserGrp(FAllUserGrps.Item[Idx].Data);
end;

function TUCMgr.GetUserGrpCount: Integer;
begin
  Result := FAllUserGrps.Count;
end;

procedure TUCMgr.NewApp(App: TApp; ParentNode: TTreeNode);
begin
  FMaxAppId := FMaxAppId + 1;
  App.Id := FMaxAppId;
  App.ParentId := TApp(ParentNode.Data).Id;
  DisplayApp(FAllApps.AddChildObject(ParentNode, '', App));
end;

function TUCMgr.NewDataFile(DataFile: TDataFile;
  ListItems: TListItems): TListItem;
var
  ListItem: TListItem;
begin
  FMaxDataFileId := FMaxDataFileId + 1;
  DataFile.Id := FMaxDataFileId;
  ListItem := ListItems.Add;
  ListItem.Data := DataFile;
  DisplayDataFile(ListItem);
  FAllDataFiles.AddObject('', DataFile);
  Result := ListItem;
end;

procedure TUCMgr.NewFunc(Func: TFunc; ListItems: TListItems);
var
  ListItem: TListItem;
begin
  FMaxFuncId := FMaxFuncId + 1;
  Func.Id := FMaxFuncId;
  ListItem := ListItems.Add;
  ListItem.Data := Func;
  DisplayFunc(ListItem);
  FAllFuncs.AddObject('', Func);
end;

procedure TUCMgr.NewUser(User: TUser; ListItems: TListItems);
var
  ListItem: TListItem;
begin
  FMaxUserId := FMaxUserId + 1;
  User.Id := FMaxUserId;
  ListItem := ListItems.Add;
  ListItem.Data := User;
  DisplayUser(ListItem);
  FAllUsers.AddObject('', User);  
end;

procedure TUCMgr.NewUserGrp(UserGrp: TUserGrp; ParentNode: TTreeNode);
var
  ii: Integer;
  ParentUserGrp: TUserGrp;
begin
  FMaxUserGrpId := FMaxUserGrpId + 1;
  UserGrp.Id := FMaxUserGrpId;
  UserGrp.ParentId := TUserGrp(ParentNode.Data).Id;
  DisplayUserGrp(FAllUserGrps.AddChildObject(ParentNode, '', UserGrp));
  {add all func of parent usergrp to the new usergrp}
  ParentUserGrp := TUserGrp(ParentNode.Data);
  for ii:=0 to ParentUserGrp.FuncCount-1 do
    UserGrp.AddFunc(ParentUserGrp.Funcs[ii]);
end;

function TUCMgr.ExtractApp(Str: string): TApp;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TApp.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.Desc := Attrs[1];
  Result.ParentId := StrToInt(Attrs[2]);
  Attrs.Free;
end;

function TUCMgr.ExtractFunc(Str: string): TFunc;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TFunc.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.Desc := Attrs[1];
  Result.ExeName := Attrs[2];
  Result.Param := Attrs[3];
  Result.IconName := Attrs[4];
  Attrs.Free;
end;

function TUCMgr.ExtractDataFile(Str: string): TDataFile;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TDataFile.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.FileName := Attrs[1];
  Result.Desc := Attrs[2];
  Result.AutoReg := StrToInt(Attrs[3]) = 1;
  Attrs.Free;
end;

function TUCMgr.ExtractUser(Str: string): TUser;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TUser.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.Code := Attrs[1];
  Result.Desc := Attrs[2];
  Result.Password := Attrs[3];
  Attrs.Free;
end;

function TUCMgr.ExtractUserGrp(Str: string): TUserGrp;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TUserGrp.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.Desc := Attrs[1];
  Result.ParentId := StrToInt(Attrs[2]);
  Attrs.Free;
end;

procedure TUCMgr.GetAllDataFiles;
var
  Cnt: Integer;
  Str: string;
  DataFile: TDataFile;
begin
  Cnt := 0;
  while True do
  begin
    Str := FIniFile.ReadString('DataFile', 'DataFile'+IntToStr(Cnt), '');
    if Str = '' then Break;
    DataFile := ExtractDataFile(Str);
    FMaxDataFileId := Max(DataFile.Id, FMaxDataFileId);
    FAllDataFiles.AddObject('', DataFile);
    Cnt := Cnt + 1;
  end;
end;

procedure TUCMgr.GetDataFilesOfFunc(Func: TFunc; FuncIdx: Integer);
var
  Str: string;
  DataFileIdxs: TStrings;
  ii, Idx: Integer;
begin
  Str := FIniFile.ReadString('DataFile Of Func', 'Func'+IntToStr(FuncIdx), '');
  if Str = '' then Exit;
  DataFileIdxs := TStringList.Create;
  ParseStrings(Str, ',', DataFileIdxs);
  for ii:=0 to DataFileIdxs.Count-1 do
  begin
    Idx := StrToInt(DataFileIdxs[ii]);
    if Idx >= DataFileCount then Continue;
    Func.AddDataFile(DataFiles[Idx]);
  end;
  DataFileIdxs.Free;
end;

procedure TUCMgr.GetAllFuncs;
var
  Cnt: Integer;
  Str: string;
  Func: TFunc;
begin
  Cnt := 0;
  while True do
  begin
    Str := FIniFile.ReadString('Func', 'Func'+IntToStr(Cnt), '');
    if Str = '' then Break;
    Func := ExtractFunc(Str);
    FMaxFuncId := Max(Func.Id, FMaxFuncId);
    GetDataFilesOfFunc(Func, Cnt);
    FAllFuncs.AddObject('', Func);
    Cnt := Cnt + 1;
  end;
end;

procedure TUCMgr.GetFuncsOfApp(App: TApp; AppIdx: Integer);
var
  Str: string;
  FuncIdxs: TStrings;
  ii, Idx: Integer;
begin
  Str := FIniFile.ReadString('Func Of App', 'App'+IntToStr(AppIdx), '');
  if Str = '' then Exit;
  FuncIdxs := TStringList.Create;
  ParseStrings(Str, ',', FuncIdxs);
  for ii:=0 to FuncIdxs.Count-1 do
  begin
    Idx := StrToInt(FuncIdxs[ii]);
    if Idx >= FuncCount then Continue;
    App.AddFunc(Funcs[Idx]);
  end;
  FuncIdxs.Free;
end;

function TUCMgr.GetParentAppNode(ParentId: Integer): TTreeNode;
var
  ii: Integer;
begin
  for ii:=0 to FAllApps.Count-1 do
  begin
    Result := FAllApps[ii];
    if TApp(FAllApps[ii].Data).Id = ParentId then Exit;
  end;
  Result := nil;
end;

procedure TUCMgr.GetAllApps;
var
  ii, Cnt: Integer;
  Str: string;
  App: TApp;
begin

  App := TApp.Create;
  FAllApps.BeginUpdate;
  FAllApps.AddObject(nil, '所有系統', App);
  for ii:=0 to FuncCount-1 do App.AddFunc(Funcs[ii]);

  FMaxAppId := 0;
  Cnt := 0;
  while True do
  begin
    Str := FIniFile.ReadString('App', 'App'+IntToStr(Cnt), '');
    if Str = '' then Break;
    App := ExtractApp(Str);
    FMaxAppId := Max(App.Id, FMaxAppId);
    GetFuncsOfApp(App, Cnt);
    FAllApps.AddChildObject(GetParentAppNode(App.ParentId), App.Desc, App);
    Cnt := Cnt + 1;
  end;
  FAllApps.Owner.FullExpand;
  FAllApps.EndUpdate;  
end;

function TUCMgr.GetParentUserGrpNode(ParentId: Integer): TTreeNode;
var
  ii: Integer;
begin
  for ii:=0 to FAllUserGrps.Count-1 do
  begin
    Result := FAllUserGrps[ii];
    if TUserGrp(FAllUserGrps[ii].Data).Id = ParentId then Exit;
  end;
  Result := nil;
end;

procedure TUCMgr.GetFuncsOfUserGrp(UserGrp: TUserGrp; UserGrpIdx: Integer);
var
  Str: string;
  FuncIdxs: TStrings;
  ii, Idx: Integer;
begin
  Str := FIniFile.ReadString('Func Of UserGrp', 'UserGrp'+IntToStr(UserGrpIdx), '');
  if Str = '' then Exit;
  FuncIdxs := TStringList.Create;
  ParseStrings(Str, ',', FuncIdxs);
  for ii:=0 to FuncIdxs.Count-1 do
  begin
    Idx := StrToInt(FuncIdxs[ii]);
    if Idx >= FuncCount then Continue;
    UserGrp.AddFunc(Funcs[Idx]);
  end;
  FuncIdxs.Free;
end;

procedure TUCMgr.GetUsersOfUserGrp(UserGrp: TUserGrp; UserGrpIdx: Integer);
var
  Str: string;
  UserIdxs: TStrings;
  ii, Idx: Integer;
begin
  Str := FIniFile.ReadString('User Of UserGrp', 'UserGrp'+IntToStr(UserGrpIdx), '');
  if Str = '' then Exit;
  UserIdxs := TStringList.Create;
  ParseStrings(Str, ',', UserIdxs);
  for ii:=0 to UserIdxs.Count-1 do
  begin
    Idx := StrToInt(UserIdxs[ii]);
    if Idx >= UserCount then Continue;
    UserGrp.AddUser(Users[Idx]);
  end;
  UserIdxs.Free;
end;

procedure TUCMgr.GetAllUserGrps;
var
  ii, Cnt: Integer;
  Str: string;
  UserGrp: TUserGrp;
begin

  UserGrp := TUserGrp.Create;
  FAllUserGrps.BeginUpdate;
  FAllUserGrps.AddObject(nil, '所有群組', UserGrp);
  for ii:=0 to UserCount-1 do UserGrp.AddUser(Users[ii]);

  FMaxUserGrpId := 0;
  Cnt := 0;
  while True do
  begin
    Str := FIniFile.ReadString('UserGrp', 'UserGrp'+IntToStr(Cnt), '');
    if Str = '' then Break;
    UserGrp := ExtractUserGrp(Str);
    FMaxUserGrpId := Max(UserGrp.Id, FMaxUserGrpId);
    GetUsersOfUserGrp(UserGrp, Cnt);
    GetFuncsOfUserGrp(UserGrp, Cnt);
    FAllUserGrps.AddChildObject(GetParentUserGrpNode(UserGrp.ParentId),
      UserGrp.Desc, UserGrp);
    Cnt := Cnt + 1;
  end;
  FAllUserGrps.Owner.FullExpand;
  FAllUserGrps.EndUpdate;  
end;

procedure TUCMgr.GetAllUsers;
var
  Cnt: Integer;
  Str: string;
  User: TUser;
begin
  Cnt := 0;
  while True do
  begin
    Str := FIniFile.ReadString('User', 'User'+IntToStr(Cnt), '');
    if Str = '' then Break;
    User := ExtractUser(Str);
    FMaxUserId := Max(User.Id, FMaxUserId);
    FAllUsers.AddObject('', User);
    Cnt := Cnt + 1;
  end;
end;

procedure TUCMgr.DeleteDataFileFromFunc(FuncItem, DataFileItem: TListItem);
begin
  TFunc(FuncItem.Data).DeleteDataFile(TDataFile(DataFileItem.Data));
  if DataFileItem.ListView.Checkboxes then DataFileItem.Checked := False
  else DataFileItem.Delete;
end;

procedure TUCMgr.DeleteFuncFromApp(AppNode: TTreeNode; FuncItem: TListItem);
begin
  TApp(AppNode.Data).DeleteFunc(TFunc(FuncItem.Data));
  if FuncItem.ListView.Checkboxes then FuncItem.Checked := False
  else FuncItem.Delete;
end;

procedure TUCMgr.DeleteFuncFromUserGrp(UserGrpNode: TTreeNode;
  FuncItem: TListItem);
begin
  TUserGrp(UserGrpNode.Data).DeleteFunc(TFunc(FuncItem.Data));
  if FuncItem.ListView.Checkboxes then FuncItem.Checked := False
  else FuncItem.Delete;
end;

procedure TUCMgr.DeleteUserFromUserGrp(UserGrpNode: TTreeNode;
  UserItem: TListItem);
begin
  TUserGrp(UserGrpNode.Data).DeleteUser(TUser(UserItem.Data));
  if UserItem.ListView.Checkboxes then UserItem.Checked := False
  else UserItem.Delete;
end;

procedure TUCMgr.DisplayApp(AppNode: TTreeNode);
var
  App: TApp;
begin
  App := TApp(AppNode.Data);
  AppNode.Text := App.Desc;
end;

procedure TUCMgr.DisplayDataFile(DataFileItem: TListItem);
var
  DataFile: TDataFile;
begin
  DataFile := TDataFile(DataFileItem.Data);
  DataFileItem.Caption := DataFile.FileName;
  if TListView(DataFileItem.ListView).ViewStyle <> vsReport then Exit;
  if DataFileItem.SubItems.Count = 0 then
  begin
    DataFileItem.SubItems.Add('');
    DataFileItem.SubItems.Add('');
  end;
  DataFileItem.SubItems[0] := DataFile.Desc;
  if DataFile.AutoReg then DataFileItem.SubItems[1] := 'Yes'
  else DataFileItem.SubItems[1] := 'No';
end;

procedure TUCMgr.DisplayFunc(FuncItem: TListItem);
var
  Func: TFunc;
begin
  Func := TFunc(FuncItem.Data);
  FuncItem.Caption := Func.Desc;
  if TListView(FuncItem.ListView).ViewStyle <> vsReport then Exit;
  if FuncItem.SubItems.Count = 0 then
  begin
    FuncItem.SubItems.Add('');
    FuncItem.SubItems.Add('');
    FuncItem.SubItems.Add('');
  end;
  FuncItem.SubItems[0] := Func.ExeName;
  FuncItem.SubItems[1] := Func.Param;
  FuncItem.SubItems[2] := Func.IconName;  
end;

procedure TUCMgr.DisplayUser(UserItem: TListItem);
var
  User: TUser;
begin
  User := TUser(UserItem.Data);
  UserItem.Caption := User.Code;
  if TListView(UserItem.ListView).ViewStyle <> vsReport then Exit;
  if UserItem.SubItems.Count = 0 then UserItem.SubItems.Add('');
  UserItem.SubItems[0] := User.Desc;
end;

procedure TUCMgr.DisplayUserGrp(UserGrpNode: TTreeNode);
var
  UserGrp: TUserGrp;
begin
  UserGrp := TUserGrp(UserGrpNode.Data);
  UserGrpNode.Text := UserGrp.Desc;
end;

procedure TUCMgr.AddFuncToApp(AppNode: TTreeNode; Func: TFunc;
  ListItems: TListItems);
var
  ListItem: TListItem;
begin
  ListItem := ListItems.Add;
  ListItem.Data := Func;
  TApp(AppNode.Data).AddFunc(Func);
  DisplayFunc(ListItem);
end;

procedure TUCMgr.AddFuncToUserGrp(UserGrpNode: TTreeNode; FuncItem: TListItem);
var
  ii: Integer;
  Func: TFunc;
begin
  {Add the func to all child of the user group}
  for ii:=0 to UserGrpNode.Count-1 do
    AddFuncToUserGrp(UserGrpNode.Item[ii], FuncItem);
  {Add the func to the specified user group}
  Func := TFunc(FuncItem.Data);
  TUserGrp(UserGrpNode.Data).AddFunc(Func);
end;

procedure TUCMgr.AddUserToUserGrp(UserGrpNode: TTreeNode; User: TUser;
  ListItems: TListItems);
var
  ListItem: TListItem;
begin
  ListItem := ListItems.Add;
  ListItem.Data := User;
  TUserGrp(UserGrpNode.Data).AddUser(User);
  DisplayUser(ListItem);
end;

function TUCMgr.HasUserFunc(User: TUser; Func: TFunc): Boolean;
var
  ii: Integer;
begin
  Result := False;
  for ii:=0 to UserGrpCount-1 do
    if UserGrps[ii].HasUser(User) and UserGrps[ii].HasFunc(Func) then
    begin
      Result := True;
      Exit;
    end;
end;

procedure TUCMgr.ClearFile;
var
  ii: Integer;
  Sects: TStrings;
begin
  Sects := TStringList.Create;
  FIniFile.ReadSections(Sects);
  for ii:=0 to Sects.Count-1 do
    FIniFile.EraseSection(Sects[ii]);
end;

procedure TUCMgr.SaveToFile;
var
  ii, jj: Integer;
  Str: string;
begin
  FIniFile := TIniFile.Create(GetConfigFileName);
  ClearFile;

  {save all app}
  for ii:=1 to AppCount-1 do
    FIniFile.WriteString('App', 'App'+IntToStr(ii-1), Apps[ii].GetString);

  {save all func}
  for ii:=0 to FuncCount-1 do
    FIniFile.WriteString('Func', 'Func'+IntToStr(ii), Funcs[ii].GetString);

  {save all datafile}
  for ii:=0 to DataFileCount-1 do
    FIniFile.WriteString('DataFile', 'DataFile'+IntToStr(ii),
      DataFiles[ii].GetString);

  {save all usergrp}
  for ii:=1 to UserGrpCount-1 do
    FIniFile.WriteString('UserGrp', 'UserGrp'+IntToStr(ii-1),
      UserGrps[ii].GetString);

  {save all user}
  for ii:=0 to UserCount-1 do
    FIniFile.WriteString('User', 'User'+IntToStr(ii), Users[ii].GetString);

  {save func of app}
  for ii:=1 to AppCount-1 do
  begin
    Str := '';
    for jj:=0 to Apps[ii].FuncCount-1 do
    begin
      if Str <> '' then Str := Str + ',';
      Str := Str + IntToStr(FAllFuncs.IndexOfObject(Apps[ii].Funcs[jj]));
    end;
    if Str <> '' then
      FIniFile.WriteString('Func Of App', 'App'+IntToStr(ii-1), Str);
  end;

  {save datafile of func}
  for ii:=0 to FuncCount-1 do
  begin
    Str := '';
    for jj:=0 to Funcs[ii].DataFileCount-1 do
    begin
      if Str <> '' then Str := Str + ',';
      Str := Str + IntToStr(FAllDataFiles.IndexOfObject(Funcs[ii].DataFiles[jj]));
    end;
    if Str <> '' then
      FIniFile.WriteString('DataFile Of Func', 'Func'+IntToStr(ii), Str);
  end;

  {save user of usergrp}
  for ii:=1 to UserGrpCount-1 do
  begin
    Str := '';
    for jj:=0 to UserGrps[ii].UserCount-1 do
    begin
      if Str <> '' then Str := Str + ',';
      Str := Str + IntToStr(FAllUsers.IndexOfObject(UserGrps[ii].Users[jj]));
    end;
    if Str <> '' then
      FIniFile.WriteString('User Of UserGrp', 'UserGrp'+IntToStr(ii-1), Str);
  end;

  {save func of usergrp}
  for ii:=1 to UserGrpCount-1 do
  begin
    Str := '';
    for jj:=0 to UserGrps[ii].FuncCount-1 do
    begin
      if Str <> '' then Str := Str + ',';
      Str := Str + IntToStr(FAllFuncs.IndexOfObject(UserGrps[ii].Funcs[jj]));
    end;
    if Str <> '' then
      FIniFile.WriteString('Func Of UserGrp', 'UserGrp'+IntToStr(ii-1), Str);
  end;

  FIniFile.Free;
end;

function TUCMgr.HasUser(UserCode, Password: string): TUser;
var
  ii: Integer;
  User: TUser;
begin
  for ii:=0 to FAllUsers.Count-1 do
  begin
    User := TUser(FAllUsers.Objects[ii]);
    Result := User;
    if (UpperCase(UserCode) = UpperCase(UserCode)) and
      (UpperCase(User.Password) = UpperCase(Password)) then Exit;
  end;
  Result := nil;
end;

{ TUserMgr }

function TUserMgr.ExtractUser(Str: string): TUser;
var
  Attrs: TStrings;
begin
  Attrs := TStringList.Create;
  Result := TUser.Create;
  ParseStrings(Str, ',' , Attrs);
  Result.Id := StrToInt(Attrs[0]);
  Result.Code := Attrs[1];
  Result.Desc := Attrs[2];
  Result.Password := Attrs[3];
  Attrs.Free;
end;

constructor TUserMgr.Create;
var
  IniFile: TIniFile;
  Cnt: Integer;
  Str: string;
  User: TUser;
begin
  FAllUsers := TStringList.Create;
  IniFile := TIniFile.Create(GetConfigFileName);
  Cnt := 0;
  while True do
  begin
    Str := IniFile.ReadString('User', 'User'+IntToStr(Cnt), '');
    if Str = '' then Break;
    User := ExtractUser(Str);
    FAllUsers.AddObject('', User);
    Cnt := Cnt + 1;
  end;
  IniFile.Free;
end;

destructor TUserMgr.Destroy;
var
  ii: Integer;
begin
  for ii:=0 to FAllUsers.Count-1 do
    FAllUsers.Objects[ii].Free;
  FAllUsers.Free;
end;

function TUserMgr.HasUser(UserCode, Password: string): TUser;
var
  ii: Integer;
  User: TUser;
begin
  for ii:=0 to FAllUsers.Count-1 do
  begin
    User := TUser(FAllUsers.Objects[ii]);
    Result := User;
    if (UpperCase(UserCode) = UpperCase(UserCode)) and
      (UpperCase(User.Password) = UpperCase(Password)) then Exit;
  end;
  Result := nil;
end;

end.


