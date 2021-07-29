unit prjActMainSvr_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : $Revision:   1.130  $
// File generated on 2002/1/25 ¤W¤È 10:44:27 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Account\Howard\ApServer\ActMainSvr\prjActMainSvr.tlb (1)
// LIBID: {B0527FA5-F54A-48DD-9576-52D45CB41262}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v1.0 Midas, (D:\WINNT\System32\midas.dll)
//   (2) v2.0 stdole, (D:\WINNT\SYSTEM32\stdole2.tlb)
//   (3) v4.0 StdVCL, (D:\WINNT\System32\stdvcl40.dll)
// Errors:
//   Error creating palette bitmap of (TActMainSvr) : No Server registered for this CoClass
//   Error creating palette bitmap of (Tabc) : No Server registered for this CoClass
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

interface

uses ActiveX, Classes, Graphics, Midas, OleServer, StdVCL, Variants, 
Windows;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  prjActMainSvrMajorVersion = 1;
  prjActMainSvrMinorVersion = 0;

  LIBID_prjActMainSvr: TGUID = '{B0527FA5-F54A-48DD-9576-52D45CB41262}';

  IID_IActMainSvr: TGUID = '{947C8F34-E50E-467B-8524-E330BBE7ACD1}';
  CLASS_ActMainSvr: TGUID = '{C70E369C-7449-42C5-BA46-259AB166D5C6}';
  IID_Iabc: TGUID = '{A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}';
  CLASS_abc: TGUID = '{68E3694B-9137-4E54-8F17-84DD53DF05AA}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IActMainSvr = interface;
  IActMainSvrDisp = dispinterface;
  Iabc = interface;
  IabcDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  ActMainSvr = IActMainSvr;
  abc = Iabc;


// *********************************************************************//
// Interface: IActMainSvr
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {947C8F34-E50E-467B-8524-E330BBE7ACD1}
// *********************************************************************//
  IActMainSvr = interface(IAppServer)
    ['{947C8F34-E50E-467B-8524-E330BBE7ACD1}']
    procedure connect2db(var bI_Connected: OleVariant); safecall;
    procedure checkUserInfo(sI_UserID: OleVariant; sI_Passwd: OleVariant; 
                            var sI_UserName: OleVariant; var sI_GroupID: OleVariant; 
                            var sI_CompCode: OleVariant; var sI_CompName: OleVariant; 
                            var bI_Correct: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  IActMainSvrDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {947C8F34-E50E-467B-8524-E330BBE7ACD1}
// *********************************************************************//
  IActMainSvrDisp = dispinterface
    ['{947C8F34-E50E-467B-8524-E330BBE7ACD1}']
    procedure connect2db(var bI_Connected: OleVariant); dispid 1;
    procedure checkUserInfo(sI_UserID: OleVariant; sI_Passwd: OleVariant; 
                            var sI_UserName: OleVariant; var sI_GroupID: OleVariant; 
                            var sI_CompCode: OleVariant; var sI_CompName: OleVariant; 
                            var bI_Correct: OleVariant); dispid 2;
    function  AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                              MaxErrors: Integer; out ErrorCount: Integer; var OwnerData: OleVariant): OleVariant; dispid 20000000;
    function  AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                            Options: Integer; const CommandText: WideString; 
                            var Params: OleVariant; var OwnerData: OleVariant): OleVariant; dispid 20000001;
    function  AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant; dispid 20000002;
    function  AS_GetProviderNames: OleVariant; dispid 20000003;
    function  AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant; dispid 20000004;
    function  AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                            var OwnerData: OleVariant): OleVariant; dispid 20000005;
    procedure AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                         var Params: OleVariant; var OwnerData: OleVariant); dispid 20000006;
  end;

// *********************************************************************//
// Interface: Iabc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}
// *********************************************************************//
  Iabc = interface(IAppServer)
    ['{A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}']
  end;

// *********************************************************************//
// DispIntf:  IabcDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}
// *********************************************************************//
  IabcDisp = dispinterface
    ['{A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}']
    function  AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                              MaxErrors: Integer; out ErrorCount: Integer; var OwnerData: OleVariant): OleVariant; dispid 20000000;
    function  AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                            Options: Integer; const CommandText: WideString; 
                            var Params: OleVariant; var OwnerData: OleVariant): OleVariant; dispid 20000001;
    function  AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant; dispid 20000002;
    function  AS_GetProviderNames: OleVariant; dispid 20000003;
    function  AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant; dispid 20000004;
    function  AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                            var OwnerData: OleVariant): OleVariant; dispid 20000005;
    procedure AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                         var Params: OleVariant; var OwnerData: OleVariant); dispid 20000006;
  end;

// *********************************************************************//
// The Class CoActMainSvr provides a Create and CreateRemote method to          
// create instances of the default interface IActMainSvr exposed by              
// the CoClass ActMainSvr. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoActMainSvr = class
    class function Create: IActMainSvr;
    class function CreateRemote(const MachineName: string): IActMainSvr;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TActMainSvr
// Help String      : ActMainSvr Object
// Default Interface: IActMainSvr
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TActMainSvrProperties= class;
{$ENDIF}
  TActMainSvr = class(TOleServer)
  private
    FIntf:        IActMainSvr;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TActMainSvrProperties;
    function      GetServerProperties: TActMainSvrProperties;
{$ENDIF}
    function      GetDefaultInterface: IActMainSvr;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IActMainSvr);
    procedure Disconnect; override;
    function  AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                              MaxErrors: Integer; out ErrorCount: Integer; var OwnerData: OleVariant): OleVariant;
    function  AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                            Options: Integer; const CommandText: WideString; 
                            var Params: OleVariant; var OwnerData: OleVariant): OleVariant;
    function  AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant;
    function  AS_GetProviderNames: OleVariant;
    function  AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant;
    function  AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                            var OwnerData: OleVariant): OleVariant;
    procedure AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                         var Params: OleVariant; var OwnerData: OleVariant);
    procedure connect2db(var bI_Connected: OleVariant);
    procedure checkUserInfo(sI_UserID: OleVariant; sI_Passwd: OleVariant; 
                            var sI_UserName: OleVariant; var sI_GroupID: OleVariant; 
                            var sI_CompCode: OleVariant; var sI_CompName: OleVariant; 
                            var bI_Correct: OleVariant);
    property  DefaultInterface: IActMainSvr read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TActMainSvrProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TActMainSvr
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TActMainSvrProperties = class(TPersistent)
  private
    FServer:    TActMainSvr;
    function    GetDefaultInterface: IActMainSvr;
    constructor Create(AServer: TActMainSvr);
  protected
  public
    property DefaultInterface: IActMainSvr read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class Coabc provides a Create and CreateRemote method to          
// create instances of the default interface Iabc exposed by              
// the CoClass abc. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  Coabc = class
    class function Create: Iabc;
    class function CreateRemote(const MachineName: string): Iabc;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : Tabc
// Help String      : abc Object
// Default Interface: Iabc
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TabcProperties= class;
{$ENDIF}
  Tabc = class(TOleServer)
  private
    FIntf:        Iabc;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TabcProperties;
    function      GetServerProperties: TabcProperties;
{$ENDIF}
    function      GetDefaultInterface: Iabc;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: Iabc);
    procedure Disconnect; override;
    function  AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                              MaxErrors: Integer; out ErrorCount: Integer; var OwnerData: OleVariant): OleVariant;
    function  AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                            Options: Integer; const CommandText: WideString; 
                            var Params: OleVariant; var OwnerData: OleVariant): OleVariant;
    function  AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant;
    function  AS_GetProviderNames: OleVariant;
    function  AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant;
    function  AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                            var OwnerData: OleVariant): OleVariant;
    procedure AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                         var Params: OleVariant; var OwnerData: OleVariant);
    property  DefaultInterface: Iabc read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TabcProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : Tabc
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TabcProperties = class(TPersistent)
  private
    FServer:    Tabc;
    function    GetDefaultInterface: Iabc;
    constructor Create(AServer: Tabc);
  protected
  public
    property DefaultInterface: Iabc read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

implementation

uses ComObj;

class function CoActMainSvr.Create: IActMainSvr;
begin
  Result := CreateComObject(CLASS_ActMainSvr) as IActMainSvr;
end;

class function CoActMainSvr.CreateRemote(const MachineName: string): IActMainSvr;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ActMainSvr) as IActMainSvr;
end;

procedure TActMainSvr.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{C70E369C-7449-42C5-BA46-259AB166D5C6}';
    IntfIID:   '{947C8F34-E50E-467B-8524-E330BBE7ACD1}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TActMainSvr.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IActMainSvr;
  end;
end;

procedure TActMainSvr.ConnectTo(svrIntf: IActMainSvr);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TActMainSvr.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TActMainSvr.GetDefaultInterface: IActMainSvr;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TActMainSvr.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TActMainSvrProperties.Create(Self);
{$ENDIF}
end;

destructor TActMainSvr.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TActMainSvr.GetServerProperties: TActMainSvrProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function  TActMainSvr.AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                                      MaxErrors: Integer; out ErrorCount: Integer; 
                                      var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_ApplyUpdates(ProviderName, Delta, MaxErrors, ErrorCount, OwnerData);
end;

function  TActMainSvr.AS_GetRecords(const ProviderName: WideString; Count: Integer; 
                                    out RecsOut: Integer; Options: Integer; 
                                    const CommandText: WideString; var Params: OleVariant; 
                                    var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_GetRecords(ProviderName, Count, RecsOut, Options, CommandText, 
                                           Params, OwnerData);
end;

function  TActMainSvr.AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_DataRequest(ProviderName, Data);
end;

function  TActMainSvr.AS_GetProviderNames: OleVariant;
begin
  Result := DefaultInterface.AS_GetProviderNames;
end;

function  TActMainSvr.AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_GetParams(ProviderName, OwnerData);
end;

function  TActMainSvr.AS_RowRequest(const ProviderName: WideString; Row: OleVariant; 
                                    RequestType: Integer; var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_RowRequest(ProviderName, Row, RequestType, OwnerData);
end;

procedure TActMainSvr.AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                                 var Params: OleVariant; var OwnerData: OleVariant);
begin
  DefaultInterface.AS_Execute(ProviderName, CommandText, Params, OwnerData);
end;

procedure TActMainSvr.connect2db(var bI_Connected: OleVariant);
begin
  DefaultInterface.connect2db(bI_Connected);
end;

procedure TActMainSvr.checkUserInfo(sI_UserID: OleVariant; sI_Passwd: OleVariant; 
                                    var sI_UserName: OleVariant; var sI_GroupID: OleVariant; 
                                    var sI_CompCode: OleVariant; var sI_CompName: OleVariant; 
                                    var bI_Correct: OleVariant);
begin
  DefaultInterface.checkUserInfo(sI_UserID, sI_Passwd, sI_UserName, sI_GroupID, sI_CompCode, 
                                 sI_CompName, bI_Correct);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TActMainSvrProperties.Create(AServer: TActMainSvr);
begin
  inherited Create;
  FServer := AServer;
end;

function TActMainSvrProperties.GetDefaultInterface: IActMainSvr;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

class function Coabc.Create: Iabc;
begin
  Result := CreateComObject(CLASS_abc) as Iabc;
end;

class function Coabc.CreateRemote(const MachineName: string): Iabc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_abc) as Iabc;
end;

procedure Tabc.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{68E3694B-9137-4E54-8F17-84DD53DF05AA}';
    IntfIID:   '{A7FADC08-7F31-4BD4-BDA3-B89D93211EAD}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure Tabc.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as Iabc;
  end;
end;

procedure Tabc.ConnectTo(svrIntf: Iabc);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure Tabc.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function Tabc.GetDefaultInterface: Iabc;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor Tabc.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TabcProperties.Create(Self);
{$ENDIF}
end;

destructor Tabc.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function Tabc.GetServerProperties: TabcProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function  Tabc.AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; 
                               MaxErrors: Integer; out ErrorCount: Integer; 
                               var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_ApplyUpdates(ProviderName, Delta, MaxErrors, ErrorCount, OwnerData);
end;

function  Tabc.AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                             Options: Integer; const CommandText: WideString; 
                             var Params: OleVariant; var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_GetRecords(ProviderName, Count, RecsOut, Options, CommandText, 
                                           Params, OwnerData);
end;

function  Tabc.AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_DataRequest(ProviderName, Data);
end;

function  Tabc.AS_GetProviderNames: OleVariant;
begin
  Result := DefaultInterface.AS_GetProviderNames;
end;

function  Tabc.AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_GetParams(ProviderName, OwnerData);
end;

function  Tabc.AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                             var OwnerData: OleVariant): OleVariant;
begin
  Result := DefaultInterface.AS_RowRequest(ProviderName, Row, RequestType, OwnerData);
end;

procedure Tabc.AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                          var Params: OleVariant; var OwnerData: OleVariant);
begin
  DefaultInterface.AS_Execute(ProviderName, CommandText, Params, OwnerData);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TabcProperties.Create(AServer: Tabc);
begin
  inherited Create;
  FServer := AServer;
end;

function TabcProperties.GetDefaultInterface: Iabc;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TActMainSvr, Tabc]);
end;

end.
