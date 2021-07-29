unit Crypt32DLL_TLB;

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

// PASTLWTR : 1.2
// File generated on 2007/3/29 ¤U¤È 03:51:01 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\CommoUnit\Crypt32DLL.tlb (1)
// LIBID: {A035928C-A622-4D7A-8BF3-0996C4494627}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\STDOLE2.TLB)
// Errors:
//   Error creating palette bitmap of (TCrypt32) : Server C:\CSMISV40\common\crypt32dll.dll contains no icons
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
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  Crypt32DLLMajorVersion = 7;
  Crypt32DLLMinorVersion = 0;

  LIBID_Crypt32DLL: TGUID = '{A035928C-A622-4D7A-8BF3-0996C4494627}';

  IID__Crypt32: TGUID = '{7F0AC8C6-D806-471A-A984-B06B223133B8}';
  CLASS_Crypt32: TGUID = '{A47EE28E-374D-4952-9361-8F728AFA3800}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _Crypt32 = interface;
  _Crypt32Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Crypt32 = _Crypt32;


// *********************************************************************//
// Interface: _Crypt32
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {7F0AC8C6-D806-471A-A984-B06B223133B8}
// *********************************************************************//
  _Crypt32 = interface(IDispatch)
    ['{7F0AC8C6-D806-471A-A984-B06B223133B8}']
    function AValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; safecall;
    function BValue(const a: WideString; var b: OleVariant; const c: WideString): WideString; safecall;
    function CValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; safecall;
    function DValue(const a: WideString; var b: OleVariant; const c: WideString): WideString; safecall;
    function EValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; safecall;
    function FValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; safecall;
    function GValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  _Crypt32Disp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {7F0AC8C6-D806-471A-A984-B06B223133B8}
// *********************************************************************//
  _Crypt32Disp = dispinterface
    ['{7F0AC8C6-D806-471A-A984-B06B223133B8}']
    function AValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; dispid 1610809344;
    function BValue(const a: WideString; var b: OleVariant; const c: WideString): WideString; dispid 1610809345;
    function CValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; dispid 1610809346;
    function DValue(const a: WideString; var b: OleVariant; const c: WideString): WideString; dispid 1610809347;
    function EValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; dispid 1610809348;
    function FValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; dispid 1610809349;
    function GValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString; dispid 1610809350;
  end;

// *********************************************************************//
// The Class CoCrypt32 provides a Create and CreateRemote method to          
// create instances of the default interface _Crypt32 exposed by              
// the CoClass Crypt32. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCrypt32 = class
    class function Create: _Crypt32;
    class function CreateRemote(const MachineName: string): _Crypt32;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCrypt32
// Help String      : 
// Default Interface: _Crypt32
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TCrypt32Properties= class;
{$ENDIF}
  TCrypt32 = class(TOleServer)
  private
    FIntf:        _Crypt32;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TCrypt32Properties;
    function      GetServerProperties: TCrypt32Properties;
{$ENDIF}
    function      GetDefaultInterface: _Crypt32;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _Crypt32);
    procedure Disconnect; override;
    function AValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
    function BValue(const a: WideString; var b: OleVariant; const c: WideString): WideString;
    function CValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
    function DValue(const a: WideString; var b: OleVariant; const c: WideString): WideString;
    function EValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
    function FValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
    function GValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
    property DefaultInterface: _Crypt32 read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TCrypt32Properties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TCrypt32
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TCrypt32Properties = class(TPersistent)
  private
    FServer:    TCrypt32;
    function    GetDefaultInterface: _Crypt32;
    constructor Create(AServer: TCrypt32);
  protected
  public
    property DefaultInterface: _Crypt32 read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoCrypt32.Create: _Crypt32;
begin
  Result := CreateComObject(CLASS_Crypt32) as _Crypt32;
end;

class function CoCrypt32.CreateRemote(const MachineName: string): _Crypt32;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Crypt32) as _Crypt32;
end;

procedure TCrypt32.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{A47EE28E-374D-4952-9361-8F728AFA3800}';
    IntfIID:   '{7F0AC8C6-D806-471A-A984-B06B223133B8}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCrypt32.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _Crypt32;
  end;
end;

procedure TCrypt32.ConnectTo(svrIntf: _Crypt32);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCrypt32.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCrypt32.GetDefaultInterface: _Crypt32;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TCrypt32.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TCrypt32Properties.Create(Self);
{$ENDIF}
end;

destructor TCrypt32.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TCrypt32.GetServerProperties: TCrypt32Properties;
begin
  Result := FProps;
end;
{$ENDIF}

function TCrypt32.AValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
begin
  Result := DefaultInterface.AValue(a, b, c);
end;

function TCrypt32.BValue(const a: WideString; var b: OleVariant; const c: WideString): WideString;
begin
  Result := DefaultInterface.BValue(a, b, c);
end;

function TCrypt32.CValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
begin
  Result := DefaultInterface.CValue(a, b, c);
end;

function TCrypt32.DValue(const a: WideString; var b: OleVariant; const c: WideString): WideString;
begin
  Result := DefaultInterface.DValue(a, b, c);
end;

function TCrypt32.EValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
begin
  Result := DefaultInterface.EValue(a, b, c);
end;

function TCrypt32.FValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
begin
  Result := DefaultInterface.FValue(a, b, c);
end;

function TCrypt32.GValue(var a: OleVariant; var b: OleVariant; var c: OleVariant): WideString;
begin
  Result := DefaultInterface.GValue(a, b, c);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TCrypt32Properties.Create(AServer: TCrypt32);
begin
  inherited Create;
  FServer := AServer;
end;

function TCrypt32Properties.GetDefaultInterface: _Crypt32;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TCrypt32]);
end;

end.
