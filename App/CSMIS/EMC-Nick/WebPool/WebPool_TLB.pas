unit WebPool_TLB;

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

// $Rev: 5081 $
// File generated on 2007/08/30 ¤U¤È 12:00:27 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\App\EMC\WebPool\WebPool.tlb (1)
// LIBID: {E1D3786E-6458-48B9-B163-49A0CAAD0CFD}
// LCID: 0
// Helpfile: 
// HelpString: CableSoft WebPool COM+ Component
// DepndLst: 
//   (1) v1.0 Midas, (C:\WINDOWS\system32\midas.dll)
//   (2) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, Midas, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  WebPoolMajorVersion = 1;
  WebPoolMinorVersion = 0;

  LIBID_WebPool: TGUID = '{E1D3786E-6458-48B9-B163-49A0CAAD0CFD}';

  IID_IPoolManager: TGUID = '{B4900D21-022B-471D-8C0E-368DE7176C9D}';
  CLASS_PoolManager: TGUID = '{3169B209-86A6-4CAB-A717-C2CA6F0C3684}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IPoolManager = interface;
  IPoolManagerDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  PoolManager = IPoolManager;


// *********************************************************************//
// Interface: IPoolManager
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B4900D21-022B-471D-8C0E-368DE7176C9D}
// *********************************************************************//
  IPoolManager = interface(IAppServer)
    ['{B4900D21-022B-471D-8C0E-368DE7176C9D}']
    function GetRecordSet(aSQL: OleVariant): OleVariant; safecall;
    function GetConnection: OleVariant; safecall;
    function GetCommand: OleVariant; safecall;
  end;

// *********************************************************************//
// DispIntf:  IPoolManagerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B4900D21-022B-471D-8C0E-368DE7176C9D}
// *********************************************************************//
  IPoolManagerDisp = dispinterface
    ['{B4900D21-022B-471D-8C0E-368DE7176C9D}']
    function GetRecordSet(aSQL: OleVariant): OleVariant; dispid 302;
    function GetConnection: OleVariant; dispid 301;
    function GetCommand: OleVariant; dispid 303;
    function AS_ApplyUpdates(const ProviderName: WideString; Delta: OleVariant; MaxErrors: Integer; 
                             out ErrorCount: Integer; var OwnerData: OleVariant): OleVariant; dispid 20000000;
    function AS_GetRecords(const ProviderName: WideString; Count: Integer; out RecsOut: Integer; 
                           Options: Integer; const CommandText: WideString; var Params: OleVariant; 
                           var OwnerData: OleVariant): OleVariant; dispid 20000001;
    function AS_DataRequest(const ProviderName: WideString; Data: OleVariant): OleVariant; dispid 20000002;
    function AS_GetProviderNames: OleVariant; dispid 20000003;
    function AS_GetParams(const ProviderName: WideString; var OwnerData: OleVariant): OleVariant; dispid 20000004;
    function AS_RowRequest(const ProviderName: WideString; Row: OleVariant; RequestType: Integer; 
                           var OwnerData: OleVariant): OleVariant; dispid 20000005;
    procedure AS_Execute(const ProviderName: WideString; const CommandText: WideString; 
                         var Params: OleVariant; var OwnerData: OleVariant); dispid 20000006;
  end;

// *********************************************************************//
// The Class CoPoolManager provides a Create and CreateRemote method to          
// create instances of the default interface IPoolManager exposed by              
// the CoClass PoolManager. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoPoolManager = class
    class function Create: IPoolManager;
    class function CreateRemote(const MachineName: string): IPoolManager;
  end;

implementation

uses ComObj;

class function CoPoolManager.Create: IPoolManager;
begin
  Result := CreateComObject(CLASS_PoolManager) as IPoolManager;
end;

class function CoPoolManager.CreateRemote(const MachineName: string): IPoolManager;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PoolManager) as IPoolManager;
end;

end.
