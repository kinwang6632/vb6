unit Emc_Separate_TLB;

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

// PASTLWTR : $Revision:   1.130.3.0.1.0  $
// File generated on 2003/04/10 19:49:47 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\EMC\©î±b\project\Server\Emc_Separate.tlb (1)
// LIBID: {AF226D1C-A473-41AE-B700-3D06C324A5D6}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v1.0 Midas, (D:\WINNT\System32\midas.dll)
//   (2) v2.0 stdole, (D:\WINNT\SYSTEM32\STDOLE2.TLB)
//   (3) v4.0 StdVCL, (D:\WINNT\System32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

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
  Emc_SeparateMajorVersion = 1;
  Emc_SeparateMinorVersion = 0;

  LIBID_Emc_Separate: TGUID = '{AF226D1C-A473-41AE-B700-3D06C324A5D6}';

  IID_IEmc_Separate1: TGUID = '{778311E5-514A-4445-8A1E-11ED9A04547D}';
  CLASS_Emc_Separate1: TGUID = '{775D7E70-4ED9-41BF-B4EF-C49216E38A77}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IEmc_Separate1 = interface;
  IEmc_Separate1Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Emc_Separate1 = IEmc_Separate1;


// *********************************************************************//
// Interface: IEmc_Separate1
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {778311E5-514A-4445-8A1E-11ED9A04547D}
// *********************************************************************//
  IEmc_Separate1 = interface(IAppServer)
    ['{778311E5-514A-4445-8A1E-11ED9A04547D}']
    procedure connectToDB(sI_User: OleVariant; sI_CompCode: OleVariant; sI_DbUserName: OleVariant; 
                          sI_DbPassword: OleVariant; sI_DbAlias: OleVariant; 
                          var sI_Result: OleVariant); safecall;
    procedure runSF_CALCULATESO131(P_COMPCODE: OleVariant; P_SERVICETYPE: OleVariant; 
                                   P_COMPUTEYM: OleVariant; P_STARTDATE: OleVariant; 
                                   P_ENDDATE: OleVariant; P_CHARGEITEMSQL: OleVariant; 
                                   P_REALORSHOULDDATE: OleVariant; var p_RetCode: OleVariant; 
                                   var p_RetMsg: OleVariant); safecall;
    procedure getSO131Data(sI_SQL: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  IEmc_Separate1Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {778311E5-514A-4445-8A1E-11ED9A04547D}
// *********************************************************************//
  IEmc_Separate1Disp = dispinterface
    ['{778311E5-514A-4445-8A1E-11ED9A04547D}']
    procedure connectToDB(sI_User: OleVariant; sI_CompCode: OleVariant; sI_DbUserName: OleVariant; 
                          sI_DbPassword: OleVariant; sI_DbAlias: OleVariant; 
                          var sI_Result: OleVariant); dispid 2;
    procedure runSF_CALCULATESO131(P_COMPCODE: OleVariant; P_SERVICETYPE: OleVariant; 
                                   P_COMPUTEYM: OleVariant; P_STARTDATE: OleVariant; 
                                   P_ENDDATE: OleVariant; P_CHARGEITEMSQL: OleVariant; 
                                   P_REALORSHOULDDATE: OleVariant; var p_RetCode: OleVariant; 
                                   var p_RetMsg: OleVariant); dispid 1;
    procedure getSO131Data(sI_SQL: OleVariant); dispid 3;
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
// The Class CoEmc_Separate1 provides a Create and CreateRemote method to          
// create instances of the default interface IEmc_Separate1 exposed by              
// the CoClass Emc_Separate1. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoEmc_Separate1 = class
    class function Create: IEmc_Separate1;
    class function CreateRemote(const MachineName: string): IEmc_Separate1;
  end;

implementation

uses ComObj;

class function CoEmc_Separate1.Create: IEmc_Separate1;
begin
  Result := CreateComObject(CLASS_Emc_Separate1) as IEmc_Separate1;
end;

class function CoEmc_Separate1.CreateRemote(const MachineName: string): IEmc_Separate1;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Emc_Separate1) as IEmc_Separate1;
end;

end.
