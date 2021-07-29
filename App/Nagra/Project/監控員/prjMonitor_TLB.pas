unit prjMonitor_TLB;

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
// File generated on 2003/03/14 08:48:28 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\APP\Nagra\Project\ºÊ±±­û\prjMonitor.tlb (1)
// LIBID: {47850BF5-4103-4BA1-9736-7BDF6731E629}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

interface

uses ActiveX, Classes, Graphics, StdVCL, Variants, Windows;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  prjMonitorMajorVersion = 1;
  prjMonitorMinorVersion = 0;

  LIBID_prjMonitor: TGUID = '{47850BF5-4103-4BA1-9736-7BDF6731E629}';

  IID_IMonitor: TGUID = '{35848C21-7659-4B65-A265-176D85E9DBC1}';
  CLASS_Monitor: TGUID = '{4FBF8FC5-B886-4380-805D-8D0250B38F82}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IMonitor = interface;
  IMonitorDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Monitor = IMonitor;


// *********************************************************************//
// Interface: IMonitor
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {35848C21-7659-4B65-A265-176D85E9DBC1}
// *********************************************************************//
  IMonitor = interface(IDispatch)
    ['{35848C21-7659-4B65-A265-176D85E9DBC1}']
    procedure RunCa(sI_IniFileName: OleVariant); safecall;
  end;

// *********************************************************************//
// DispIntf:  IMonitorDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {35848C21-7659-4B65-A265-176D85E9DBC1}
// *********************************************************************//
  IMonitorDisp = dispinterface
    ['{35848C21-7659-4B65-A265-176D85E9DBC1}']
    procedure RunCa(sI_IniFileName: OleVariant); dispid 1;
  end;

// *********************************************************************//
// The Class CoMonitor provides a Create and CreateRemote method to          
// create instances of the default interface IMonitor exposed by              
// the CoClass Monitor. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMonitor = class
    class function Create: IMonitor;
    class function CreateRemote(const MachineName: string): IMonitor;
  end;

implementation

uses ComObj;

class function CoMonitor.Create: IMonitor;
begin
  Result := CreateComObject(CLASS_Monitor) as IMonitor;
end;

class function CoMonitor.CreateRemote(const MachineName: string): IMonitor;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Monitor) as IMonitor;
end;

end.
