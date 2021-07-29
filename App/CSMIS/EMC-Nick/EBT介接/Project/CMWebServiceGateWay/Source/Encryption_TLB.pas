unit Encryption_TLB;

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
// File generated on 2001/11/14 ¤U¤È 03:15:57 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Delphi\ep.dll (1)
// LIBID: {493813DA-4890-11D5-B178-00485456B028}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (D:\WINNT\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (D:\WINNT\System32\stdvcl40.dll)
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

uses ActiveX, Classes, Graphics, OleServer, StdVCL, Variants, Windows;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  EncryptionMajorVersion = 1;
  EncryptionMinorVersion = 0;

  LIBID_Encryption: TGUID = '{493813DA-4890-11D5-B178-00485456B028}';

  IID__Password: TGUID = '{493813DB-4890-11D5-B178-00485456B028}';
  CLASS_Password: TGUID = '{493813DC-4890-11D5-B178-00485456B028}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _Password = interface;
  _PasswordDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Password = _Password;


// *********************************************************************//
// Interface: _Password
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {493813DB-4890-11D5-B178-00485456B028}
// *********************************************************************//
  _Password = interface(IDispatch)
    ['{493813DB-4890-11D5-B178-00485456B028}']
    function  Encrypt(const Plain: WideString; var sEncKey: WideString): WideString; safecall;
    function  Decrypt(const Encrypted: WideString; var sEncKey: WideString): WideString; safecall;
  end;

// *********************************************************************//
// DispIntf:  _PasswordDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {493813DB-4890-11D5-B178-00485456B028}
// *********************************************************************//
  _PasswordDisp = dispinterface
    ['{493813DB-4890-11D5-B178-00485456B028}']
    function  Encrypt(const Plain: WideString; var sEncKey: WideString): WideString; dispid 1610809344;
    function  Decrypt(const Encrypted: WideString; var sEncKey: WideString): WideString; dispid 1610809345;
  end;

// *********************************************************************//
// The Class CoPassword provides a Create and CreateRemote method to          
// create instances of the default interface _Password exposed by              
// the CoClass Password. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoPassword = class
    class function Create: _Password;
    class function CreateRemote(const MachineName: string): _Password;
  end;

implementation

uses ComObj;

class function CoPassword.Create: _Password;
begin
  Result := CreateComObject(CLASS_Password) as _Password;
end;

class function CoPassword.CreateRemote(const MachineName: string): _Password;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Password) as _Password;
end;

end.
