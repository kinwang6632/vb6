{******************************************************************}
{                                                       	   }
{       Borland Delphi Runtime Library                  	   }
{       Windows Terminal Services interface unit                   }
{ 								   }
{ Portions created by Microsoft are 				   }
{ Copyright (C) 1995-1999 Microsoft Corporation. 		   }
{ All Rights Reserved. 						   }
{ 								   }
{ The original file is: wtsapi32.h, released June 2000. 	   }
{ The original Pascal code is: WtsApi32.pas, released Dec 2000     }
{ The initial developers of the Pascal code are Marcel van Brakel, }
{ Nathan Allan and Fabio Bomfim Nunes.        			   }
{ 								   }
{ Portions created by Marcel van Brakel are			   }
{ Copyright (C) 1999, 2000 Marcel van Brakel.            	   }
{ 								   }
{ Contributions:      					           }
{ 								   }
{   Fabio Bomfim Nunes (fbn@alba.com.br)                           }
{   Nathan Allan (nallan@enol.com)                                 }
{ 								   }
{ Obtained through:                               	           }
{ Joint Endeavour of Delphi Innovators (Project JEDI)              }
{								   }
{ You may retrieve the latest version of this file at the Project  }
{ JEDI home page, located at http://delphi-jedi.org                }
{								   }
{ The contents of this file are used with permission, subject to   }
{ the Mozilla Public License Version 1.1 (the "License"); you may  }
{ not use this file except in compliance with the License. You may }
{ obtain a copy of the License at                                  }
{ http://www.mozilla.org/MPL/MPL-1.1.html 	                   }
{                                                                  }
{ Software distributed under the License is distributed on an 	   }
{ "AS IS" basis, WITHOUT WARRANTY OF ANY KIND, either express or   }
{ implied. See the License for the specific language governing     }
{ rights and limitations under the License. 			   }
{ 								   }
{******************************************************************}

unit WtsApi32;

interface

uses
  Windows;

//==============================================================================
// Defines
//==============================================================================

//
//  Specifies the current server
//

const
  WTS_CURRENT_SERVER        = THandle(0);
  {$EXTERNALSYM WTS_CURRENT_SERVER}
  WTS_CURRENT_SERVER_HANDLE = THandle(0);
  {$EXTERNALSYM WTS_CURRENT_SERVER_HANDLE}
  WTS_CURRENT_SERVER_NAME   = '';
  {$EXTERNALSYM WTS_CURRENT_SERVER_NAME}
  SM_REMOTESESSION = $1000;
//
//  Specifies the current session (SessionId)
//

  WTS_CURRENT_SESSION = DWORD(-1);
  {$EXTERNALSYM WTS_CURRENT_SESSION}

//
//  Possible pResponse values from WTSSendMessage()
//

  IDTIMEOUT = 32000;
  {$EXTERNALSYM IDTIMEOUT}
  IDASYNC   = 32001;
  {$EXTERNALSYM IDASYNC}

//
//  Shutdown flags
//

  WTS_WSD_LOGOFF = $00000001;           // log off all users except
  {$EXTERNALSYM WTS_WSD_LOGOFF}         // current user; deletes
                                        // WinStations (a reboot is
                                        // required to recreate the
                                        // WinStations)
  WTS_WSD_SHUTDOWN = $00000002;         // shutdown system
  {$EXTERNALSYM WTS_WSD_SHUTDOWN}
  WTS_WSD_REBOOT   = $00000004;         // shutdown and reboot
  {$EXTERNALSYM WTS_WSD_REBOOT}
  WTS_WSD_POWEROFF = $00000008;         // shutdown and power off (on
  {$EXTERNALSYM WTS_WSD_POWEROFF}
                                        // machines that support power
                                        // off through software)
  WTS_WSD_FASTREBOOT = $00000010;       // reboot without logging users
  {$EXTERNALSYM WTS_WSD_FASTREBOOT}     // off or shutting down

//==============================================================================
// WTS_CONNECTSTATE_CLASS - Session connect state
//==============================================================================

type
  _WTS_CONNECTSTATE_CLASS = (
    WTSActive,              // User logged on to WinStation
    WTSConnected,           // WinStation connected to client
    WTSConnectQuery,        // In the process of connecting to client
    WTSShadow,              // Shadowing another WinStation
    WTSDisconnected,        // WinStation logged on without client
    WTSIdle,                // Waiting for client to connect
    WTSListen,              // WinStation is listening for connection
    WTSReset,               // WinStation is being reset
    WTSDown,                // WinStation is down due to error
    WTSInit);               // WinStation in initialization
  {$EXTERNALSYM _WTS_CONNECTSTATE_CLASS}
  WTS_CONNECTSTATE_CLASS = _WTS_CONNECTSTATE_CLASS;
  {$EXTERNALSYM WTS_CONNECTSTATE_CLASS}
  TWtsConnectStateClass = WTS_CONNECTSTATE_CLASS;

//==============================================================================
// WTS_SERVER_INFO - returned by WTSEnumerateServers (version 1)
//==============================================================================

//
//  WTSEnumerateServers() returns two variables: pServerInfo and Count.
//  The latter is the number of WTS_SERVER_INFO structures contained in
//  the former.  In order to read each server, iterate i from 0 to
//  Count-1 and reference the server name as
//  pServerInfo[i].pServerName; for example:
//
//  for ( i=0; i < Count; i++ ) {
//      _tprintf( TEXT("%s "), pServerInfo[i].pServerName );
//  }
//
//  The memory returned looks like the following.  P is a pServerInfo
//  pointer, and D is the string data for that pServerInfo:
//
//  P1 P2 P3 P4 ... Pn D1 D2 D3 D4 ... Dn
//
//  This makes it easier to iterate the servers, using code similar to
//  the above.
//

type
  PWTS_SERVER_INFOW = ^WTS_SERVER_INFOW;
  {$EXTERNALSYM PWTS_SERVER_INFOW}
  _WTS_SERVER_INFOW = record
    pServerName: PWideChar; // server name
  end;
  {$EXTERNALSYM _WTS_SERVER_INFOW}
  WTS_SERVER_INFOW = _WTS_SERVER_INFOW;
  {$EXTERNALSYM WTS_SERVER_INFOW}
  TWtsServerInfoW = WTS_SERVER_INFOW;
  PWtsServerInfoW = PWTS_SERVER_INFOW;

  PWTS_SERVER_INFOA = ^WTS_SERVER_INFOA;
  {$EXTERNALSYM PWTS_SERVER_INFOA}
  _WTS_SERVER_INFOA = record
    pServerName: PAnsiChar; // server name
  end;
  {$EXTERNALSYM _WTS_SERVER_INFOA}
  WTS_SERVER_INFOA = _WTS_SERVER_INFOA;
  {$EXTERNALSYM WTS_SERVER_INFOA}
  TWtsServerInfoA = WTS_SERVER_INFOA;
  PWtsServerInfoA = PWTS_SERVER_INFOA;

{$IFDEF UNICODE}
  WTS_SERVER_INFO = WTS_SERVER_INFOW;
  {$EXTERNALSYM WTS_SERVER_INFO}
  PWTS_SERVER_INFO = PWTS_SERVER_INFOW;
  {$EXTERNALSYM PWTS_SERVER_INFO}
  TWtsServerInfo = TWtsServerInfoW;
  PWtsServerInfo = PWtsServerInfoW;
{$ELSE}
  WTS_SERVER_INFO = WTS_SERVER_INFOA;
  {$EXTERNALSYM WTS_SERVER_INFO}
  PWTS_SERVER_INFO = PWTS_SERVER_INFOA;
  {$EXTERNALSYM PWTS_SERVER_INFO}
  TWtsServerInfo = TWtsServerInfoA;
  PWtsServerInfo = PWtsServerInfoA;
{$ENDIF}

//==============================================================================
// WTS_SESSION_INFO - returned by WTSEnumerateSessions (version 1)
//==============================================================================

//
//  WTSEnumerateSessions() returns data in a similar format to the above
//  WTSEnumerateServers().  It returns two variables: pSessionInfo and
//  Count.  The latter is the number of WTS_SESSION_INFO structures
//  contained in the former.  Iteration is similar, except that there
//  are three parts to each entry, so it would look like this:
//
//  for ( i=0; i < Count; i++ ) {
//      _tprintf( TEXT("%-5u  %-20s  %u\n"),
//                pSessionInfo[i].SessionId,
//                pSessionInfo[i].pWinStationName,
//                pSessionInfo[i].State );
//  }
//
//  The memory returned is also segmented as the above, with all the
//  structures allocated at the start and the string data at the end.
//  We'll use S for the SessionId, P for the pWinStationName pointer
//  and D for the string data, and C for the connect State:
//
//  S1 P1 C1 S2 P2 C2 S3 P3 C3 S4 P4 C4 ... Sn Pn Cn D1 D2 D3 D4 ... Dn
//
//  As above, this makes it easier to iterate the sessions.
//

type
  PWTS_SESSION_INFOW = ^WTS_SESSION_INFOW;
  {$EXTERNALSYM PWTS_SESSION_INFOW}
  _WTS_SESSION_INFOW = record
    SessionId: DWORD;              // session id
    pWinStationName: PWideChar;       // name of WinStation this session is connected to
    State: TWtsConnectStateClass; // connection state (see enum)
  end;
  {$EXTERNALSYM _WTS_SESSION_INFOW}
  WTS_SESSION_INFOW = _WTS_SESSION_INFOW;
  {$EXTERNALSYM WTS_SESSION_INFOW}
  TWtsSessionInfoW = WTS_SESSION_INFOW;
  PWtsSessionInfoW = PWTS_SESSION_INFOW;

  PWTS_SESSION_INFOA = ^WTS_SESSION_INFOA;
  {$EXTERNALSYM PWTS_SESSION_INFOA}
  _WTS_SESSION_INFOA = record
    SessionId: DWORD;              // session id
    pWinStationName: PAnsiChar;        // name of WinStation this session is connected to
    State: WTS_CONNECTSTATE_CLASS; // connection state (see enum)
  end;
  {$EXTERNALSYM _WTS_SESSION_INFOA}
  WTS_SESSION_INFOA = _WTS_SESSION_INFOA;
  {$EXTERNALSYM WTS_SESSION_INFOA}
  TWtsSessionInfoA = WTS_SESSION_INFOA;
  PWtsSessionInfoA = PWTS_SESSION_INFOA;

{$IFDEF UNICODE}
  WTS_SESSION_INFO = WTS_SESSION_INFOW;
  PWTS_SESSION_INFO = PWTS_SESSION_INFOW;
  TWtsSessionInfo = TWtsSessionInfoW;
  PWtsSessionInfo = PWtsSessionInfoW;
{$ELSE}
  WTS_SESSION_INFO = WTS_SESSION_INFOA;
  PWTS_SESSION_INFO = PWTS_SESSION_INFOA;
  TWtsSessionInfo = TWtsSessionInfoA;
  PWtsSessionInfo = PWtsSessionInfoA;
{$ENDIF}

//==============================================================================
// WTS_PROCESS_INFO - returned by WTSEnumerateProcesses (version 1)
//==============================================================================

//
//  WTSEnumerateProcesses() also returns data similar to
//  WTSEnumerateServers().  It returns two variables: pProcessInfo and
//  Count.  The latter is the number of WTS_PROCESS_INFO structures
//  contained in the former.  Iteration is similar, except that there
//  are four parts to each entry, so it would look like this:
//
//  for ( i=0; i < Count; i++ ) {
//      GetUserNameFromSid( pProcessInfo[i].pUserSid, UserName,
//                          sizeof(UserName) );
//      _tprintf( TEXT("%-5u  %-20s  %-5u  %s\n"),
//              pProcessInfo[i].SessionId,
//              UserName,
//              pProcessInfo[i].ProcessId,
//              pProcessInfo[i].pProcessName );
//  }
//
//  The memory returned is also segmented as the above, with all the
//  structures allocated at the start and the string data at the end.
//  We'll use S for the SessionId, R for the ProcessId, P for the
//  pProcessName pointer and D for the string data, and U for pUserSid:
//
//  S1 R1 P1 U1 S2 R2 P2 U2 S3 R3 P3 U3 ... Sn Rn Pn Un D1 D2 D3 ... Dn
//
//  As above, this makes it easier to iterate the processes.
//

type
  PWTS_PROCESS_INFOW = ^WTS_PROCESS_INFOW;
  {$EXTERNALSYM PWTS_PROCESS_INFOW}
  _WTS_PROCESS_INFOW = record
    SessionId: DWORD;     // session id
    ProcessId: DWORD;     // process id
    pProcessName: PWideChar; // name of process
    pUserSid: PSID;       // user's SID
  end;
  {$EXTERNALSYM _WTS_PROCESS_INFOW}
  WTS_PROCESS_INFOW = _WTS_PROCESS_INFOW;
  {$EXTERNALSYM WTS_PROCESS_INFOW}
  TWtsProcessInfoW = WTS_PROCESS_INFOW;
  PWtsProcessInfoW = PWTS_PROCESS_INFOW;

  PWTS_PROCESS_INFOA = ^WTS_PROCESS_INFOA;
  {$EXTERNALSYM PWTS_PROCESS_INFOA}
  _WTS_PROCESS_INFOA = record
    SessionId: DWORD;    // session id
    ProcessId: DWORD;    // process id
    pProcessName: PAnsiChar; // name of process
    pUserSid: PSID;      // user's SID
  end;
  {$EXTERNALSYM _WTS_PROCESS_INFOA}
  WTS_PROCESS_INFOA = _WTS_PROCESS_INFOA;
  {$EXTERNALSYM WTS_PROCESS_INFOA}
  TWtsProcessInfoA = WTS_PROCESS_INFOA;
  PWtsProcessInfoA = PWTS_PROCESS_INFOA;

{$IFDEF UNICODE}
  WTS_PROCESS_INFO = WTS_PROCESS_INFOW;
  {$EXTERNALSYM WTS_PROCESS_INFO}
  PWTS_PROCESS_INFO = PWTS_PROCESS_INFOW;
  {$EXTERNALSYM PWTS_PROCESS_INFO}
  TWtsProcessInfo = TWtsProcessInfoW;
  PWtsProcessInfo = PWtsProcessInfoW;
{$ELSE}
  WTS_PROCESS_INFO = WTS_PROCESS_INFOA;
  {$EXTERNALSYM WTS_PROCESS_INFO}
  PWTS_PROCESS_INFO = PWTS_PROCESS_INFOA;
  {$EXTERNALSYM PWTS_PROCESS_INFO}
  TWtsProcessInfo = TWtsProcessInfoA;
  PWtsProcessInfo = PWtsProcessInfoA;
{$ENDIF}

//==============================================================================
// WTS_INFO_CLASS - WTSQuerySessionInformation
// (See additional typedefs for more info on structures)
//==============================================================================

const
  WTS_PROTOCOL_TYPE_CONSOLE = 0; // Console
  {$EXTERNALSYM WTS_PROTOCOL_TYPE_CONSOLE}
  WTS_PROTOCOL_TYPE_ICA     = 1; // ICA Protocol
  {$EXTERNALSYM WTS_PROTOCOL_TYPE_ICA}
  WTS_PROTOCOL_TYPE_RDP     = 2; // RDP Protocol
  {$EXTERNALSYM WTS_PROTOCOL_TYPE_RDP}

type
  _WTS_INFO_CLASS = (
    WTSInitialProgram,
    WTSApplicationName,
    WTSWorkingDirectory,
    WTSOEMId,
    WTSSessionId,
    WTSUserName,
    WTSWinStationName,
    WTSDomainName,
    WTSConnectState,
    WTSClientBuildNumber,
    WTSClientName,
    WTSClientDirectory,
    WTSClientProductId,
    WTSClientHardwareId,
    WTSClientAddress,
    WTSClientDisplay,
    WTSClientProtocolType);
  {$EXTERNALSYM _WTS_INFO_CLASS}
  WTS_INFO_CLASS = _WTS_INFO_CLASS;
  TWtsInfoClass = WTS_INFO_CLASS;

//==============================================================================
// WTSQuerySessionInformation - (WTSClientAddress)
//==============================================================================

type
  PWTS_CLIENT_ADDRESS = ^WTS_CLIENT_ADDRESS;
  {$EXTERNALSYM PWTS_CLIENT_ADDRESS}
  _WTS_CLIENT_ADDRESS = record
    AddressFamily: DWORD;           // AF_INET, AF_IPX, AF_NETBIOS, AF_UNSPEC
    Address: array [0..19] of Byte; // client network address
  end;
  {$EXTERNALSYM _WTS_CLIENT_ADDRESS}
  WTS_CLIENT_ADDRESS = _WTS_CLIENT_ADDRESS;
  {$EXTERNALSYM WTS_CLIENT_ADDRESS}
  TWtsClientAddress = WTS_CLIENT_ADDRESS;
  PWtsClientAddress = PWTS_CLIENT_ADDRESS;

//==============================================================================
// WTSQuerySessionInformation - (WTSClientDisplay)
//==============================================================================

type
  PWTS_CLIENT_DISPLAY = ^WTS_CLIENT_DISPLAY;
  {$EXTERNALSYM PWTS_CLIENT_DISPLAY}
  _WTS_CLIENT_DISPLAY = record
    HorizontalResolution: DWORD; // horizontal dimensions, in pixels
    VerticalResolution: DWORD;   // vertical dimensions, in pixels
    ColorDepth: DWORD;           // 1=16, 2=256, 4=64K, 8=16M
  end;
  {$EXTERNALSYM _WTS_CLIENT_DISPLAY}
  WTS_CLIENT_DISPLAY = _WTS_CLIENT_DISPLAY;
  {$EXTERNALSYM WTS_CLIENT_DISPLAY}
  TWtsClientDisplay = WTS_CLIENT_DISPLAY;
  PWtsClientDisplay = PWTS_CLIENT_DISPLAY;

//==============================================================================
// WTS_CONFIG_CLASS - WTSQueryUserConfig/WTSSetUserConfig
//==============================================================================

type
  _WTS_CONFIG_CLASS = (
    //Initial program settings
    WTSUserConfigInitialProgram,         	// string returned/expected
    WTSUserConfigWorkingDirectory,       	// string returned/expected
    WTSUserConfigfInheritInitialProgram, 	// DWORD returned/expected
    //
    WTSUserConfigfAllowLogonTerminalServer, 	//DWORD returned/expected
    //Timeout settings
    WTSUserConfigTimeoutSettingsConnections, 	//DWORD returned/expected
    WTSUserConfigTimeoutSettingsDisconnections, //DWORD returned/expected
    WTSUserConfigTimeoutSettingsIdle, 	        //DWORD returned/expected
    //Client device settings
    WTSUserConfigfDeviceClientDrives,  		//DWORD returned/expected
    WTSUserConfigfDeviceClientPrinters,         //DWORD returned/expected
    WTSUserConfigfDeviceClientDefaultPrinter,   //DWORD returned/expected
    //Connection settings
    WTSUserConfigBrokenTimeoutSettings,         //DWORD returned/expected
    WTSUserConfigReconnectSettings,             //DWORD returned/expected
    //Modem settings
    WTSUserConfigModemCallbackSettings,         //DWORD returned/expected
    WTSUserConfigModemCallbackPhoneNumber,      // string returned/expected
    //Shadow settings
    WTSUserConfigShadowingSettings,             //DWORD returned/expected
    //User Profile settings
    WTSUserConfigTerminalServerProfilePath,     // string returned/expected
    //Terminal Server home directory
    WTSUserConfigTerminalServerHomeDir,         // string returned/expected
    WTSUserConfigTerminalServerHomeDirDrive,    // string returned/expected
    WTSUserConfigfTerminalServerRemoteHomeDir); // DWORD 0:LOCAL 1:REMOTE
  {$EXTERNALSYM _WTS_CONFIG_CLASS}
  WTS_CONFIG_CLASS = _WTS_CONFIG_CLASS;
  TWtsConfigClass = WTS_CONFIG_CLASS;

  PWTS_USER_CONFIG_SET_NWSERVERW = ^WTS_USER_CONFIG_SET_NWSERVERW;
  {$EXTERNALSYM PWTS_USER_CONFIG_SET_NWSERVERW}
  _WTS_USER_CONFIG_SET_NWSERVERW = record
    pNWServerName: PWideChar;
    pNWDomainAdminName: PWideChar;
    pNWDomainAdminPassword: PWideChar;
  end;
  {$EXTERNALSYM _WTS_USER_CONFIG_SET_NWSERVERW}
  WTS_USER_CONFIG_SET_NWSERVERW = _WTS_USER_CONFIG_SET_NWSERVERW;
  {$EXTERNALSYM WTS_USER_CONFIG_SET_NWSERVERW}
  TWtsUserConfigSetNwserverW = WTS_USER_CONFIG_SET_NWSERVERW;
  PWtsUserConfigSetNwserverW = PWTS_USER_CONFIG_SET_NWSERVERW;

  PWTS_USER_CONFIG_SET_NWSERVERA = ^WTS_USER_CONFIG_SET_NWSERVERA;
  {$EXTERNALSYM PWTS_USER_CONFIG_SET_NWSERVERA}
  _WTS_USER_CONFIG_SET_NWSERVERA = record
    pNWServerName: PAnsiChar;
    pNWDomainAdminName: PAnsiChar;
    pNWDomainAdminPassword: PAnsiChar;
  end;
  {$EXTERNALSYM _WTS_USER_CONFIG_SET_NWSERVERA}
  WTS_USER_CONFIG_SET_NWSERVERA = _WTS_USER_CONFIG_SET_NWSERVERA;
  {$EXTERNALSYM WTS_USER_CONFIG_SET_NWSERVERA}
  TWtsUserConfigSetNwserverA = WTS_USER_CONFIG_SET_NWSERVERA;
  PWtsUserConfigSetNwserverA = PWTS_USER_CONFIG_SET_NWSERVERA;

{$IFDEF UNICODE}
  WTS_USER_CONFIG_SET_NWSERVER  = WTS_USER_CONFIG_SET_NWSERVERW;
  {$EXTERNALSYM WTS_USER_CONFIG_SET_NWSERVER}
  PWTS_USER_CONFIG_SET_NWSERVER = PWTS_USER_CONFIG_SET_NWSERVERW;
  {$EXTERNALSYM PWTS_USER_CONFIG_SET_NWSERVER}
  TWtsUserConfigSetNwserver = TWtsUserConfigSetNwserverW;
  PWtsUserConfigSetNwserver = PWtsUserConfigSetNwserverW;
{$ELSE}
  WTS_USER_CONFIG_SET_NWSERVER  = WTS_USER_CONFIG_SET_NWSERVERA;
  {$EXTERNALSYM WTS_USER_CONFIG_SET_NWSERVER}
  PWTS_USER_CONFIG_SET_NWSERVER = PWTS_USER_CONFIG_SET_NWSERVERA;
  {$EXTERNALSYM PWTS_USER_CONFIG_SET_NWSERVER}
  TWtsUserConfigSetNwserver = TWtsUserConfigSetNwserverA;
  PWtsUserConfigSetNwserver = PWtsUserConfigSetNwserverA;
{$ENDIF}

//==============================================================================
// WTS_EVENT - Event flags for WTSWaitSystemEvent
//==============================================================================

const
  WTS_EVENT_NONE        = $00000000; // return no event
  {$EXTERNALSYM WTS_EVENT_NONE}
  WTS_EVENT_CREATE      = $00000001; // new WinStation created
  {$EXTERNALSYM WTS_EVENT_CREATE}
  WTS_EVENT_DELETE      = $00000002; // existing WinStation deleted
  {$EXTERNALSYM WTS_EVENT_DELETE}
  WTS_EVENT_RENAME      = $00000004; // existing WinStation renamed
  {$EXTERNALSYM WTS_EVENT_RENAME}
  WTS_EVENT_CONNECT     = $00000008; // WinStation connect to client
  {$EXTERNALSYM WTS_EVENT_CONNECT}
  WTS_EVENT_DISCONNECT  = $00000010; // WinStation logged on without client
  {$EXTERNALSYM WTS_EVENT_DISCONNECT}
  WTS_EVENT_LOGON       = $00000020; // user logged on to existing WinStation
  {$EXTERNALSYM WTS_EVENT_LOGON}
  WTS_EVENT_LOGOFF      = $00000040; // user logged off from existing WinStation
  {$EXTERNALSYM WTS_EVENT_LOGOFF}
  WTS_EVENT_STATECHANGE = $00000080; // WinStation state change
  {$EXTERNALSYM WTS_EVENT_STATECHANGE}
  WTS_EVENT_LICENSE     = $00000100; // license state change
  {$EXTERNALSYM WTS_EVENT_LICENSE}
  WTS_EVENT_ALL         = $7fffffff; // wait for all event types
  {$EXTERNALSYM WTS_EVENT_ALL}
  WTS_EVENT_FLUSH       = DWORD($80000000); // unblock all waiters
  {$EXTERNALSYM WTS_EVENT_FLUSH}

//==============================================================================
// WTS_VIRTUAL_CLASS - WTSVirtualChannelQuery
//==============================================================================

type
  _WTS_VIRTUAL_CLASS = (WTSVirtualClientData);  // Virtual channel client module data (C2H data)
  {$EXTERNALSYM _WTS_VIRTUAL_CLASS}
  WTS_VIRTUAL_CLASS = _WTS_VIRTUAL_CLASS;
  {$EXTERNALSYM WTS_VIRTUAL_CLASS}
  TWtsVirtualClass = WTS_VIRTUAL_CLASS;

//==============================================================================
// Windows Terminal Server public APIs
//==============================================================================

function WTSEnumerateServersA(pDomainName: PAnsiChar; Reserved, Version: DWORD;
  var ppServerInfo: PWtsServerInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateServersA}
function WTSEnumerateServersW(pDomainName: PWideChar; Reserved, Version: DWORD;
  var ppServerInfo: PWtsServerInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateServersW}

{$IFDEF UNICODE}
function WTSEnumerateServers(pDomainName: PWideChar; Reserved, Version: DWORD;
  var ppServerInfo: PWtsServerInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateServers}
{$ELSE}
function WTSEnumerateServers(pDomainName: PAnsiChar; Reserved, Version: DWORD;
  var ppServerInfo: PWtsServerInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateServers}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSOpenServerA(pServerName: PAnsiChar): THandle; stdcall;
{$EXTERNALSYM WTSOpenServerA}
function WTSOpenServerW(pServerName: PWideChar): THandle; stdcall;
{$EXTERNALSYM WTSOpenServerW}

{$IFDEF UNICODE}
function WTSOpenServer(pServerName: PWideChar): THandle; stdcall;
{$EXTERNALSYM WTSOpenServer}
{$ELSE}
function WTSOpenServer(pServerName: PAnsiChar): THandle; stdcall;
{$EXTERNALSYM WTSOpenServer}
{$ENDIF}

//------------------------------------------------------------------------------

procedure WTSCloseServer(hServer: THandle); stdcall;
{$EXTERNALSYM WTSCloseServer}

//------------------------------------------------------------------------------

function WTSEnumerateSessionsA(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppSessionInfo: PWtsSessionInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateSessionsA}
function WTSEnumerateSessionsW(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppSessionInfo: PWtsSessionInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateSessionsW}

{$IFDEF UNICODE}
function WTSEnumerateSessions(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppSessionInfo: PWtsSessionInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateSessions}
{$ELSE}
function WTSEnumerateSessions(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppSessionInfo: PWtsSessionInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateSessions}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSEnumerateProcessesA(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppProcessInfo: PWtsProcessInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateProcessesA}
function WTSEnumerateProcessesW(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppProcessInfo: PWtsProcessInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateProcessesW}

{$IFDEF UNICODE}
function WTSEnumerateProcesses(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppProcessInfo: PWtsProcessInfoW; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateProcesses}
{$ELSE}
function WTSEnumerateProcesses(hServer: THandle; Reserved: DWORD; Version: DWORD;
  var ppProcessInfo: PWtsProcessInfoA; var pCount: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSEnumerateProcesses}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSTerminateProcess(hServer: THandle; ProcessId, ExitCode: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSTerminateProcess}

//------------------------------------------------------------------------------

function WTSQuerySessionInformationA(hServer: THandle; SessionId: DWORD;
  WTSInfoClass: TWtsInfoClass; var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQuerySessionInformationA}
function WTSQuerySessionInformationW(hServer: THandle; SessionId: DWORD;
  WTSInfoClass: TWtsInfoClass; var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQuerySessionInformationW}

{$IFDEF UNICODE}
function WTSQuerySessionInformation(hServer: THandle; SessionId: DWORD;
  WTSInfoClass: TWtsInfoClass; var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQuerySessionInformation}
{$ELSE}
function WTSQuerySessionInformation(hServer: THandle; SessionId: DWORD;
  WTSInfoClass: TWtsInfoClass; var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQuerySessionInformation}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSQueryUserConfigA(pServerName, pUserName: PAnsiChar; WTSConfigClass: TWtsConfigClass;
  var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQueryUserConfigA}
function WTSQueryUserConfigW(pServerName, pUserName: PWideChar; WTSConfigClass: TWtsConfigClass;
  var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQueryUserConfigW}

{$IFDEF UNICODE}
function WTSQueryUserConfig(pServerName, pUserName: PWideChar; WTSConfigClass: TWtsConfigClass;
  var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQueryUserConfig}
{$ELSE}
function WTSQueryUserConfig(pServerName, pUserName: PAnsiChar; WTSConfigClass: TWtsConfigClass;
  var ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSQueryUserConfig}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSSetUserConfigA(pServerName, pUserName: PAnsiChar; WTSConfigClass: TWtsConfigClass;
  pBuffer: PAnsiChar; DataLength: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSSetUserConfigA}
function WTSSetUserConfigW(pServerName, pUserName: PWideChar; WTSConfigClass: TWtsConfigClass;
  pBuffer: PWideChar; DataLength: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSSetUserConfigW}

{$IFDEF UNICODE}
function WTSSetUserConfig(pServerName, pUserName: PWideChar; WTSConfigClass: TWtsConfigClass;
  pBuffer: PWideChar; DataLength: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSSetUserConfig}
{$ELSE}
function WTSSetUserConfig(pServerName, pUserName: PAnsiChar; WTSConfigClass: TWtsConfigClass;
  pBuffer: PAnsiChar; DataLength: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSSetUserConfig}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSSendMessageA(hServer: THandle; SessionId: DWORD; pTitle: PAnsiChar;
  TitleLength: DWORD; pMessage: PAnsiChar; MessageLength: DWORD; Style: DWORD;
  Timeout: DWORD; var pResponse: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSSendMessageA}
function WTSSendMessageW(hServer: THandle; SessionId: DWORD; pTitle: PWideChar;
  TitleLength: DWORD; pMessage: PWideChar; MessageLength: DWORD; Style: DWORD;
  Timeout: DWORD; var pResponse: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSSendMessageW}

{$IFDEF UNICODE}
function WTSSendMessage(hServer: THandle; SessionId: DWORD; pTitle: PWideChar;
  TitleLength: DWORD; pMessage: PWideChar; MessageLength: DWORD; Style: DWORD;
  Timeout: DWORD; var pResponse: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSSendMessage}
{$ELSE}
function WTSSendMessage(hServer: THandle; SessionId: DWORD; pTitle: PAnsiChar;
  TitleLength: DWORD; pMessage: PAnsiChar; MessageLength: DWORD; Style: DWORD;
  Timeout: DWORD; var pResponse: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSSendMessage}
{$ENDIF}

//------------------------------------------------------------------------------

function WTSDisconnectSession(hServer: THandle; SessionId: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSDisconnectSession}

//------------------------------------------------------------------------------

function WTSLogoffSession(hServer: THandle; SessionId: DWORD; bWait: BOOL): BOOL; stdcall;
{$EXTERNALSYM WTSLogoffSession}

//------------------------------------------------------------------------------

function WTSShutdownSystem(hServer: THandle; ShutdownFlag: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSShutdownSystem}

//------------------------------------------------------------------------------

function WTSWaitSystemEvent(hServer: THandle; EventMask: DWORD;
  var pEventFlags: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSWaitSystemEvent}

//------------------------------------------------------------------------------

function WTSVirtualChannelOpen(hServer: THandle; SessionId: DWORD;
  pVirtualName: PAnsiChar): THandle; stdcall;
{$EXTERNALSYM WTSVirtualChannelOpen}

function WTSVirtualChannelClose(hChannelHandle: THandle): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelClose}

function WTSVirtualChannelRead(hChannelHandle: THandle; TimeOut: ULONG;
  Buffer: PChar; BufferSize: ULONG; var pBytesRead: ULONG): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelRead}

function WTSVirtualChannelWrite(hChannelHandle: THandle; Buffer: PChar;
  Length: ULONG; var pBytesWritten: ULONG): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelWrite}

function WTSVirtualChannelPurgeInput(hChannelHandle: THandle): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelPurgeInput}

function WTSVirtualChannelPurgeOutput(hChannelHandle: THandle): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelPurgeOutput}

function WTSVirtualChannelQuery(hChannelHandle: THandle; VirtualClass: TWtsVirtualClass;
  ppBuffer: Pointer; var pBytesReturned: DWORD): BOOL; stdcall;
{$EXTERNALSYM WTSVirtualChannelQuery}

//------------------------------------------------------------------------------

procedure WTSFreeMemory(pMemory: Pointer); stdcall;
{$EXTERNALSYM WTSFreeMemory}

implementation

const
  wtsapi = 'wtsapi32.dll';

function WTSEnumerateServersA; external wtsapi name 'WTSEnumerateServersA';
function WTSEnumerateServersW; external wtsapi name 'WTSEnumerateServersW';
{$IFDEF UNICODE}
function WTSEnumerateServers; external wtsapi name 'WTSEnumerateServersW';
{$ELSE}
function WTSEnumerateServers; external wtsapi name 'WTSEnumerateServersA';
{$ENDIF}
function WTSOpenServerA; external wtsapi name 'WTSOpenServerA';
function WTSOpenServerW; external wtsapi name 'WTSOpenServerW';
{$IFDEF UNICODE}
function WTSOpenServer; external wtsapi name 'WTSOpenServerW';
{$ELSE}
function WTSOpenServer; external wtsapi name 'WTSOpenServerA';
{$ENDIF}
procedure WTSCloseServer; external wtsapi name 'WTSCloseServer';
function WTSEnumerateSessionsA; external wtsapi name 'WTSEnumerateSessionsA';
function WTSEnumerateSessionsW; external wtsapi name 'WTSEnumerateSessionsW';
{$IFDEF UNICODE}
function WTSEnumerateSessions; external wtsapi name 'WTSEnumerateSessionsW';
{$ELSE}
function WTSEnumerateSessions; external wtsapi name 'WTSEnumerateSessionsA';
{$ENDIF}
function WTSEnumerateProcessesA; external wtsapi name 'WTSEnumerateProcessesA';
function WTSEnumerateProcessesW; external wtsapi name 'WTSEnumerateProcessesW';
{$IFDEF UNICODE}
function WTSEnumerateProcesses; external wtsapi name 'WTSEnumerateProcessesW';
{$ELSE}
function WTSEnumerateProcesses; external wtsapi name 'WTSEnumerateProcessesA';
{$ENDIF}
function WTSTerminateProcess; external wtsapi name 'WTSTerminateProcess';
function WTSQuerySessionInformationA; external wtsapi name 'WTSQuerySessionInformationA';
function WTSQuerySessionInformationW; external wtsapi name 'WTSQuerySessionInformationW';
{$IFDEF UNICODE}
function WTSQuerySessionInformation; external wtsapi name 'WTSQuerySessionInformationW';
{$ELSE}
function WTSQuerySessionInformation; external wtsapi name 'WTSQuerySessionInformationA';
{$ENDIF}
function WTSQueryUserConfigA; external wtsapi name 'WTSQueryUserConfigA';
function WTSQueryUserConfigW; external wtsapi name 'WTSQueryUserConfigW';
{$IFDEF UNICODE}
function WTSQueryUserConfig; external wtsapi name 'WTSQueryUserConfigW';
{$ELSE}
function WTSQueryUserConfig; external wtsapi name 'WTSQueryUserConfigA';
{$ENDIF}
function WTSSetUserConfigA; external wtsapi name 'WTSSetUserConfigA';
function WTSSetUserConfigW; external wtsapi name 'WTSSetUserConfigW';
{$IFDEF UNICODE}
function WTSSetUserConfig; external wtsapi name 'WTSSetUserConfigW';
{$ELSE}
function WTSSetUserConfig; external wtsapi name 'WTSSetUserConfigA';
{$ENDIF}
function WTSSendMessageA; external wtsapi name 'WTSSendMessageA';
function WTSSendMessageW; external wtsapi name 'WTSSendMessageW';
{$IFDEF UNICODE}
function WTSSendMessage; external wtsapi name 'WTSSendMessageW'
{$ELSE}
function WTSSendMessage; external wtsapi name 'WTSSendMessageA';
{$ENDIF}
function WTSDisconnectSession; external wtsapi name 'WTSDisconnectSession';
function WTSLogoffSession; external wtsapi name 'WTSLogoffSession';
function WTSShutdownSystem; external wtsapi name 'WTSShutdownSystem';
function WTSWaitSystemEvent; external wtsapi name 'WTSWaitSystemEvent';
function WTSVirtualChannelOpen; external wtsapi name 'WTSVirtualChannelOpen';
function WTSVirtualChannelClose; external wtsapi name 'WTSVirtualChannelClose';
function WTSVirtualChannelRead; external wtsapi name 'WTSVirtualChannelRead';
function WTSVirtualChannelWrite; external wtsapi name 'WTSVirtualChannelWrite';
function WTSVirtualChannelPurgeInput; external wtsapi name 'WTSVirtualChannelPurgeInput';
function WTSVirtualChannelPurgeOutput; external wtsapi name 'WTSVirtualChannelPurgeOutput';
function WTSVirtualChannelQuery; external wtsapi name 'WTSVirtualChannelQuery';
procedure WTSFreeMemory; external wtsapi name 'WTSFreeMemory';

end.
