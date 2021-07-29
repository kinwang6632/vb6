unit cbClass;

interface

uses Classes;


type

  TExpoterKind = ( ekCA, ekSA, ekRA );
  
  TDownloadStatus = procedure (const aErrorText, aFileName: String) of object;
  
  TNotify = record
    MsgType: Integer;
    MsgText: String;
  end;

  TFreq = record
    CFR: Integer;
    LastExc: TDateTime;
    NextExc: TDateTime;
  end;

  TFTP = record
    UserId: String;
    Password: String;
    Host: String;
    Comps: String;
  end;

  TWatchFile= record
     FileName: String;
     LastSee: TDateTime;
     SendEmailCount: Integer;
  end;

  TEmailAccount = record
    Account: String;
    Password: String;
    Server: String;
    From: String;
  end;

const
  aDelimiter = Chr(1);

var
  TExpoterText: array [TExpoterKind] of String =  ( 'CA', 'SA', 'RA' );

  
implementation

end.
