unit Uotheru;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Controls, Db, Forms,
  Dialogs, Math, inifiles,
  QuickRpt,QRCtrls, ExtCtrls, QRPrntr, StdCtrls,Winsock, Registry;

const
    ALIGNMENT_LEFT_FLAG = 'L';
    ALIGNMENT_CENTER_FLAG = 'C';
    ALIGNMENT_RIGHT_FLAG = 'R';

    REPORT_SECTION_NAME = 'REPORT';
    BAND_SECTION_NAME= 'BAND_INFO';
    PARENT_TAG = 'PARENT';
    CLASS_NAME_TAG = 'CLASS';
    NONE_FOOTER_BAND_FLAG = '***';
    NONE_CHILD_NAME_FLAG = '***';
Type

  TSex = (sFemale, sMale);

  TUother = class
  private

    class function LiteralChk(asId:string): Boolean;
    class function LengthChk(asId:string): Boolean ;
    class function TellSex(asId:string; var aSex:TSex): Boolean;
  public
    class function regWrite(sI_KeyPath, sI_KeyNam:string;nI_KeyStyle:integer;vI_KeyVal:Variant): boolean;
    class function regRead(sI_KeyPath, sI_KeyNam:string; nI_KeyStyle:integer): Variant;
    class function myEncrypt(sI_SourceStr, sI_Key:String):String;
    class function myDecrypt(sI_SourceStr, sI_Key:String):String;

    class function getLocalIp:String;
    class function getAlignmentFlag(I_Alignment:TAlignment):String;
    class function getRealAlignmentValue(sI_AlignmentFlag:String):TAlignment;

    class function getBandFlag(I_QRBand : TQRCustomBand):String;
    class function getRealBandType(sI_BandTypeFlag : String):TQrBandType;
//    class function getOrientationFlag(I_Orientation : TPrinterOrientation):String;
    class function CreateQRControl(QRControlClass:TControlClass; aParent:TControl; I_QuickRep:TQuickRep) : TControl ;
    class procedure restoreControlInfo(I_ParentControl:TControl; sI_IniFileName:String);//將 I_ParentControl 的 child control 的位置..等 information 從 ini file 中回存
    class procedure saveControlInfo(I_ParentControl:TControl; sI_IniFileName:String);//將 I_ParentControl 的 child control 的位置..等 information 存到 ini file 中
  
    class function RoundFloat(Src : Double; nRetainPos : Integer): Double;
    class function CheckIdNo(asId:string; var iSex:TSex):Boolean;
    class function RunShell(sI_ExeName, sI_Param: string; bI_WithUI, bI_WaitProcess: Boolean; var sI_ErrorMsg:String): Boolean;
    class procedure PauseAction(var aFlag:boolean);
    class procedure DumpRecord(aDs:TDataSet);
    class procedure SaveTxtLog(sMessage:string);
    class procedure CopyFile(SrcFileName, DescFileName : String);//SrcFileName and DestFileName均包含path
    class function GetWinOSVer : Integer;
    class function MessageDlg(Msg, Caption: string; AType: TMsgDlgType;
      AButtons: TMsgDlgButtons): TModalResult; 
  end;


implementation

uses UMsgCodeu ;

{ TUother }

class function TUother.MessageDlg(Msg, Caption: string; AType: TMsgDlgType;
  AButtons: TMsgDlgButtons): TModalResult;
var
  aFrm: TForm;
begin
  aFrm := CreateMessageDialog(Msg, AType, AButtons);
  aFrm.Caption := Caption;
  Result := aFrm.ShowModal;
  aFrm.Free;
end;


class function TUother.CheckIdNo(asId: string; var iSex: TSex): boolean;
begin
  if TellSex(asId, iSex) then
  begin
    if LiteralChk(asId) then  Result := True
    else Result := False;
  end else Result := False;
end;

class function TUother.TellSex(asId: string; var aSex: TSex): Boolean;
begin
  if (asId[2] = '1') or (asId[2] = '2') then
  begin
    if (asId[2] = '1') then aSex := sMale else aSex := sFemale;
    Result := True;
  end else
    Result := False;
end;

class function TUother.LengthChk(asId: string): Boolean;
begin
  if Length(asId) = 10 then  Result := True
  else Result := False;
end;

class function TUother.LiteralChk(asId: string): Boolean;
type
  TFirstChk = set of Char;
var
  FirstChk : TFirstChk;
  letter, ii,j : Integer;
  vp : PChar ;
begin
  letter := 0;
  j := 9;
  vp := @asId[1];
  FirstChk := ['A'..'Z'];
  if LengthChk(asId) then
    if vp^ in FirstChk then
    begin
      case vp^ of
      'A' : letter := 1;
      'B' : letter := 10;
      'C' : letter := 19;
      'D' : letter := 28;
      'E' : letter := 37;
      'F' : letter := 46;
      'G' : letter := 55;
      'H' : letter := 64;
      'I' : letter := 39;
      'J' : letter := 73;
      'K' : letter := 82;
      'L' : letter := 2;
      'M' : letter := 11;
      'N' : letter := 20;
      'O' : letter := 48;
      'P' : letter := 29;
      'Q' : letter := 38;
      'R' : letter := 47;
      'S' : letter := 56;
      'T' : letter := 65;
      'U' : letter := 74;
      'V' : letter := 83;
      'W' : letter := 21;
      'X' : letter := 3;
      'Y' : letter := 12;
      'Z' : letter := 30;
      end;
      for ii := 1 to 9 do
      begin
        try
          Inc(vp);
          Dec(j);
          if j = 0 then
            j := 1;
          letter := letter + StrToInt(vp^)*j;
        except
           Result := False;
           Exit;
        end;
      end;
      letter := letter mod 10;
      if letter = 0 then
        Result := True
      else Result := False;
    end else Result := False
  else Result := False;
end;

class procedure TUother.DumpRecord(aDs: TDataSet);
var
  aStr: string;
  ii: integer;
begin
  aStr := '' ;
  for ii:=0 to aDs.FieldCount-1 do
  begin
    aStr := aStr + aDs.Fields[ii].FieldName + ': ' + aDs.Fields[ii].AsString + #13#10;
  end;
  Application.MessageBox(@aStr[1], '訊息', MB_OKCANCEL + MB_DEFBUTTON1);
end;

class procedure TUother.PauseAction(var aFlag: boolean);
begin
  while True do
  begin
    if not aFlag then Break;
    Application.ProcessMessages;
  end;
end;


class procedure TUother.SaveTxtLog(sMessage: string);
var
  f: TextFile;
begin
  AssignFile(f, 'ErrorLog.txt');
    Append(f);
  Writeln(f, sMessage);
  CloseFile(f);
end;

class procedure TUother.CopyFile(SrcFileName, DescFileName: String);
var
  tmpStrList : TStringList;
begin
  tmpStrList := nil;
  tmpStrList := TStringList.Create;
  tmpStrList.LoadFromFile(SrcFileName);
  tmpStrList.SaveToFile(DescFileName);
  tmpStrList.Free;
  tmpStrList := nil;
end;

class function TUother.GetWinOSVer: Integer;
var
   tmpInfo :OSVersionInfo;
   ret : LongBool;
begin
   Result := -1;
   tmpinfo.dwOSVersionInfoSize := SizeOf(OSVersionInfo);
   ret := Windows.GetVersionEx(tmpInfo);
   if not ret then
     Result := -1 ;//Function failed
   case tmpInfo.dwPlatformId of
    VER_PLATFORM_WIN32_WINDOWS ://(Win95 or Win98...)
      Result := 1;
    VER_PLATFORM_WIN32_NT ://WinNt
      Result := 2;
    VER_PLATFORM_WIN32s ://Others
      Result := 0;
   end;
end;

class function TUother.RoundFloat(Src: Double;
  nRetainPos: Integer): Double;
var
   bL_Minus : Boolean;
   sL_Tmp:String;
   sSrc : String;
   nPos : Integer;
   fTmp_Result, fL_RetainNum : double;
begin
     //ex : Src=258.124, nRetainPos=2
     sSrc := FloatToStr(Src);//sSrc='258.124'

     bL_Minus := False;
     if sSrc[1]='-' then//為負值
     begin
      bL_Minus := True;
      Src := ABS(Src);      
      sSrc := Copy(sSrc,2,Length(sSrc)-1);

     end;
          
     nPos := AnsiPos('.',sSrc);//nPos=4
     if nPos=0 then//找不到'.'
     begin
       if bL_Minus then
         Result := 0 - Src
       else
         Result := Src;

       Exit;
     end;


     sL_Tmp := Copy(sSrc,nPos+1,nRetainPos);//sL_Tmp='12'
     if Length(sL_Tmp)<nRetainPos then
     begin
       if bL_Minus then
         Result := 0 - Src
       else
         Result := Src;
       Exit;
     end;
     sSrc := Copy(sSrc,1,nPos+nRetainPos);//sSrc='258.12'
     if FloatToStr(Src)[nPos+nRetainPos+1]<='4' then//以下捨去
       fTmp_Result := StrToFloat(sSrc)
     else
     begin
       fL_RetainNum := Power(10, (-1*nRetainPos));
       fTmp_Result := StrToFloat(sSrc) + fL_RetainNum;
     end;
     if bL_Minus then
       fTmp_Result := 0 - fTmp_Result;
     Result := fTmp_Result;
end;

class procedure TUother.restoreControlInfo(I_ParentControl: TControl;
  sI_IniFileName: String);
var
    Ini:TiniFile;
    ii: Integer;
    L_TempComp: TComponent;
    sL_OrientationFlag, sL_SectionName, sL_IniName : String;
begin
    if not FileExists(sI_IniFileName) then Exit;
    sL_IniName := sI_IniFileName;
    with I_ParentControl do
    begin
      Ini:=TIniFile.Create(sL_IniName);
      for ii := ComponentCount - 1 downto 0 do
      begin
        L_TempComp := Components[ii];
        if (L_TempComp is TControl) then
        begin
          sL_SectionName:=L_TempComp.Name;
          (L_TempComp as TControl).Top:=Ini.ReadInteger(sL_SectionName,'Top',(L_TempComp as TControl).Top);
          (L_TempComp as TControl).Left:=Ini.ReadInteger(sL_SectionName,'Left',(L_TempComp as TControl).Left);
          (L_TempComp as TControl).Width:=Ini.ReadInteger(sL_SectionName,'Width',(L_TempComp as TControl).Width);
          (L_TempComp as TControl).Height:=Ini.ReadInteger(sL_SectionName,'Height',(L_TempComp as TControl).Height);
        end;
        if (L_TempComp is TQRDBText) then
        begin
          sL_SectionName:=L_TempComp.Name;
          {
          if (L_TempComp as TQRDBText).DataSet<>nil then
            (L_TempComp as TQRDBText).DataSet.Name:=Ini.ReadString(sL_SectionName,'DataSet',(L_TempComp as TQRDBText).DataSet.Name);
          }  
          (L_TempComp as TQRDBText).DataField:=Ini.ReadString(sL_SectionName,'DataField',(L_TempComp as TQRDBText).DataField);
        end;
        if (L_TempComp is TQRExpr) then
        begin
          sL_SectionName:=L_TempComp.Name;
          {
          if (L_TempComp as TQRExpr).Master<>nil then
            (L_TempComp as TQRExpr).Master.Name:=Ini.ReadString(sL_SectionName,'Master',(L_TempComp as TQRExpr).Master.name);
          }              
          (L_TempComp as TQRExpr).Expression:=Ini.ReadString(sL_SectionName,'Expression',(L_TempComp as TQRExpr).Expression);
        end;

        if (L_TempComp is TQRGroup) then
        begin
          sL_SectionName:=L_TempComp.Name;
          {
          if (L_TempComp as TQRGroup).Master<>nil then
            (L_TempComp as TQRGroup).Master.Name:=Ini.ReadString(sL_SectionName,'Master',(L_TempComp as TQRGroup).Master.name);
          }
          (L_TempComp as TQRGroup).Expression:=Ini.ReadString(sL_SectionName,'Expression',(L_TempComp as TQRGroup).Expression);
        end;
        if (L_TempComp is TQRLabel) then
        begin
          sL_SectionName:=L_TempComp.Name;
          (L_TempComp as TQRLabel).Caption :=Ini.ReadString(sL_SectionName,'Caption',(L_TempComp as TQRLabel).Caption);
        end;
      end;
      Ini.Destroy;
    end;
end;

class procedure TUother.saveControlInfo(I_ParentControl: TControl;
  sI_IniFileName: String);
var
    sL_ParentName, sL_ComponentName, sL_FooterBandName, sL_BandType : String;
    Ini:TiniFile;
    sL_AlegnmentFlag : String;
    nL_BandTag,nL_BandHeight, ii: Integer;
    L_TempComp: TComponent;
    sL_ChildBandName, sL_SectionName, sL_IniName, sL_BandInfo : String;
    eL_LeftMargin, eL_RightMargin, eL_BottomMargin, eL_TopMargin : Extended;
    eL_RptWidth, eL_RptLength : Extended;
begin
    sL_IniName := sI_IniFileName;
    with I_ParentControl do
    begin
      Ini:=TIniFile.Create(sL_IniName);
      if I_ParentControl is TQuickRep then
      begin
        sL_SectionName := REPORT_SECTION_NAME;
        eL_LeftMargin := (I_ParentControl as TQuickRep).Page.LeftMargin;
        eL_RightMargin := (I_ParentControl as TQuickRep).Page.RightMargin;
        eL_BottomMargin := (I_ParentControl as TQuickRep).Page.BottomMargin;
        eL_TopMargin := (I_ParentControl as TQuickRep).Page.TopMargin;
        eL_RptWidth := (I_ParentControl as TQuickRep).Page.Width;
        eL_RptLength := (I_ParentControl as TQuickRep).Page.Length;
        Ini.WriteFloat(sL_SectionName,'LM',eL_LeftMargin);
        Ini.WriteFloat(sL_SectionName,'RM',eL_RightMargin);
        Ini.WriteFloat(sL_SectionName,'BM',eL_BottomMargin);
        Ini.WriteFloat(sL_SectionName,'TM',eL_TopMargin);
        Ini.WriteFloat(sL_SectionName,'RPTWIDTH',eL_RptWidth);
        Ini.WriteFloat(sL_SectionName,'RPTLENGTH',eL_RptLength);                                                
      end;
      
      for ii := ComponentCount - 1 downto 0 do
      begin
        sL_SectionName := '';
        L_TempComp := Components[ii];
        if (L_TempComp is TControl) then
        begin
          sL_SectionName:=L_TempComp.Name;
          sL_ParentName := (L_TempComp as TControl).Parent.Name;
          Ini.WriteInteger(sL_SectionName,'Top',(L_TempComp as TControl).Top);
          Ini.WriteInteger(sL_SectionName,'Left',(L_TempComp as TControl).Left);
          Ini.WriteInteger(sL_SectionName,'Width',(L_TempComp as TControl).Width);
          Ini.WriteInteger(sL_SectionName,'Height',(L_TempComp as TControl).Height);
          Ini.WriteString(sL_SectionName,PARENT_TAG,sL_ParentName);
          Ini.WriteString(sL_SectionName,CLASS_NAME_TAG,(L_TempComp as TControl).ClassName);

        end;
        if (L_TempComp is TQRBand) then
        begin
          sL_SectionName := BAND_SECTION_NAME;
          sL_ComponentName :=L_TempComp.Name;
          sL_BandType := getBandFlag(L_TempComp as TQRBand);
          nL_BandTag := (L_TempComp as TQRBand).Tag;
          nL_BandHeight := (L_TempComp as TQRBand).Height;
          sL_FooterBandName := NONE_FOOTER_BAND_FLAG;//無意義,只是為了一致性
          if (L_TempComp as TQRBand).HasChild then
            sL_ChildBandName := (L_TempComp as TQRBand).ChildBand.Name
          else
            sL_ChildBandName := NONE_CHILD_NAME_FLAG;
          
          sL_BandInfo := sL_BandType + ',' + IntToStr(nL_BandTag) + ',' + IntToStr(nL_BandHeight) + ',' + sL_FooterBandName + ',' + sL_ChildBandName;
          Ini.WriteString(sL_SectionName,sL_ComponentName,sL_BandInfo);
        end;
        if (L_TempComp is TQRGroup) then
        begin
          sL_SectionName := BAND_SECTION_NAME;
          sL_ComponentName :=L_TempComp.Name;
          sL_BandType := getBandFlag(L_TempComp as TQRGroup);
          nL_BandTag := (L_TempComp as TQRGroup).Tag;
          nL_BandHeight := (L_TempComp as TQRGroup).Height;
          if (L_TempComp as TQRGroup).FooterBand <> nil then
            sL_FooterBandName := (L_TempComp as TQRGroup).FooterBand.Name
          else
            sL_FooterBandName := '';  

          if (L_TempComp as TQRGroup).HasChild then
            sL_ChildBandName := (L_TempComp as TQRGroup).ChildBand.Name
          else
            sL_ChildBandName := NONE_CHILD_NAME_FLAG;

          sL_BandInfo := sL_BandType + ',' + IntToStr(nL_BandTag) + ',' + IntToStr(nL_BandHeight) + ',' + sL_FooterBandName + ',' + sL_ChildBandName;
          Ini.WriteString(sL_SectionName,sL_ComponentName,sL_BandInfo);
        end;

        if (L_TempComp is TQRChildBand) then
        begin
          sL_SectionName := BAND_SECTION_NAME;
          sL_ComponentName :=L_TempComp.Name;
          sL_BandType := getBandFlag(L_TempComp as TQRChildBand);
          nL_BandTag := (L_TempComp as TQRChildBand).Tag;
          nL_BandHeight := (L_TempComp as TQRChildBand).Height;
          sL_FooterBandName := NONE_FOOTER_BAND_FLAG;//無意義,只是為了一致性
          if (L_TempComp as TQRChildBand).HasChild then
            sL_ChildBandName := (L_TempComp as TQRChildBand).ChildBand.Name
          else
            sL_ChildBandName := NONE_CHILD_NAME_FLAG;  
          sL_BandInfo := sL_BandType + ',' + IntToStr(nL_BandTag) + ',' + IntToStr(nL_BandHeight) + ',' + sL_FooterBandName +',' + sL_ChildBandName;
          Ini.WriteString(sL_SectionName,sL_ComponentName,sL_BandInfo);
        end;
        
        if (L_TempComp is TQRDBText) then
        begin
          sL_SectionName:=L_TempComp.Name;
          if (L_TempComp as TQRDBText).DataSet<>nil then
            Ini.WriteString(sL_SectionName,'DataSet',(L_TempComp as TQRDBText).DataSet.Name);
          Ini.WriteString(sL_SectionName,'DataField',(L_TempComp as TQRDBText).DataField);
          sL_AlegnmentFlag := getAlignmentFlag((L_TempComp as TQRDBText).Alignment);
          Ini.WriteString(sL_SectionName,'Alignment',sL_AlegnmentFlag);
          
        end;
        if (L_TempComp is TQRExpr)  then
        begin
          sL_SectionName:=L_TempComp.Name;
          if (L_TempComp as TQRExpr).Master<> nil then
            Ini.WriteString(sL_SectionName,'Master',(L_TempComp as TQRExpr).Master.Name);
          Ini.WriteString(sL_SectionName,'Expression',(L_TempComp as TQRExpr).Expression);
          sL_AlegnmentFlag := getAlignmentFlag((L_TempComp as TQRExpr).Alignment);
          Ini.WriteString(sL_SectionName,'Alignment',sL_AlegnmentFlag);

        end;
        if (L_TempComp is TQRGroup) then
        begin
          sL_SectionName:=L_TempComp.Name;
          if (L_TempComp as TQRGroup).Master<> nil then
            Ini.WriteString(sL_SectionName,'Master',(L_TempComp as TQRGroup).Master.Name);
          Ini.WriteString(sL_SectionName,'Expression',(L_TempComp as TQRGroup).Expression);
        end;

        if (L_TempComp is TQRLabel) then
        begin
          sL_SectionName:=L_TempComp.Name;
          Ini.WriteString(sL_SectionName,'Caption',(L_TempComp as TQRLabel).Caption);
          sL_AlegnmentFlag := getAlignmentFlag((L_TempComp as TQRLabel).Alignment);
          Ini.WriteString(sL_SectionName,'Alignment',sL_AlegnmentFlag);          
        end;


      end;
      Ini.Destroy;
    end;
end;

class function TUother.RunShell(sI_ExeName, sI_Param: string;
  bI_WithUI, bI_WaitProcess: Boolean; var sI_ErrorMsg:String): Boolean;
var
  pi: TProcessInformation;
  si: TStartupInfo;
  bRet: Boolean;
  abc :  Cardinal;
  _Bool : BOOL;
  PMsgBuf : PChar;

begin
    sI_ErrorMsg := '';
    Result := False;

    FillChar(si, SizeOf(TStartupInfo), 0);
    si.cb := SizeOf(TStartupInfo);
    si.wShowWindow := SW_NORMAL;

    bRet := CreateProcess(nil, PChar(sI_ExeName + ' ' + sI_Param), nil, nil, False,
      0, nil, nil, si, pi);

      _Bool := GetExitCodeProcess(pi.hProcess,abc);

  //    TerminateProcess(pi.hProcess,abc);  //mark by Howard 0316
    if not bRet then
    begin
      pMsgBuf := nil ;
      FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER or FORMAT_MESSAGE_FROM_SYSTEM,
        nil,
        GetLastError(),
        LOCALE_USER_DEFAULT,
        PMsgBuf,
        0,
        nil );
  //      showmessage(PMsgBuf);
        sI_ErrorMsg := PMsgBuf;
        LocalFree(HLOCAL(PMsgBuf));

        Result := false;
        Exit;
    end;
  //  if bI_WithUI then Hide;   //mark by Howard 0329

    { wait until the process terminate}
    if bI_WaitProcess then
      WaitForSingleObject(pi.hProcess, INFINITE);

  //  if bI_WithUI then Show;

    Result := True ;
end;

class function TUother.getBandFlag(I_QRBand: TQRCustomBand): String;
var
    sL_BandFlag : String;
begin
  case I_QRBand.BandType of
    rbChild :
     sL_BandFlag := '1';
    rbColumnHeader :
     sL_BandFlag := '2';
    rbDetail :
     sL_BandFlag := '3';
    rbGroupFooter :
     sL_BandFlag := '4';
    rbGroupHeader :
     sL_BandFlag := '5';
    rbOverlay :
     sL_BandFlag := '6';
    rbPageFooter :
     sL_BandFlag := '7';
    rbPageHeader :
     sL_BandFlag := '8';
    rbSubDetail :
     sL_BandFlag := '9';
    rbSummary :
     sL_BandFlag := '10';
    rbTitle :
     sL_BandFlag := '11';
  end;
  result := sL_BandFlag;
end;

class function TUother.getRealBandType(
  sI_BandTypeFlag: String): TQrBandType;
var
    L_BandType : TQrBandType;
begin
    if sI_BandTypeFlag='1' then
      L_BandType := rbChild
    else if sI_BandTypeFlag='2' then
      L_BandType := rbColumnHeader
    else if sI_BandTypeFlag='3' then
      L_BandType := rbDetail
    else if sI_BandTypeFlag='4' then
      L_BandType := rbGroupFooter
    else if sI_BandTypeFlag='5' then
      L_BandType := rbGroupHeader
    else if sI_BandTypeFlag='6' then
      L_BandType := rbOverlay
    else if sI_BandTypeFlag='7' then
      L_BandType := rbPageFooter
    else if sI_BandTypeFlag='8' then
      L_BandType := rbPageHeader
    else if sI_BandTypeFlag='9' then
      L_BandType := rbSubDetail
    else if sI_BandTypeFlag='10' then
      L_BandType := rbSummary
    else if sI_BandTypeFlag='11' then
      L_BandType := rbTitle;
  result := L_BandType;
end;


class function TUother.CreateQRControl(QRControlClass: TControlClass;
  aParent: TControl; I_QuickRep:TQuickRep): TControl;
var
  aCtrl : TControl ;
begin
  aCtrl := QRControlClass.Create(nil) ;
  TQRPrintable(aCtrl).ParentReport := I_QuickRep ;
  TWinControl(aParent).InsertControl(aCtrl);
  aCtrl.Visible := True ;
  Result := aCtrl ;
end ;




class function TUother.getAlignmentFlag(I_Alignment: TAlignment): String;
var
    sL_Result : String;
begin
    case I_Alignment of
      taLeftJustify :
        sL_Result := ALIGNMENT_LEFT_FLAG;
      taCenter :
        sL_Result := ALIGNMENT_CENTER_FLAG;
      taRightJustify :
        sL_Result := ALIGNMENT_RIGHT_FLAG;
    end;
    result := sL_Result;
end;

class function TUother.getRealAlignmentValue(
  sI_AlignmentFlag: String): TAlignment;
var
    L_Alignment : TAlignment;
begin
    if sI_AlignmentFlag = ALIGNMENT_CENTER_FLAG then
      L_Alignment := taCenter
    else if sI_AlignmentFlag = ALIGNMENT_RIGHT_FLAG then
      L_Alignment := taRightJustify
    else
      L_Alignment := taLeftJustify;
      
    result := L_Alignment;
end;

class function TUother.getLocalIp: String;
type
    TaPInAddr = array [0..10] of PInAddr;
    PaPInAddr = ^TaPInAddr;
var
    phe  : PHostEnt;
    pptr : PaPInAddr;
    Buffer : array [0..63] of char;
    I    : Integer;
    GInitData      : TWSADATA;

begin
    WSAStartup($101, GInitData);
    Result := '';
    GetHostName(Buffer, SizeOf(Buffer));
    phe :=GetHostByName(buffer);
    if phe = nil then Exit;
    pptr := PaPInAddr(Phe^.h_addr_list);
    I := 0;
    while pptr^[I] <> nil do begin
      result:=StrPas(inet_ntoa(pptr^[I]^));
      Inc(I);
    end;
    WSACleanup;
end;

class function TUother.regRead(sI_KeyPath, sI_KeyNam: string;
  nI_KeyStyle: integer): Variant;
var
    L_Registry: TRegistry;
begin
	L_Registry:=TRegistry.Create;
    L_Registry.RootKey:=HKEY_LOCAL_MACHINE;
    //  （預設位置）Default RootKey:= HKEY_CURRENT_USER;
    //將資料讀取出來，如果沒有該目錄則建立新目錄
    try
        L_Registry.OpenKey(sI_KeyPath,True);
      //將上次的資料讀取出來
      case nI_KeyStyle of
      1: {數值}
        Result:=L_Registry.ReadInteger(sI_KeyNam);
      2: {字串}
        Result:=L_Registry.ReadString(sI_KeyNam);
      3: {日期}
        Result:=L_Registry.ReadDate(sI_KeyNam);
      end;
    except
      showmessage('Error:Registry Read Error');
      Result:='?';
    end;
    L_Registry.Free;
end;

class function TUother.regWrite(sI_KeyPath, sI_KeyNam: string;
  nI_KeyStyle: integer; vI_KeyVal: Variant): boolean;
var
    L_Registry: TRegistry;
begin
	L_Registry:=TRegistry.Create;
    L_Registry.RootKey:=HKEY_LOCAL_MACHINE;
 	//  （預設位置）Default RootKey:= HKEY_CURRENT_USER;
    //將資料讀取出來，如果沒有該目錄則建立新目錄
    try
      L_Registry.OpenKey(sI_KeyPath,True);
          //將目前視窗的資訊寫到系統的登錄資料中
      case nI_KeyStyle of
        1: {數值}
              L_Registry.WriteInteger(sI_KeyNam,vI_KeyVal);
        2: {字串}
            L_Registry.WriteString(sI_KeyNam,vI_KeyVal);
        3: {日期}
            L_Registry.WriteDate(sI_KeyNam,vI_KeyVal);
      end;
      Result:=True;
    except
      Result:=False;
    end;
    L_Registry.Free;
end;

class function TUother.myDecrypt(sI_SourceStr, sI_Key: String): String;
var
    ii, nL_KeyPos, nL_SourceLength, nL_KeyLength : Integer;
    sL_Reault : String;
begin
  nL_KeyPos:=1; // initialize count
  nL_SourceLength := length(sI_SourceStr);
  nL_KeyLength := Length(sI_Key);

  for ii:=1 to nL_SourceLength do
  begin
    sL_Reault := sL_Reault + Chr(Ord(sI_SourceStr[ii])-Ord(sI_Key[nL_KeyPos]));
    Inc(nL_KeyPos);
    if nL_KeyPos=(nL_KeyLength)then nL_KeyPos:=1;
  end;
  result := sL_Reault;

end;

class function TUother.myEncrypt(sI_SourceStr, sI_Key: String): String;
var
    ii, nL_KeyPos, nL_SourceLength, nL_KeyLength : Integer;
    sL_Reault : String;

begin
  nL_KeyPos:=1; // initialize count
  nL_KeyLength := length(sI_Key);
  nL_SourceLength := length(sI_SourceStr);
  
  for ii:=1 to nL_SourceLength do
  begin
    sL_Reault := sL_Reault + Chr(Ord(sI_SourceStr[ii])+Ord(sI_Key[nL_KeyPos]));
    Inc(nL_KeyPos);
    if nL_KeyPos=(nL_KeyLength)then nL_KeyPos:=1;
  end;
  result := sL_Reault;
end;

end.


