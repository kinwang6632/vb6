unit frmReport1U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, ExtCtrls, Db, Grids, DBGrids, IniFiles;

const
    ACTIVITY_CODE_HEADER = 'Activity-Code-';
    SPECIAL_DATA_ID  ='999';
    
type

  TReport1Obj  =  class(TObject)
      FullDataString  :  String;

  end;

  TfrmReport1 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label1: TLabel;
    dtpRptSDate: TDateTimePicker;
    rgpComp: TRadioGroup;
    chbShowError: TCheckBox;
    Label2: TLabel;
    Label3: TLabel;
    dtpRptEDate6: TDateTimePicker;
    rgpDataSrc: TRadioGroup;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
    procedure LoadAppInfo;
  public
    { Public declarations }
    G_Report1StrList :TStringList;
  end;

var
  frmReport1: TfrmReport1;

implementation

uses dtmMainU, rptReport1U, frmMainU;

{$R *.DFM}

procedure TfrmReport1.BitBtn1Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmReport1.FormShow(Sender: TObject);
begin
    dtpRptSDate.date := date;
    dtpRptEDate6.date := date;
    LoadAppInfo;
end;

procedure TfrmReport1.BitBtn2Click(Sender: TObject);
var
    nL_CompNdx : Integer;
    bL_HasRecord : boolean;
    nL_NtyTOccurrences, nL_LscTOccurrences,   nL_BestLscTOccurrences : Integer;
    nL_CyLscTOccurrences,  nL_ShLscTOccurrences, nL_BestTOccurrences,nL_ShTOccurrences,nL_CyTOccurrences : Integer;
    nL_DataSrc : Integer;

begin
    { 舊的
    Application  ID
                        Stanley
    NTY                 10004
    BEST                10015
    SH                  10003
    CY                  10011
    LSC                 10007
    BEST_LSC            10012
    SH_LSC              10000
    CY_LSC              10008

    SERVICE_NTY         10005
    SERVICE_BEST        10014
    SERVICE_SH          10002
    SERVICE_CY          10010
    }

    { 新的
    NTY + LSC + SERVICE_NTY           10004 + 10007 + 10005
    BEST + BEST_LSC + SERVICE_BEST    10015 + 10012 + 10014
    SH + SH_LSC + SERVICE_SH          10003 + 10000 + 10002
    CY + CY_LSC + SERVICE_CY          10011 + 10008 + 10010

    LSC                 10007
    BEST_LSC            10012
    SH_LSC              10000
    CY_LSC              10008

    SERVICE_NTY         10005
    SERVICE_BEST        10014
    SERVICE_SH          10002
    SERVICE_CY          10010

    }

    nL_NtyTOccurrences := 0;
    nL_LscTOccurrences := 0;
    nL_CompNdx := rgpComp.ItemIndex;
    nL_DataSrc := rgpDataSrc.ItemIndex ;
    bL_HasRecord := dtmMain.activeActivityCodeStat(nL_DataSrc, nL_CompNdx, dtpRptSDate.date, dtpRptEDate6.date, nL_NtyTOccurrences, nL_BestTOccurrences,nL_ShTOccurrences,nL_CyTOccurrences);
//    bL_HasRecord := dtmMain.activeActivityCodeStat(nL_CompNdx, dtpRptSDate.date, nL_NtyTOccurrences, nL_LscTOccurrences);
    if (bL_HasRecord) then
    begin
      rptReport1 := TrptReport1.Create(Application);
      rptReport1.bG_ShowInvalidDetail := chbShowError.Checked;

      rptReport1.sG_RptSDate := DateToStr(dtpRptSDate.date);
      rptReport1.sG_RptEDate := DateToStr(dtpRptEDate6.date);      
      rptReport1.nG_GroupOccurrance :=0;
      rptReport1.fG_GroupPercent :=0;

      rptReport1.nG_CompNdx := nL_CompNdx;
      rptReport1.nG_NtyTOccurrences := nL_NtyTOccurrences;
      rptReport1.nG_BestTOccurrences := nL_BestTOccurrences;
      rptReport1.nG_ShTOccurrences :=   nL_ShTOccurrences;
      rptReport1.nG_CyTOccurrences := nL_CyTOccurrences;

      rptReport1.nG_LscTOccurrences := nL_LscTOccurrences;
      rptReport1.nG_BestLscTOccurrences := nL_BestLscTOccurrences;
      rptReport1.nG_ShLscTOccurrences := nL_ShLscTOccurrences;
      rptReport1.nG_CyLscTOccurrences := nL_CyLscTOccurrences;      

      G_Report1StrList :=  TStringList.Create;
      rptReport1.Preview;
//      frmMain.transReportFormat(rptReport1, ACTIVITY_CODE_HEADER,dtpRptSDate.date, dtpRptEDate6.date);  //用此方式轉出的文字檔資料會有遺漏=>是quick  report  元件的問題


      frmMain.transReport1Format(G_Report1StrList, ACTIVITY_CODE_HEADER,dtpRptSDate.date, dtpRptEDate6.date);  
      G_Report1StrList.Free;

      rptReport1.free;
    end
    else
      MessageDlg('沒有資料可供列印!', mtInformation,[mbOK], 0);
end;

procedure TfrmReport1.LoadAppInfo;
var
  aFullName: String;
  aIniFile: TIniFile;
  aList: TStringList;
  aIndex, aPos: Integer;
  aObj: PAppInfo;
  aName, aValue: String;
begin
  aFullName := IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + APPINFO_FILENAME;
  aIniFile := TIniFile.Create( aFullName );
  try
    aList := TStringList.Create;
    try
      aIniFile.ReadSectionValues( 'ApplicationId', aList );
      if ( aList.Count = 0 ) then
      begin
        { NTY }
        New( aObj );
        aObj.AppId := '10004';
        aObj.AppName := 'NTY';
        aObj.SubAppId := '10004,10007,10005';
        AppInfoList.Add( aObj );
        { BEST }
        New( aObj );
        aObj.AppId := '10015';
        aObj.AppName := 'NTY';
        aObj.SubAppId := '10015,10012,10014';
        AppInfoList.Add( aObj );
        { SH }
        New( aObj );
        aObj.AppId := '10003';
        aObj.AppName := 'NTY';
        aObj.SubAppId := '10003,10000,10002';
        AppInfoList.Add( aObj );
        { CY }
        New( aObj );
        aObj.AppId := '10011';
        aObj.AppName := 'NTY';
        aObj.SubAppId := '10011,10008,10010';
        AppInfoList.Add( aObj );
        dtmMain.SaveAppObjToFile;
      end else
      begin
        for aIndex := 0 to aList.Count - 1 do
        begin
          aPos := Pos( '=', aList[aIndex] );
          aName := Copy( aList[aIndex], 1, aPos - 1 );
          aValue := Copy( aList[aIndex], aPos + 1 , Length( aList[aIndex] ) - aPos );
          New( aObj );
          aObj.AppId := aName;
          aObj.SubAppId := aValue;
          AppInfoList.Add( aObj );
        end;
      end;
    finally
      aList.Free;
    end;
  finally
    aIniFile.Free;
  end;
end;

end.
