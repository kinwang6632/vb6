unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls,  qrextra, QRExport,
  qrprntr, Quickrpt;

type
  TfrmMain = class(TForm)
    BitBtn1: TBitBtn;
    rgpOption: TRadioGroup;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure transReportFormat(I_QuickRep : TQuickRep; sI_ReportHeader:String; dI_SDate, dI_EDate:TDate);
    procedure transReport1Format(I_DataStrList : TStringList; sI_ReportHeader:String; dI_SDate, dI_EDate:TDate);    
  end;

var
  frmMain: TfrmMain;

implementation

uses frmReport1U, frmReport2U;

{$R *.DFM}

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
    case rgpOption.ItemIndex of
      0:
       begin
        frmReport1 := TfrmReport1.Create(Application);
        frmReport1.ShowModal;
        frmReport1.free;
       end;
      1:
       begin
        frmReport2 := TfrmReport2.Create(Application);
        frmReport2.ShowModal;
        frmReport2.free;
       end;

    end;
end;

procedure TfrmMain.transReport1Format(I_DataStrList: TStringList;
  sI_ReportHeader: String; dI_SDate, dI_EDate: TDate);
var

  wL_Year, wL_Month, wL_Day : Word;
  sL_SDate, sL_EDate : String;
  sL_RptPath , sL_RptFileName : String;
  L_StrList : TStringList;
  ii  :  Integer;

begin
   if I_DataStrList=nil  then  Exit;
   if  I_DataStrList.Count<=0 then Exit;
           
   decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
   sL_SDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);

   decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
   sL_EDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);


   sL_RptPath := 'C:\Documents and Settings\seccbackup\орн▒\';

   sL_RptFileName := sL_RptPath + sI_ReportHeader + sL_SDate + '~' + sL_EDate + '.txt';


   try
    L_StrList := TStringList.Create;
    for  ii:=0  to  I_DataStrList.Count-1 do
    begin

      L_StrList.Add((I_DataStrList.Objects[ii] as TReport1Obj).FullDataString);

    end;
    L_StrList.SaveToFile(sL_RptFileName);
    L_StrList.free;
   finally

   end;


end;

procedure TfrmMain.transReportFormat(I_QuickRep: TQuickRep;
  sI_ReportHeader: String; dI_SDate, dI_EDate: TDate);
var
  AsciiExportFilter : TQRAsciiExportFilter;
  wL_Year, wL_Month, wL_Day : Word;
  sL_SDate, sL_EDate : String;
  sL_RptPath , sL_RptFileName : String;

begin

   decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
   sL_SDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);

   decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
   sL_EDate := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);

   sL_RptPath := 'C:\Documents and Settings\seccbackup\орн▒\';
   //sL_RptPath := 'C:\';

   sL_RptFileName := sL_RptPath + sI_ReportHeader + sL_SDate + '~' + sL_EDate + '.txt';

   AsciiExportFilter := TQRAsciiExportFilter.Create(sL_RptFileName);
   try
      I_QuickRep.ExportToFilter(AsciiExportFilter);
   finally
      AsciiExportFilter.Free;
   end;

{
Other filters:
HTML: TQRHTMLDocumentFilter
ASCII: TQRAsciiExportFilter
CSV: TQRCommaSeparatedFilter

In Professional Version:
RTF: TQRRTFExportFilter
WMF: TQRWMFExportFilter
Excel: TQRXLSFilter
}

end;

end.
