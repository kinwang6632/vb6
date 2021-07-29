{--------------------------------------------------------------------------------
* Description : TscExcelExport : Export Delphi Dataset to MS Excel
* Dates : February 2000 - June 2001
* Version : 2.4 (Delphi 5.0 and 6.0)

* Author : Stefan Cruysberghs
* Email : stefancr@yucom.be
* Website : http://www.stefancr.yucom.be
--------------------------------------------------------------------------------
* $Archive: /Component Library/SC/scExcelExport.pas $
* $Author: Administrator $
* $Date: 12/05/01 11:12 $
* $Modtime: 12/05/01 11:06 $
* $Revision: 7 $
--------------------------------------------------------------------------------
This component is free of charge.
The author doesn't give a warranty for error free running
of this component and he doesn't give any support.
Suggestions and bugs can be send by email.
--------------------------------------------------------------------------------
Version 1.2
  - Improved connection to Excel
Version 1.3
  - Added Orientation of titles
  - Added StyleColumnWidth
Version 1.4
  - Improved GetColumnCharacters
  - Added Border properties and background colors for Titles, Data and Summary
  - Added Summary properties (SummarySelection, SummaryFields, SummaryCalculation)
Version 1.5
  - Improved speed of exporting data
  - Improved exporting string-fields
  - Added ConnectTo property
Version 1.6
  - Suppression of reference to Dataset.RecordCount (thanks to Gérard Beneteau)
  - Added VisibleFieldsOnly property
Version 1.7
  - Notification when disconnecting dataset
  - Very fast export by using a variant matrix (thanks to Frank van den Bergh)
Version 1.8
  - Bug in exporting titles (thanks to Roberto Parola)
  - Setting format of cells before exporting data (thanks to Asbjørn Heggvik)
  - Added BlockOfRecords property : little bit faster and more control of memory
  - Added properties to set begin row of titles and data
Version 1.9
  - Added properties (Orientation, WrapText, Alignment) to font (thanks to David Balick)
    (property OrientationTitles is removed)
  - Added HeaderText (thanks to David Balick)
  - Improved some routines
Version 2.0
  - Added read only property with row number of last data row
  - Added property with the Excel tabsheet so after the export
    it is possible to access the cells in Excel
  - Added event OnExportRecords
Version 2.1
  - Bugfixes
Version 2.2
  - Bugfix when exporting filtered dataset (thanks to Heinz Faerber)
  - New column width styles : cwFieldDisplayWidth, cwFieldDataSize, cwEnhAutoFit (idea from Scott Stanek)
Version 2.3
  - Added properties begin column data/titles and header (idea from Knjazev Sergey)
  - Added property ShowTitles
Version 2.4
  - Support for Delphi 6.0
-------------------------------------------------------------------------------}

{
  Example 1 : easiest way to export a dataset to Excel

  scExcelExport1.Dataset:=Table1;
  scExcelExport1.ExportDataset;
  scExcelExport1.Disconnect;
}

{
  Example 2 : use layout properties, add summary cells and save file

  scExcelExport1.WorksheetName := 'MyDataset';
  scExcelExport1.Dataset:=Table1;
  scExcelExport1.StyleColumnWidth:=cwOwnerWidth;
  scExcelExport1.ColumnWidth := 20;
  scExcelExport1.HeaderText.Text := 'Header';
  scExcelExport1.BeginRowHeader := 2;
  scExcelExport1.FontTitles := LabelTitle.Font;
  scExcelExport1.FontTitles.Orientation := 45;
  scExcelExport1.BorderTitles.BackColor := clYellow;
  scExcelExport1.BorderTitles.BorderColor := clRed;
  scExcelExport1.BorderTitles.LineStyle := blLine;
  scExcelExport1.BeginRowTitles := 5;
  scExcelExport1.BeginColumnData := 3;
  scExcelExport1.FontData := LabelData.Font;
  scExcelExport1.SummarySelection := ssValues;
  scExcelExport1.SummaryCalculation := scMAX;
  scExcelExport1.ExcelVisible:=False;
  try
    scExcelExport1.ExportDataset;

    if Assigned(scExcelExport1.ExcelWorkSheet) then
      scExcelExport1.ExcelWorkSheet.Range['A1','A10'].Value := 'Delphi';

    scExcelExport1.SaveAs('c:\test.xls',ffXLS);
  finally
    scExcelExport1.Disconnect;
  end;
}

{
  Example 3 : export more datasets 

  scExcelExport1.ExcelVisible:=True;

  try
    scExcelExport1.Dataset:=Table1;
    scExcelExport1.WorksheetName:='1';
    scExcelExport1.ConnectTo := ctNewExcel;
    scExcelExport1.ExportDataset;

    scExcelExport1.Dataset:=Table2;
    scExcelExport1.WorksheetName:='2';
    scExcelExport1.ConnectTo := ctNewWorkbook;
    scExcelExport1.ExportDataset;

    scExcelExport1.Dataset:=Table3;
    scExcelExport1.WorksheetName:='3';
    scExcelExport1.ConnectTo := ctNewWorksheet;
    scExcelExport1.ExportDataset;
  finally
    scExcelExport1.Disconnect;
  end;

}

unit scExcelExport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, DB,
  StdCtrls, Variants, 
  OleServer, // used for TConnectKind
  {$IFDEF VER140}
  // Delphi 11
    Excel2000,
    Variants;
  {$ELSEIFDEF VER185}
  // Delphi 6
    Excel2000,
    Variants;
  {$ELSE}
  // Delphi 5
    Excel2000; // Excel97 or Excel2000
  {$ENDIF}

  // Delphi 5.0 notes
  // ----------------
  // Delphi5 standard installation installs Excel97
  // If you wants to use Excel2000 you need to apply the following modifications
  // described in upd1rdme.txt coming with Delphi Update Pack 1
  //
  //   Office 2000 Components
  //   ----------------------
  //   To install the Office 2000 components package:
  //
  //   1. Select File | Close all
  //   2. Select Component | Install package
  //   3. Remove package DclAxServer50.bpl to avoid
  //      name conflicts with Ofice 97 components.
  //   4. Add package Dcloffice2k50.bpl, which contains
  //      the Office 2000 components.

type
  TFileFormat = (ffXLS, ffHTM, ffCSV, ffXL97);
  TStyleColumnWidth = (cwDefault, cwOwnerWidth, cwAutoFit, cwFieldDisplayWidth, cwFieldDataSize, cwEnhAutoFit);
    // cwOwnerWidth : width specified with property ColumnWidth
    // cwAutoFit : Excel autofit
    // cwFieldDisplayWidth : width of DisplayWidth of TField
    // cwFieldDataSize : width of Datasize of TField
    //    Datasize = amount of memory to store value, for datetime fields width is set to 16
    // cwEnhAutoFit (enhanced autofit) : width of DisplayWidth of TField except when title is larger
  TBorderWeight = (bwThin, bwMedium, bwThick);
    // xlHairline, xlMedium, xlThick, xlThin
  TBorderLineStyle = (blNone, blDash, blLine, blDoubleLine);
    // xlContinuous, xlDash, xlDashDot, xlDashDotDot, xlDot, xlDouble, xlSlantDashDot, xlLineStyleNone
  TSummarySelection = (ssNone, ssValues, ssGiven);
  TSummaryCalculation = (scSUM, scMIN, scMAX);
  THAlignment = (haGeneral, haLeft, haRight, haCenter);
  TConnectTo = (ctNewExcel, ctNewWorkbook, ctNewWorksheet);

  TxlFont = class(TFont)
  private
    FAlignment: THAlignment;
    FBlnWrapText: Boolean;
    FIntOrientation: Integer;
  published
    property Alignment: THAlignment read FAlignment write FAlignment;
    property WrapText: Boolean read FBlnWrapText write FBlnWrapText;
    property Orientation: Integer read FIntOrientation write FIntOrientation;
  end;

  TCellBorder = class(TPersistent)
  private
    FBackColor : TColor;
    FBorderColor : TColor;
    FBorderWeight : TBorderWeight;
    FBorderLineStyle : TBorderLineStyle;
  published
    property BackColor : TColor read FBackColor write FBackColor default clWhite;
    property BorderColor : TColor read FBorderColor write FBorderColor default clBlack;
    property Weight : TBorderWeight read FBorderWeight write FBorderWeight default bwMedium;
    property LineStyle : TBorderLineStyle read FBorderLineStyle write FBorderLineStyle default blNone;
  end;

  TOnExportEvent = procedure(Sender : TObject; IntRecordNumber : Integer) of object;

  TscExcelExport = class(TComponent)
  private
    FDataset : TDataset;
    FIntRecordNo : integer;
    FIntTotRecNo: Integer;
    FIntBeginRowHeader : Integer;
    FIntBeginRowTitles : Integer;
    FIntBeginRowData : Integer;
    FIntBeginColumnData : Integer;
    FIntBeginColumnHeader : Integer;

    FIntEndRowData : Integer;

    FExcelApplication : TExcelApplication;
    FExcelWorkbook : TExcelWorkbook;
    FExcelWorksheet : TExcelWorksheet;

    FFieldNames : TStrings;

    FBlnShowTitles : Boolean;
    FBlnExcelVisible : Boolean;
    FStrWorksheetName : String;
    FIntColumnWidth : Integer;
    FStyleColumnWidth : TStyleColumnWidth;
    FConnectTo : TConnectTo;

    FFontHeader: TxlFont;
    FBorderHeader : TCellBorder;
    FStrHeaderText: TStrings;

    FFontTitles : TxlFont;
    FBorderTitles : TCellBorder;

    FFontData : TxlFont;
    FBorderData : TCellBorder;

    FFontSummary : TxlFont;
    FBorderSummary : TCellBorder;
    FSummarySelection : TSummarySelection;
    FSummaryFields : TStrings;
    FSummaryCalculation : TSummaryCalculation;

    FIntBlockOfRecords : Integer;

    FStrBeginColumnDataChar : String;

    LCID : Integer;
    FVisibleFieldsOnly: Boolean;

    FOnExportRecords : TOnExportEvent;
    procedure SetBeginColumnData(const Value: Integer);
    procedure SetBeginColumnHeader(const Value: Integer);
  protected
    procedure SetFontHeader(const Value: TxlFont);
    procedure SetHeaderText(const Value: TStrings);
    procedure SetFontTitles(Value : TxlFont);
    procedure SetFontData(Value : TxlFont);
    procedure SetFontSummary(Value : TxlFont);
    procedure SetSummaryFields(Value : TStrings);
    procedure SetVisibleFieldsOnly(const Value: Boolean);
    procedure SetBeginRowHeader(const Value: Integer);
    procedure SetBeginRowTitles(const Value: Integer);
    procedure SetBeginRowData(const Value: Integer);

    procedure SetColumnWidth;
    function GetColumnCharacters(IntNumber : Integer) : String;
    procedure SetFontAndBorderRange(DelphiFont : TxlFont; Border : TCellBorder; StrBeginCell, StrEndCell : String);
    procedure SetFormat;

    function CanConvertFieldToCell(Field : TField) : Boolean;
    function IsValueField(Field : TField) : Boolean;
    function GetWidthFromDatasize(Field : TField) : Integer;

    procedure ExportHeader;
    procedure ExportTitles;
    procedure ExportFieldData;
    procedure ExportSummary;
  public
    constructor Create(Owner: TComponent); override;
    destructor Destroy; override;

    // Read-only properties
    // Line number of last row of data is filled in after the export
    property EndRowData : Integer read FIntEndRowData;
    // Link to the Excel worksheet, can be used to add extra data after the export
    property ExcelWorkSheet : TExcelWorksheet read FExcelWorksheet write FExcelWorksheet;

    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    procedure Disconnect; virtual;
    procedure ExportDataset(BlnOpenedExcel : Boolean = False); virtual;
    procedure SaveAs(const StrFileName : String; const FileFormat : TFileFormat); virtual;
    procedure PrintPreview(const BlnPrintGridLines : Boolean); virtual;
  published
    // Show or hide excel (default True)
    property ExcelVisible : Boolean read FBlnExcelVisible write FBlnExcelVisible default True;
    // New instance of Excel application, new workbook or new worksheet (default is new instance of Excel)
    property ConnectTo : TConnectTo read FConnectTo write FConnectTo default ctNewExcel;

    // Name of worksheet
    property WorksheetName : String read FStrWorksheetName write FStrWorksheetName;
    // Dataset which will be exported (TTable, TQuery, TClientDataset, TADODataset, ...)
    property Dataset : TDataset read FDataset write FDataset;

    // Style of columnswidth : Excel default, width of property ColumnWidth, AutoFit
    property StyleColumnWidth : TStyleColumnWidth read FStyleColumnWidth write FStyleColumnWidth;
    property ColumnWidth : Integer read FIntColumnWidth write FIntColumnWidth;

    // Export only visible fields or all fields
    property VisibleFieldsOnly : Boolean read FVisibleFieldsOnly write SetVisibleFieldsOnly default True;

    // Font and border of header
    // Fill in header texts
    property FontHeader: TxlFont read FFontHeader write SetFontHeader;
    property HeaderText: TStrings read FStrHeaderText write SetHeaderText;
    property BorderHeader : TCellBorder read FBorderHeader write FBorderHeader;

    // Font and border of titles
    property ShowTitles : Boolean read FBlnShowTitles write FBlnShowTitles default True;
    property FontTitles : TxlFont read FFontTitles write SetFontTitles;
    property BorderTitles : TCellBorder read FBorderTitles write FBorderTitles;

    // Font and border of data
    property FontData : TxlFont read FFontData write SetFontData;
    property BorderData : TCellBorder read FBorderData write FBorderData;

    // Font and border of summary
    property FontSummary : TxlFont read FFontSummary write SetFontSummary;
    property BorderSummary : TCellBorder read FBorderSummary write FBorderSummary;
    // Which fields will be summerized : all numeric fields, the fields of SummaryFields, none
    property SummarySelection : TSummarySelection read FSummarySelection write FSummarySelection;
    property SummaryFields : TStrings read FSummaryFields write SetSummaryFields;
    // Calculation : SUM, MIN, MAX
    property SummaryCalculation : TSummaryCalculation read FSummaryCalculation write FSummaryCalculation;

    // Number of records which will be exported in one variant matrix (default 20)
    // Try to increase and decrease this property for the optimal speed
    property BlockOfRecords : Integer read FIntBlockOfRecords write FIntBlockOfRecords default 20;

    // Begin row of titles and data
    property BeginRowHeader : Integer read FIntBeginRowHeader write SetBeginRowHeader default 1;
    property BeginRowTitles : Integer read FIntBeginRowTitles write SetBeginRowTitles default 1;
    property BeginRowData : Integer read FIntBeginRowData write SetBeginRowData default 2;

    // Begin column header and data/titles
    property BeginColumnHeader : Integer read FIntBeginColumnHeader write SetBeginColumnHeader default 1;
    property BeginColumnData : Integer read FIntBeginColumnData write SetBeginColumnData default 1;

    // Event which is triggered after each export of a record
    property OnExportRecords : TOnExportEvent read FOnExportRecords write FOnExportRecords;
  end;

procedure Register;

implementation

uses ComObj; {, ActiveX;}

type
  TOleEnum = type Integer; // Copied from ActiveX unit

procedure Register;
begin
  RegisterComponents('SC', [TscExcelExport]);
end;

{ $R scExcelExport.dcr }

//------------------------------------------------------------------------------
constructor TscExcelExport.Create(Owner: TComponent);
begin
  inherited;
  FStrHeaderText := TStringList.Create;

  FFontHeader := TxlFont.Create;
  FFontTitles := TxlFont.Create;
  FFontData := TxlFont.Create;
  FFontSummary := TxlFont.Create;

  FBorderHeader := TCellBorder.Create;
  FBorderTitles := TCellBorder.Create;
  FBorderData := TCellBorder.Create;
  FBorderSummary := TCellBorder.Create;

  FBorderTitles.FBackColor := clWhite;
  FBorderTitles.FBorderColor := clBlack;
  FBorderTitles.FBorderWeight := bwMedium;
  FBorderTitles.FBorderLineStyle := blNone;

  FBorderData.FBackColor := clWhite;
  FBorderData.FBorderColor := clBlack;
  FBorderData.FBorderWeight := bwMedium;
  FBorderData.FBorderLineStyle := blNone;

  FBorderSummary.FBackColor := clWhite;
  FBorderSummary.FBorderColor := clBlack;
  FBorderSummary.FBorderWeight := bwMedium;
  FBorderSummary.FBorderLineStyle := blNone;

  FBlnExcelVisible := True;
  FConnectTo := ctNewExcel;
  FStrWorksheetName := '';
  FStyleColumnWidth := cwDefault;
  FVisibleFieldsOnly := True;
  FBlnShowTitles := True;
  FIntColumnWidth := 0;
  FIntBlockOfRecords := 20;
  FIntBeginRowHeader := 1;
  FIntBeginRowTitles := 1;
  FIntBeginRowData := 2;
  FIntBeginColumnHeader := 1;
  FIntBeginColumnData := 1;

  FExcelApplication := TExcelApplication.Create(Self);
  FExcelWorkbook := TExcelWorkbook.Create(Self);
  FExcelWorksheet := TExcelWorksheet.Create(Self);

  FFieldNames:=TStringList.Create;
  FSummaryFields:=TStringList.Create;
end;

//------------------------------------------------------------------------------
destructor TscExcelExport.Destroy;
begin
  FStrHeaderText.Free;

  FFontHeader.Free;
  FFontTitles.Free;
  FFontData.Free;
  FFontSummary.Free;

  FBorderHeader.Free;
  FBorderTitles.Free;
  FBorderData.Free;
  FBorderSummary.Free;

  FExcelWorksheet.Free;
  FExcelWorkbook.Free;
  FExcelApplication.Free;

  FFieldNames.Free;
  FSummaryFields.Free;

  inherited;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited;
  if (Operation = opRemove) and (Assigned(FDataset)) and (AComponent = FDataset) then
  begin
    FDataset := nil;
  end;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetHeaderText(const Value: TStrings);
begin
  FStrHeaderText.Assign(Value);

  if FStrHeaderText.Count = 0 then
    FIntBeginRowHeader := 1;

  if FIntBeginRowTitles < FIntBeginRowHeader + FStrHeaderText.Count then
    SetBeginRowTitles(FIntBeginRowHeader + FStrHeaderText.Count);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFontHeader(const Value: TxlFont);
begin
  FFontHeader.Assign(Value);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFontTitles(Value : TxlFont);
begin
  FFontTitles.Assign(Value);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFontData(Value : TxlFont);
begin
  FFontData.Assign(Value);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFontSummary(Value : TxlFont);
begin
  FFontSummary.Assign(Value);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetSummaryFields(Value : TStrings);
begin
  FSummaryFields.Assign(Value);
  FSummaryFields.Text := UpperCase(FSummaryFields.Text);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFontAndBorderRange(DelphiFont : TxlFont; Border : TCellBorder; StrBeginCell, StrEndCell : String);
const
  ALIGNARR: array[THAlignment] of Cardinal = (xlHAlignGeneral, xlHAlignLeft, xlHAlignRight, xlHAlignCenter);
begin
  // Convert Delphi font to the Excel font
  with FExcelWorksheet.Range[StrBeginCell, StrEndCell].Font do
  begin
    Name := DelphiFont.Name;
    Size := DelphiFont.Size;
    Color := DelphiFont.Color;
    Bold :=  fsBold in DelphiFont.Style;
    Italic := fsItalic in DelphiFont.Style;
    Underline := fsUnderline in DelphiFont.Style;
  end;

  if Border.FBackColor <> clWhite then
    FExcelWorksheet.Range[StrBeginCell, StrEndCell].Interior.Color := Border.FBackColor;

  if Border.LineStyle <> blNone then
  begin
    with FExcelWorksheet.Range[StrBeginCell, StrEndCell].Borders do
    begin
      Color := Border.BorderColor;
      case Border.Weight of
        // possible values xlHairline, xlMedium, xlThick, xlThin
        // Abs to prevent warning 'Constant expression violates subrange bounds'
        bwThin : Weight := Abs(xlThin);
        bwMedium : Weight := Abs(xlMedium);
        bwThick : Weight := Abs(xlThick);
      end;

      case Border.LineStyle of
        // possible values xlContinuous, xlDash, xlDashDot, xlDashDotDot, xlDot, xlDouble, xlSlantDashDot, xlLineStyleNone
        blDash : LineStyle := Abs(xlDash);
        blLine : LineStyle := Abs(xlContinuous);
        blDoubleLine : LineStyle := Abs(xlDouble);
      end;
    end;
  end;

  FExcelWorksheet.Range[StrBeginCell, StrEndCell].Orientation:=DelphiFont.Orientation;
  FExcelWorksheet.Range[StrBeginCell, StrEndCell].WrapText := DelphiFont.WrapText;
  FExcelWorksheet.Range[StrBeginCell, StrEndCell].HorizontalAlignment := TOleEnum(ALIGNARR[DelphiFont.Alignment]);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetColumnWidth;
begin
  if FStyleColumnWidth = cwOwnerWidth then
    FExcelWorksheet.Range['A1',GetColumnCharacters(FFieldNames.Count)+'1'].ColumnWidth:=FIntColumnWidth
  else
    if FStyleColumnWidth = cwAutoFit then
      FExcelWorksheet.Range['A1',GetColumnCharacters(FFieldNames.Count)+'1'].EntireColumn.Autofit;
    // else cwFieldDisplayWidth, cwFieldDataSize and cwEnhAutoFit are set in ExportTitles
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetFormat;
var
  IntColumn : Integer;
  StrBeginCell, StrEndCell : String;
begin
  with Dataset do
  begin
    for IntColumn := 1 to FFieldNames.Count do
    begin
      if FieldByName(FFieldNames[IntColumn-1]).DataType = ftString then
      begin
        StrBeginCell := GetColumnCharacters(IntColumn)+IntToStr(FIntBeginRowData);
        StrEndCell := GetColumnCharacters(IntColumn)+IntToStr(FIntTotRecNo + FIntBeginRowData - 1);
        FExcelWorksheet.Range[StrBeginCell,StrEndCell].NumberFormat := '@' //other cases automatic 'general'
      end;
    end;
  end;
end;

//------------------------------------------------------------------------------
function TscExcelExport.CanConvertFieldToCell(Field : TField) : Boolean;
begin
  if (Field.DataType = ftString) or
    (Field.DataType = ftSmallint) or
    (Field.DataType = ftInteger) or
    (Field.DataType = ftWord) or
    (Field.DataType = ftAutoInc) or
    (Field.DataType = ftBoolean) or
    (Field.DataType = ftFloat) or
    (Field.DataType = ftCurrency) or
    (Field.DataType = ftBCD) or
    (Field.DataType = ftDate) or
    (Field.DataType = ftTime) or
    (Field.DataType = ftDateTime) or
    (Field.DataType = ftLargeInt) or
    (Field.DataType = ftWideString) or
    (Field.DataType = ftVariant) then
  begin
    Result:=True;
  end
  else
    Result:=False;
end;

//------------------------------------------------------------------------------
function TscExcelExport.IsValueField(Field : TField) : Boolean;
begin
  if (Field.DataType = ftSmallint) or
    (Field.DataType = ftInteger) or
    (Field.DataType = ftWord) or
    (Field.DataType = ftFloat) or
    (Field.DataType = ftCurrency) then
  begin
    Result:=True;
  end
  else
    Result:=False;
end;

//------------------------------------------------------------------------------
function TscExcelExport.GetWidthFromDatasize(Field : TField) : Integer;
begin
  // Datasize for datetime is to small when also time is saved
  if Field.DataType = ftDateTime then
    Result := 16
  else
    // For all other fieldtypes, just use the datasize
    // Datasize = amount of memory to store value
    Result := Field.Datasize;
end;

//------------------------------------------------------------------------------
// Get Column-character for giving index
//------------------------------------------------------------------------------
function TscExcelExport.GetColumnCharacters(IntNumber : Integer) : String;
begin
  if IntNumber < 1 then
    Result:='A'
  else
  begin
    if IntNumber > 702 then
      Result:='ZZ'
    else
    begin
      if IntNumber > 26 then begin
       if (IntNumber mod 26)=0 then
        Result:=Chr(64 + (IntNumber div 26)-1)
       else
        Result:=Chr(64 + (IntNumber div 26));
         if (IntNumber mod 26)=0 then
          result:=result+chr(64+26)
         else
          result:=Result+Chr(64 + (IntNumber mod 26));
      end
      else
        Result:=Chr(64 + IntNumber);
    end;
  end;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.Disconnect;
begin
  if not FExcelApplication.Visible[LCID] then
  begin
    FExcelApplication.DisplayAlerts[LCID] := False;
    FExcelApplication.Quit;
  end;
  FExcelWorksheet.Disconnect;
  FExcelWorkbook.Disconnect;
  FExcelApplication.Disconnect;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.ExportDataset;
var
  CurPrev : TCursor;
begin
  CurPrev := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    with FDataset do
      if not Active then
        Exit;

    LCID := LOCALE_USER_DEFAULT; //GetUserDefaultLCID;

    // Try to connect to Excel and create new Worksheet
    try
      if FConnectTo = ctNewExcel then
      begin
        FExcelApplication.ConnectKind := ckNewInstance;
        FExcelApplication.Connect;
        FExcelWorkbook.ConnectTo(FExcelApplication.Workbooks.Add(TOleEnum(xlWBATWorksheet), LCID));
        FExcelWorksheet.ConnectTo(FExcelWorkbook.Worksheets[1] as _Worksheet);
      end
      else
      begin
        if FConnectTo = ctNewWorkbook then
        begin
          FExcelApplication.ConnectKind := ckRunningOrNew;
          FExcelApplication.Connect;
          FExcelWorkbook.ConnectTo(FExcelApplication.Workbooks.Add(TOleEnum(xlWBATWorksheet), LCID));
          FExcelWorksheet.ConnectTo(FExcelWorkbook.Worksheets[1] as _Worksheet);
        end
        else
        begin
          FExcelApplication.ConnectKind := ckRunningOrNew;
          FExcelApplication.Connect;
          FExcelWorkbook.ConnectTo(FExcelApplication.ActiveWorkbook);
          FExcelWorksheet.ConnectTo(FExcelWorkbook.Worksheets.Add(EmptyParam,EmptyParam,1,TOleEnum(xlWBATWorksheet),LCID) as _Worksheet);
        end;
      end;
    except
      Exit;
    end;

    FExcelApplication.ScreenUpdating[LCID] := False;

    // If property worksheetname is not filled, worksheet will have name of dataset
    if FStrWorksheetName <> '' then
      FExcelWorksheet.Name := FStrWorksheetName
    else
      FExcelWorksheet.Name := FDataset.Name;

    // Export header
    ExportHeader;

    // Export titels
    ExportTitles;

    FIntTotRecNo := DataSet.RecordCount;
    // Set format (for string fields) of cells before exporting data
    SetFormat;
    // Export data
    ExportFieldData;

    // Calculate summary
    if FSummarySelection <> ssNone then
      ExportSummary;

    // Set width of columns
    SetColumnWidth;

    FExcelWorksheet.Names.Add('naam',EmptyParam,True,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,'a1:b10',EmptyParam);

    FExcelApplication.ScreenUpdating[LCID] := FBlnExcelVisible;
    FExcelApplication.Visible[LCID]:=FBlnExcelVisible;
  finally
    Screen.Cursor := CurPrev;
  end;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.ExportHeader;
var
  i : Integer;
  Matrix : Variant;
  IntHeaderRows : Integer;
  StrBeginColumnChar : String;
begin
  IntHeaderRows := FStrHeaderText.Count;

  if IntHeaderRows = 0 then
    Exit;

  Matrix := VarArrayCreate([1, IntHeaderRows, 1, 1], varOleStr);

  for i := 1 to IntHeaderRows do
    Matrix[i, 1] := FStrHeaderText[i - 1];

  // Get character corresponding with column index (A ... ZZZZ)
  StrBeginColumnChar := GetColumnCharacters(FIntBeginColumnHeader);

  FExcelWorksheet.Range[
    StrBeginColumnChar + IntToStr(FIntBeginRowHeader),
    StrBeginColumnChar + IntToStr(FIntBeginRowHeader+IntHeaderRows-1)].Value := Matrix;
  SetFontAndBorderRange(FFontHeader, FBorderHeader,
    StrBeginColumnChar + IntToStr(FIntBeginRowHeader),
    StrBeginColumnChar + IntToStr(FIntBeginRowHeader+IntHeaderRows-1));
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.ExportTitles;
var
  IntColumn : Integer;
  IntFieldIndex : Integer;
  StrCell : String;
  StrColumn : String;
  StrTitle : String;
  FltFontSizeFactor : Real;
  FltTitleFontSizeFactor : Real;
begin
  with FDataset do
  begin
    FStrBeginColumnDataChar := GetColumnCharacters(FIntBeginColumnData);

    FFieldNames.Clear;

    if FBlnShowTitles then
    begin
      for IntColumn := FIntBeginColumnData to (Fields.Count + FIntBeginColumnData -1) do
      begin
        IntFieldIndex := IntColumn - FIntBeginColumnData;

        // Only export fields which are writable in an Excel cell
        // Don't export non visible fields if VisibleFieldsOnly is True
        // Add these fields to a list, so this list can be used when exporting data
        if CanConvertFieldToCell(Fields[IntFieldIndex]) and
           ((not VisibleFieldsOnly) or (VisibleFieldsOnly and Fields[IntFieldIndex].Visible))
        then
        begin
          StrColumn := GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData);
          StrCell:=StrColumn+IntToStr(FIntBeginRowTitles);

          FFieldNames.Add(Fields[IntFieldIndex].FieldName);

          // Use DisplayName of column if this is filled in, otherwise use FieldName
          if Fields[IntFieldIndex].DisplayName <> '' then
            StrTitle := Fields[IntFieldIndex].DisplayName
          else
            StrTitle := Fields[IntFieldIndex].FieldName;

          FExcelWorksheet.Range[StrCell,StrCell].Value := StrTitle;

          // Use DisplayField of each field to set the column width
          if FStyleColumnWidth = cwFieldDisplayWidth then
          begin
            // Value of datasize fits when font size = 10, so calculate factor when font size is larger
            FltFontSizeFactor := FFontData.Size / 10;

            FExcelWorksheet.Range[StrCell,StrCell].ColumnWidth:=Integer(Round(Fields[IntFieldIndex].DisplayWidth * FltFontSizeFactor));
          end
          else
          begin
            // Use Datasize of each field to set the column width
            if FStyleColumnWidth = cwFieldDataSize then
            begin
              // Value of datasize fits when font size = 10, so calculate factor when font size is larger
              FltFontSizeFactor := FFontData.Size / 10;

              FExcelWorksheet.Range[StrCell,StrCell].ColumnWidth:=Integer(Round(GetWidthFromDatasize(Fields[IntFieldIndex]) * FltFontSizeFactor));
            end
            else
            begin
              // Style = adaptive -> use DisplayWidth of TField except when title of column is larger
              if FStyleColumnWidth = cwEnhAutoFit then
              begin
                // Value of datasize fits when font size = 10, so calculate factor when font size is larger
                FltFontSizeFactor := FFontData.Size / 10;
                FltTitleFontSizeFactor := FFontTitles.Size / 10;

                if ((Length(StrTitle) + 1) * FltTitleFontSizeFactor) > (Fields[IntFieldIndex].DisplayWidth * FltFontSizeFactor) then
                  FExcelWorksheet.Range[StrCell,StrCell].ColumnWidth:=Integer(Round((Length(StrTitle) + 1) * FltTitleFontSizeFactor) + 1)
                else
                  FExcelWorksheet.Range[StrCell,StrCell].ColumnWidth:=Integer(Round(Fields[IntFieldIndex].DisplayWidth * FltFontSizeFactor));
              end
              // else cwDefault, cwOwnerWidth, cwAutoFit
              // These columns widths are set after exporting all data in the procedure SetColumnWidth
            end;
          end;
        end;
      end;
      SetFontAndBorderRange(FFontTitles, FBorderTitles, FStrBeginColumnDataChar+IntToStr(FIntBeginRowTitles),
        GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData -1)+IntToStr(FIntBeginRowTitles));
    end
    else
    begin
      // Titles will not be visible, but run through fields
      for IntColumn := FIntBeginColumnData to (Fields.Count + FIntBeginColumnData -1) do
      begin
        IntFieldIndex := IntColumn - FIntBeginColumnData;

        // Only export fields which are writable in an Excel cell
        // Don't export non visible fields if VisibleFieldsOnly is True
        // Add these fields to a list, so this list can be used when exporting data
        if CanConvertFieldToCell(Fields[IntFieldIndex]) and
           ((not VisibleFieldsOnly) or (VisibleFieldsOnly and Fields[IntFieldIndex].Visible))
        then
        begin
          StrColumn := GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData);
          FFieldNames.Add(Fields[IntFieldIndex].FieldName);
        end;
      end;
    end;
  end;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.ExportFieldData;
var
  IntColumn : Integer;
  IntBeginRow, IntEndRow : Integer;
  IntMatrixRow : Integer;
  PtBookmark : TBookmark;
  Matrix : Variant;
begin
  FIntRecordNo := 0;
  with FDataset do
  begin
    DisableControls;
    PtBookmark := GetBookmark;
    try
      // Create a matrix of variants
      // -  Columns = number of fields
      // -  Rows    = block of records (FIntBlockOfRecords)
      Matrix := VarArrayCreate([1,FIntBlockOfRecords,1,FFieldNames.Count],varVariant);

      IntBeginRow := FIntBeginRowData;
      IntEndRow := FIntBeginRowData + FIntBlockOfRecords-1 ;

      IntMatrixRow := 0;

      First;
      while not EOF do
      begin
        Inc(FIntRecordNo);
        Inc(IntMatrixRow);
        if Assigned(FOnExportRecords) then
          FOnExportRecords(Self,FIntRecordNo);

        for IntColumn := 1 to FFieldNames.Count do
          Matrix[IntMatrixRow,IntColumn] := FieldByName(FFieldNames[IntColumn - 1]).AsVariant;

        // Create a new block of records to export to Excel
        // Don't export all data to one variant matrix because memory has it limitations
        // Property FIntBlockOfRecords is default 20 records

        // check if matrix is full, and if so, write the block to excel
        if (FIntRecordNo mod FIntBlockOfRecords = 0) then
        begin
          FExcelWorksheet.Range[FStrBeginColumnDataChar+IntToStr(IntBeginRow),GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData - 1)+IntToStr(IntEndRow)].Value := Matrix;
          IntBeginRow := IntBeginRow + FIntBlockOfRecords;      // next insert starts here
          IntEndRow := IntBeginRow   + FIntBlockOfRecords -1;   // next block ends here
          IntMatrixRow := 0;                                    // reset index into matrix
        end;

        Next;
      end;

      // Now that EOF is true, so check if the matrix has remaining data to write
      if (IntMatrixRow > 0) then
      begin
        // recalculate the block's end
        IntEndRow := IntBeginRow + IntMatrixRow-1;
        // Write remaining block
        FExcelWorksheet.Range[FStrBeginColumnDataChar+IntToStr(IntBeginRow),GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData - 1)+IntToStr(IntEndRow)].Value := Matrix;
      end;

    finally
      GotoBookmark(PtBookmark);
      FreeBookmark(PtBookmark);
      EnableControls;
    end;
  end;

  FIntEndRowData := IntEndRow;

  SetFontAndBorderRange(FFontData, FBorderData,FStrBeginColumnDataChar+IntToStr(FIntBeginRowData),
    GetColumnCharacters(FFieldNames.Count + FIntBeginColumnData -1)+IntToStr(FIntEndRowData));
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.ExportSummary;
const
  SUM_ARR: Array[TSummaryCalculation] of String = ('SUM', 'MIN', 'MAX');
var
  IntColumn : Integer;
  StrCell : String;
  StrCalc : String;
  StrBeginCell, StrEndCell : String;

  function Summarized(aColumn: Integer): Boolean;
  begin
    case FSummarySelection of
    ssValues:
      Result := IsValueField(
        FDataSet.FieldByName(FFieldNames[aColumn - 1]));
    ssGiven:
      Result := FSummaryFields.IndexOf(
        UpperCase(FFieldNames[aColumn - 1])) > -1;
    else
      Result := False;
    end;
  end;

begin
  with FDataset do
  begin
    for IntColumn := 1 to FFieldNames.Count do
    begin
      if Summarized(IntColumn) then
      begin
        StrCell:=GetColumnCharacters(IntColumn)+IntToStr(FIntRecordNo + FIntBeginRowData);
        StrCalc := SUM_ARR[FSummaryCalculation];

        StrBeginCell := GetColumnCharacters(IntColumn)+IntToStr(FIntBeginRowData);
        StrEndCell := GetColumnCharacters(IntColumn)+IntToStr(FIntRecordNo + FIntBeginRowData - 1);
        FExcelWorksheet.Range[StrCell,StrCell].Value := Format('=%s(%s:%s)', [StrCalc, StrBeginCell, StrEndCell]);
      end;
    end;
  end;
  SetFontAndBorderRange(FFontSummary, FBorderSummary, 'A'+IntToStr(FIntRecordNo + FIntBeginRowData),
    GetColumnCharacters(FFieldNames.Count)+IntToStr(FIntRecordNo + FIntBeginRowData));
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SaveAs(const StrFileName : String; const FileFormat : TFileFormat);
begin
  FExcelApplication.DisplayAlerts[LCID] := False;
  // Export data to a file
  case FileFormat of
    ffXLS : FExcelWorksheet.SaveAs(StrFileName,TOleEnum(xlWorkbookNormal));
    // For 97 and 2000 compatible format
    ffXL97: FExcelWorksheet.SaveAs(StrFileName,TOleEnum(xlExcel9795));
    ffCSV : FExcelWorksheet.SaveAs(StrFileName,TOleEnum(xlCSV));
//    ffHTM : FExcelWorksheet.SaveAs(StrFileName,TOleEnum(xlHtml));  // Only works with Excel2000
  end;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.PrintPreview(const BlnPrintGridLines : Boolean);
begin
  // Show PrintPreview of Excel
  FExcelWorksheet.PageSetup.PrintGridlines:=BlnPrintGridLines;
  FExcelWorksheet.PageSetup.CenterHeader:=FExcelWorksheet.Name;
  FExcelApplication.ScreenUpdating[LCID] := True;
  FExcelApplication.Visible[LCID]:=True;
  FExcelWorksheet.PrintPreview;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetVisibleFieldsOnly(const Value: Boolean);
begin
  FVisibleFieldsOnly := Value;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetBeginRowHeader(const Value: Integer);
begin
  if FStrHeaderText.Count > 0 then
  begin
    if Value > 0 then
      FIntBeginRowHeader := Value
    else
      FIntBeginRowHeader := 1;

    if FIntBeginRowTitles < FIntBeginRowHeader + FStrHeaderText.Count - 1 then
      SetBeginRowTitles(FIntBeginRowHeader + FStrHeaderText.Count);
  end
  else
    FIntBeginRowHeader := 1;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetBeginRowTitles(const Value: Integer);
begin
  if Value < FIntBeginRowHeader + FStrHeaderText.Count then
    FIntBeginRowTitles := FIntBeginRowHeader + FStrHeaderText.Count
  else
    FIntBeginRowTitles := Value;

  if FIntBeginRowTitles >= FIntBeginRowData then
     SetBeginRowData(FIntBeginRowTitles + 1);
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetBeginRowData(const Value: Integer);
begin
  if Value <= FIntBeginRowTitles then
    FIntBeginRowData := FIntBeginRowTitles + 1
  else
    FIntBeginRowData := Value;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetBeginColumnData(const Value: Integer);
begin
  if Value < 1 then
    FIntBeginColumnData := 1
  else
    FIntBeginColumnData := Value;
end;

//------------------------------------------------------------------------------
procedure TscExcelExport.SetBeginColumnHeader(const Value: Integer);
begin
  if Value < 1 then
    FIntBeginColumnHeader := 1
  else
    FIntBeginColumnHeader := Value;
end;

end.


