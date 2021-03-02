unit mainform;

{  Program generates spreadsheet file based on data contained in report_data_xml_file.
   Report_data_xml_file is built by QM program NXTREPORT
   Program is normally called by NXTREPORT on completion of report_data_xml_file creation.

Command line: NxtReportGen.exe -f "report_data_xml_file" -s xlsx
Options: -d for debugging
         -x for display xml parsing
         -t for use timeout timer


         This is free and unencumbered software released into the public domain.

         Anyone is free to copy, modify, publish, use, compile, sell, or
         distribute this software, either in source code form or as a compiled
         binary, for any purpose, commercial or non-commercial, and by any
         means.

         In jurisdictions that recognize copyright laws, the author or authors
         of this software dedicate any and all copyright interest in the
         software to the public domain. We make this dedication for the benefit
         of the public at large and to the detriment of our heirs and
         successors. We intend this dedication to be an overt act of
         relinquishment in perpetuity of all present and future rights to this
         software under copyright law.

         THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
         EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
         MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
         IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
         OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
         ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
         OTHER DEALINGS IN THE SOFTWARE.

         For more information, please refer to <http://unlicense.org/>
}
{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ActnList, Buttons, StrUtils,
  fpspreadsheetgrid, fpspreadsheetctrls, fpscell,
  DOM, XMLRead, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    CBShowXML: TCheckBox;
    CBDebug: TCheckBox;
    CheckBox1: TCheckBox;
    Memo1: TMemo;
    Panel1: TPanel;
    PopupMenu1: TPopupMenu;
    Timer1: TTimer;
    WorksheetGrid1: TsWorksheetGrid;

    WorksheetGrid2: TsWorksheetGrid;
    //    procedure BtnNewClick(Sender: TObject);

    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { private declarations }
    procedure LoadFile(const AFileName: string);
    procedure ProcessCell(const fmtcell, loccell, cellvalue: string);
    procedure Process;
    procedure ProecssXML;
    procedure SaveReport;
    procedure SavePageFmt;
  public
    { public declarations }
  end;

const
  INT_QUERY_FUNCTION = 998;
  INT_LOOKUP_FUNCTION = 999;
  TimeoutInterval = 30000;  //timeout 30 seconds

var
  Form1: TForm1;
  FormatFileName: string;
  ReportFileName: string;
  SS_Type: string;   // spreadsheet type by extenstion xlsx, xls, ods
  LogFile: string;

//  WorksheetGrid1: TsWorksheetGrid;
implementation

uses
  fpcanvas, lazutf8,
  fpstypes, fpsutils, fpsexprparser, fpspreadsheet;
//fpsRegFileFormats,


// REMEMBER !!! WorksheetGrid.ReadFormulas := True!

// Dummy QMQuery Function to allow save of =qmquery() text in cell record
//RegisterFunction('qmQUERY', 'S', 'S', INT_QUERY_FUNCTION, @fpsQUERY);
procedure fpsQUERY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY QUERY FUNCTION
var
  s: string;

begin
  s := '';
  if Args[1].ResultType = rtError then
  begin
    Result := ErrorResult(Args[1].ResError);
    exit;
  end;
  s := ArgToString(Args[1]);
  Result := StringResult(s);

end;

// Dummy QMLOOKUP Function to allow save of =QMLOOKUP text in cell record
//RegisterFunction('qmLOOKUP', 'S', 'S', INT_LOOKUP_FUNCTION, @fpsLOOKUP);
procedure fpsLOOKUP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY LOOKUP FUNCTION
var
  s: string;

begin
  s := '';
  if Args[1].ResultType = rtError then
  begin
    Result := ErrorResult(Args[1].ResError);
    exit;
  end;
  s := ArgToString(Args[1]);
  Result := StringResult(s);
end;
// Dummy QMREPORT Function to allow save of =qmreport() text in cell record
//RegisterFunction('qmREPORT', 'S', 'S', INT_LOOKUP_FUNCTION, @fpsREPORT);
procedure fpsReport(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY LOOKUP FUNCTION
var
  s: string;

begin
  s := '';
  if Args[1].ResultType = rtError then
  begin
    Result := ErrorResult(Args[1].ResError);
    exit;
  end;
  s := ArgToString(Args[1]);
  Result := StringResult(s);
end;
procedure fpsREPLACE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY REPLACE FUNCTION
var
  s: string;
  i: integer;
begin
  s := '';
  for i := 0 to Length(Args) - 1 do
  begin
    if Args[i].ResultType = rtError then
    begin
      Result := ErrorResult(Args[i].ResError);
      exit;
    end;
    //    s := s + ArgToString(Args[i]) + ', ';
    s := ExtractWord(1, ArgToString(Args[0]), [' ']);
  end;
  Result := StringResult(s);

end;

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);

begin
  RegisterFunction('QMQUERY', 'S', 'S+', INT_QUERY_FUNCTION, @fpsQUERY);
  RegisterFunction('QMLOOKUP', 'S', 'S+', INT_QUERY_FUNCTION, @fpsLOOKUP);
  RegisterFunction('QMREPORT', 'S', 'S+', INT_QUERY_FUNCTION, @fpsREPORT);
  RegisterFunction('MVREPLACE', 'S', 'S+', INT_QUERY_FUNCTION, @fpsREPLACE);
  LogFile := ExtractFilePath(ParamStr(0)) + 'NxtReportGen.Log';
  Memo1.Lines.Clear;
  Process();
end;

procedure TForm1.Timer1Timer(Sender: TObject);
// if report generation does not occur within TimeoutInterval, abort!
begin
  Memo1.Lines.Add('NxtReportGen Timed Out: ' + DateTimeToStr(Now));
  // always write the log
  Memo1.Lines.SaveToFile(LogFile);

  application.Terminate;
end;

procedure TForm1.Process;
var
  LogMsg, timercounts: string;
  I, timercount: integer;
begin

  Memo1.Lines.Add('NxtReportGen Started: ' + DateTimeToStr(Now));
  LogMsg := '';
  for I := 0 to ParamCount do
    LogMsg := logMsg + ' ' + ParamStr(I);
  Memo1.Lines.Add(LogMsg);

  // test command line for valid parameters (
  LogMsg := Application.CheckOptions('fs:txd::',
    'file: sstype: timeout:: xml:: debug::');
  if length(LogMsg) > 0 then
  begin
    // if there is a command line error for log creation
    Memo1.Lines.Add(LogMsg);
    CBDebug.Checked := True;
  end
  else
  begin
    // debugging ??
    if Application.HasOption('d', 'debug') then
      CBDebug.Checked := True;
    if Application.HasOption('x', 'xml') then
      CBShowXML.Checked := True;

    if Application.HasOption('t', 'timeout') then
    begin
      timercounts := Application.GetOptionValue('t', 'timeout');
      if TryStrToInt(timercounts, timercount) then
        timer1.Interval := timercount
      else
        Timer1.Interval := TimeoutInterval;  // use default value
      Timer1.Enabled := True;
    end;
    // spreadsheet type?
    SS_Type := 'xlsx';
    if Application.HasOption('s', 'sstype') then
    begin
      SS_Type := Application.GetOptionValue('s', 'sstype');
      if not ((SS_Type = 'xlsx') or (SS_Type = 'xls') or (SS_Type = 'ods')) then
      begin
        Memo1.Lines.Add(SS_Type + ' Invalid spreadsheet type, using xlsx');
        SS_Type := 'xlsx';
      end;
    end
    else
      Memo1.Lines.Add('Missing spreadsheet type specifier, using xlsx');

    // get and test file name
    if Application.HasOption('f', 'file') then
    begin
      ReportFileName := Application.GetOptionValue('f', 'file');
      if not (FileExists(ReportFileName)) then
      begin
        Memo1.Lines.Add(ReportFileName + ' Does Not Exist!');
      end
      else
      begin
        //           WorksheetGrid1 := TsWorksheetGrid.Create(nil);
        try
          ProecssXML();
        finally
          //               WorksheetGrid1.Free
        end;
      end;

    end
    else
      Memo1.Lines.Add('NxtReportGen Called Without File Parameter');

  end;

  // Delete .xml file if not debugging
  If not CBDebug.Checked Then
    if DeleteFile(ReportFileName) then
       Memo1.Lines.Add(ReportFileName + ' - deleted')
    else
       Memo1.Lines.Add(ReportFileName + ' - failed to delete!');


  Memo1.Lines.Add('NxtReportGen Complete: ' + DateTimeToStr(Now));
  // always write the log
  Memo1.Lines.SaveToFile(LogFile);

  application.Terminate;
end;

procedure TForm1.ProecssXML;
// Procedure loads and parses the report xml stream
// important elements / attributes:
// filename  - name of .ods file which contains the report format
//             contained in the File element:
//              <File filename="C:\Nextjob\Laz-Projects\NxtReportGen\test1.ods">
// fmt - cell location in filename to use for formating the data
// loc - cell location to be created in generated report ss
//       both attributes contianed within the CEL element:
//       <CEL fmt="C3" loc="C5">C5 VALUE</CEL>
var
  Doc: TXMLDocument;
  Child: TDOMNode;
  fmtcell: string;   // cell format address to read format from
  loccell: string;   // cell location to write to
  cellvalue: string; // value to be writen to report cell
  //  fmtfilename : String; // Full path and name of ods format file found in xml data
  j, i: integer;
begin
  FormatFileName := '';
  try
    try
      ReadXMLFile(Doc, ReportFileName);
      // using FirstChild and NextSibling properties
      Child := Doc.DocumentElement.FirstChild;
      while Assigned(Child) do
      begin
        for i := 0 to Child.Attributes.Length - 1 do
        begin
          if CBShowXML.Checked then
            Memo1.Lines.Add(Child.NodeName + ' ' + Child.Attributes.Item[i].NodeName +
              ' ' + Child.Attributes.Item[i].NodeValue);
          if Child.Attributes.Item[i].NodeName = 'filename' then
            FormatFileName := Child.Attributes.Item[i].NodeValue;

        end;
        if CBDebug.Checked then
          Memo1.Lines.Add('filename = ' + FormatFileName);
        if Length(FormatFileName) > 0 then
          if FileExists(FormatFileName) then
            LoadFile(FormatFileName)
          else
          begin
            if CBDebug.Checked then
              Memo1.Lines.Add('filename = ' + FormatFileName + ' Not Found, Abort Report');
            exit;
          end;
        // using ChildNodes method
        with Child.ChildNodes do
          try
            for j := 0 to (Count - 1) do
            begin
              cellvalue := '';
              if CBShowXML.Checked then
                Memo1.Lines.Add(Item[j].NodeName + ' ' + Item[j].FirstChild.NodeValue);
              if Item[j].NodeName = 'CEL' then
                cellvalue := Item[j].FirstChild.NodeValue;
              fmtcell := '';
              loccell := '';
              if Item[j].HasAttributes then
              begin
                for i := 0 to Item[j].Attributes.Length - 1 do
                begin
                  if CBShowXML.Checked then
                    Memo1.Lines.Add('Attributes - ' + Item[j].Attributes.Item[i].NodeName +
                      ' ' + Item[j].Attributes.Item[i].NodeValue);
                  if Item[j].Attributes.Item[i].NodeName = 'fmt' then
                    fmtcell := Item[j].Attributes.Item[i].NodeValue
                  else if Item[j].Attributes.Item[i].NodeName = 'loc' then
                    loccell := Item[j].Attributes.Item[i].NodeValue;
                end;
                // move the data to the report spreadsheet
                if (System.Length(fmtcell) > 0) and (System.Length(loccell) > 0) then
                begin
                  if CBDebug.Checked then
                    Memo1.Lines.Add('Adding ' + Item[j].FirstChild.NodeValue +
                      ' to Report at  --loc = ' + loccell + ' Using -- fmt = ' + fmtcell);
                  ProcessCell(fmtcell, loccell, cellvalue);
                end;
              end;
            end;
          finally
            Free;
          end;
        Child := Child.NextSibling;
      end;
      // save the created report
      SaveReport;
    finally
      Doc.Free;
    end;

  except
    // handle xml error
    on E: Exception do begin
       Memo1.Lines.Add('NxtReportGen Exception raised: ' + E.Message);
    end;
  end;
end;

procedure TForm1.ProcessCell(const fmtcell, loccell, cellvalue: string);

// procdure populates report spreadsheet with formating from format spreadsheet
// fmtcell - format spreadsheet cell to use for formating
// loccell - report spreadsheet cell to populate
// cellvalue - data value to place in report ss
var
  LocCellPtr, FmtCellPtr: PCell;
  locColumn, locRow, fmtColumn, fmtRow: cardinal;
  FmtFont: TsFont;
  nf: TsNumberFormat;
  nfs : String;

begin
  // resolve excel address notation to pointer to cell (create cell if it does not
  // exist
  FmtCellPtr := WorksheetGrid1.Worksheet.GetCell(fmtcell);  // format ss cell pointer
  LocCellPtr := WorksheetGrid2.Worksheet.GetCell(loccell);  // report ss cell pointer
  // write the value to the report
  WorksheetGrid2.Worksheet.WriteCellValueAsString(LocCellPtr, cellvalue);
  // number format
  // LocCellPtr^.NumberFormatStr:=FmtCellPtr^.NumberFormatStr;
  WorksheetGrid1.Worksheet.ReadNumFormat(FmtCellPtr, nf, nfs);
  WorksheetGrid2.Worksheet.WriteNumberFormat(LocCellPtr, nfCustom, nfs);
  // copy over font
  // first get the Font record of the format cell
  FmtFont := WorksheetGrid1.Worksheet.ReadCellFont(FmtCellPtr);
  // Now write it out to the report cell
  WorksheetGrid2.Worksheet.WriteFont(LocCellPtr, FmtFont.FontName,
    FmtFont.Size, FmtFont.Style, FmtFont.Color, FmtFont.Position);

  // Now write the Cell Fromat properties to the report cell
  WorksheetGrid2.Worksheet.WriteTextRotation(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadTextRotation(FmtCellPtr));
  WorksheetGrid2.Worksheet.WriteHorAlignment(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadHorAlignment(FmtCellPtr));
  WorksheetGrid2.Worksheet.WriteVertAlignment(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadVertAlignment(FmtCellPtr));
  WorksheetGrid2.Worksheet.WriteBorders(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadCellBorders(FmtCellPtr));
  WorksheetGrid2.Worksheet.WriteBorderStyles(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadCellBorderStyles(FmtCellPtr));

  // Backgound color
  WorksheetGrid2.Worksheet.WriteBackgroundColor(
    LocCellPtr, WorksheetGrid1.Worksheet.ReadBackgroundColor(FmtCellPtr));

  // Column width & Row Height
  ParseCellString(loccell, locRow, locColumn);
  ParseCellString(fmtcell, fmtRow, fmtColumn);

  //if CBDebug.checked then Memo1.Lines.Add('Format Column Size: ' + fmtcell + ' ' + FloatToStr(WorksheetGrid1.Worksheet.GetColWidth(fmtColumn,suChars)));
  //if CBDebug.checked then Memo1.Lines.Add('   Location Column Size: ' + loccell + ' ' + FloatToStr(WorksheetGrid2.Worksheet.GetColWidth(locColumn,suChars)));

  if int(WorksheetGrid1.Worksheet.GetColWidth(fmtColumn, suMillimeters) * 100) <>
    int(WorksheetGrid2.Worksheet.GetColWidth(locColumn, suMillimeters) * 100) then
    WorksheetGrid2.Worksheet.WriteColWidth(
      locColumn, WorksheetGrid1.Worksheet.GetColWidth(fmtColumn, suMillimeters), suMillimeters);

  //if CBDebug.checked then Memo1.Lines.Add('Final Location Column Size: ' + loccell + ' ' + FloatToStr(WorksheetGrid2.Worksheet.GetColWidth(locColumn,suChars)));

  if int(WorksheetGrid1.Worksheet.GetRowHeight(fmtRow, suMillimeters) * 100) <>
    int(WorksheetGrid2.Worksheet.GetRowHeight(locRow, suMillimeters) * 100) then
    WorksheetGrid2.Worksheet.WriteRowHeight(
      locRow, WorksheetGrid1.Worksheet.GetRowHeight(fmtRow, suMillimeters), suMillimeters);

  //  need to test out Number formating (only preform if type is numeric?)

  //  WorksheetGrid1.Worksheet.ReadNumFormat(FmtCellPtr,ANumFormat,ANumFormatStr);
  //  WorksheetGrid2.Worksheet.WriteNumberFormat(LocCellPtr,ANumFormat,ANumFormatStr);
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.SaveReport;
var
  err, MyReportFileName: string;
begin
  if WorksheetGrid2.Workbook = nil then
    exit;
  SavePageFmt;
  MyReportFileName := ChangeFileExt(ReportFileName, '.' + SS_type);
  Memo1.Lines.Add('Writing Report To: ' + MyReportFileName);
  Screen.Cursor := crHourglass;
  try
    WorksheetGrid2.SaveToSpreadsheetFile(UTF8ToAnsi(MyReportFileName));
    //WorksheetGrid.WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));     // works as well
  finally
    Screen.Cursor := crDefault;
    // Show a message in case of error(s)
    err := WorksheetGrid2.Workbook.ErrorMsg;
    if err <> '' then
      Memo1.Lines.Add(err);
  end;

end;
// routine copies page format properties from format ss to report ss
procedure TForm1.SavePageFmt;
Begin
    // Save Page Layout Settings
  WorksheetGrid2.Worksheet.PageLayout.Orientation := WorksheetGrid1.Worksheet.PageLayout.Orientation;
  // Paper Size
  WorksheetGrid2.Worksheet.PageLayout.PageWidth := WorksheetGrid1.Worksheet.PageLayout.PageWidth;
  WorksheetGrid2.Worksheet.PageLayout.PageHeight := WorksheetGrid1.Worksheet.PageLayout.PageHeight;

  // Repeated rows
   WorksheetGrid2.Worksheet.PageLayout.SetRepeatedCols(WorksheetGrid1.Worksheet.PageLayout.RepeatedCols.FirstIndex,WorksheetGrid1.Worksheet.PageLayout.RepeatedCols.LastIndex);
   WorksheetGrid2.Worksheet.PageLayout.SetRepeatedRows(WorksheetGrid1.Worksheet.PageLayout.RepeatedRows.FirstIndex,WorksheetGrid1.Worksheet.PageLayout.RepeatedRows.LastIndex);


  // set margins
    WorksheetGrid2.Worksheet.PageLayout.TopMargin := WorksheetGrid1.Worksheet.PageLayout.TopMargin;
    WorksheetGrid2.Worksheet.PageLayout.BottomMargin := WorksheetGrid1.Worksheet.PageLayout.BottomMargin;
    WorksheetGrid2.Worksheet.PageLayout.LeftMargin := WorksheetGrid1.Worksheet.PageLayout.LeftMargin;
    WorksheetGrid2.Worksheet.PageLayout.RightMargin := WorksheetGrid1.Worksheet.PageLayout.RightMargin;
  // copies
    WorksheetGrid2.Worksheet.PageLayout.Copies := WorksheetGrid1.Worksheet.PageLayout.Copies;
  // options
    WorksheetGrid2.Worksheet.PageLayout.Options := WorksheetGrid1.Worksheet.PageLayout.Options ;
    // all columns on one page width
    WorksheetGrid2.Worksheet.PageLayout.FitWidthToPages := WorksheetGrid1.Worksheet.PageLayout.FitWidthToPages;
    // use as many pages as needed
    WorksheetGrid2.Worksheet.PageLayout.FitHeightToPages := WorksheetGrid1.Worksheet.PageLayout.FitHeightToPages;

end;



// Loads first worksheet from file into grid
procedure TForm1.LoadFile(const AFileName: string);
var
  err: string;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      WorksheetGrid1.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));
      // size our new ss to match format ss
      WorksheetGrid2.NewWorkbook(WorksheetGrid1.ColCount, WorksheetGrid1.RowCount);

    except
      on E: Exception do
      begin
        // Empty worksheet instead of the loaded one
        WorksheetGrid1.NewWorkbook(26, 100);
        {Caption := 'fpsGrid - no name'; }
        // Grab the error message
        WorksheetGrid1.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := WorksheetGrid1.Workbook.ErrorMsg;
    if err <> '' then
      Memo1.Lines.Add(err);
  end;
end;


initialization
  {$I mainform.lrs}

end.
