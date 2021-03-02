unit umainform;

{
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

Note This program makes use of (copies) code from Spready  an extended application of the the entire fpspreadsheet library
showing spreadsheet files with formatting, editing of cells, etc
Speady is found in the  applications/spready folder of the Lazarus Components and Code Library
https://sourceforge.net/p/lazarus-ccr/svn/HEAD/tree/applications/spready/

Note This program makes use of Delphi4QM -A Delphi and Free Pascal Wrapper for the QMClient C API 
https://github.com/FOSS4MV/delphi4qm
}

{
Notes about QMClient use:
Most calls are defined in tisQMClient, however we need to use a few that are not defined.
    function QMCallx(subrname: PAnsiChar; argc: Integer; a1, a2, a3, a4, a5, a6, a7: PAnsiChar): Integer; cdecl; external 'qmclilib.dll' name 'QMCallx';
    function QmGetArg(ArgNo: Integer): PAnsiChar; cdecl; external 'qmclilib.dll' name 'QMGetArg';
    procedure QMFree(p: PAnsiChar); cdecl; external 'qmclilib.dll' name 'QMFree';
We need to use QMCallx beacuse:
 passed variables using QMCall are getting lost unless type cast to PAnsiChar (this appears to be a FPC problem, works with Delphi)
 changing passed variables in the called subroutine, if cast to PAnsiChar in call results in seqfault.
 Using QMCallx and  QMGetArg seems to fix these two issues
 Needed to add our own reference to QMFree becuase of its location in tisQMClientAPI causes name space confilicts
}
{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, strutils,
  StdCtrls, Menus, ExtCtrls, ActnList, Spin, Buttons, ComCtrls, Grids, typinfo,
  fpspreadsheetgrid, fpspreadsheetctrls, fpstypes, fpsutils, fpsexprparser, fpspreadsheet,
  fpsReaderWriter, fpsActions, fpscsv, fpsopendocument, fpsallformats, fpsNumFormat, tisQMClient;

type

  { TMainForm }

  TMainForm = class(TForm)
    ActionList1: TActionList;
    BtnPageSU: TButton;
    CBFunctions: TComboBox;
    edFmtType: TEdit;
    edFmtStr: TEdit;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    lblColInfo: TLabel;
    lblFmtStr: TLabel;
    lblFmtType: TLabel;
    lblPageHeight: TLabel;
    lblCellWidth: TLabel;
    lblCellHeight: TLabel;
    lblSSHeight: TLabel;
    lblPageWidth: TLabel;
    lblSSWidth: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MnuNew: TMenuItem;
    MnuFile: TMenuItem;
    MnuOpen: TMenuItem;
    MnuSave: TMenuItem;
    MnuTestReport: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    Panel1: TPanel;
    AcCopyToClipboard: TsCopyAction;
    AcNumFormatCustom: TsNumberFormatAction;
    sPaste: TsCopyAction;
    sCut: TsCopyAction;
    SelectReportDialog: TOpenDialog;
    sCellBackColor: TsCellCombobox;
    sCellTopThk: TsCellBorderAction;
    sCellLeftThk: TsCellBorderAction;
    sCellRightThk: TsCellBorderAction;
    sCellRightLn: TsCellBorderAction;
    sCellouterLn: TsCellBorderAction;
    sCellOuterThkLn: TsCellBorderAction;
    sCellAllLns: TsCellBorderAction;
    sCellTopLn: TsCellBorderAction;
    sCelBottomLn: TsCellBorderAction;
    sCellBottomThkLn: TsCellBorderAction;
    sCellBottomDblLn: TsCellBorderAction;
    sCellLeftLn: TsCellBorderAction;
    MenuItemT: TMenuItem;
    Panel3: TPanel;
    OpenDialog: TOpenDialog;
    Panel2: TPanel;
    BoarderMenu1: TPopupMenu;
    SaveDialog: TSaveDialog;
    sCelFontName: TsCellCombobox;
    sCelFontSize: TsCellCombobox;
    sCellFontColor: TsCellCombobox;
    sCellEdit1: TsCellEdit;
    sCellIndicator1: TsCellIndicator;
    sFontStyleUnderLn: TsFontStyleAction;
    sFontStyleStrike: TsFontStyleAction;
    sFontStyleItalic: TsFontStyleAction;
    sFontStyleBold: TsFontStyleAction;
    sHorAlignLeft: TsHorAlignmentAction;
    sHorAlignCenter: TsHorAlignmentAction;
    sHorAlignRight: TsHorAlignmentAction;
    sNoCellBorders: TsNoCellBordersAction;
    sVertAlignTop: TsVertAlignmentAction;
    sVertAlignCenter: TsVertAlignmentAction;
    sVertAlignBottom: TsVertAlignmentAction;
    sWorkbookSource1: TsWorkbookSource;
    ToolBar1: TToolBar;
    TBBold: TToolButton;
    TBItalic: TToolButton;
    TBStirkeOut: TToolButton;
    TBHorzAlignCenter: TToolButton;
    TBHorzAlignRight: TToolButton;
    TBCopy: TToolButton;
    TBPaste: TToolButton;
    ToolButton1: TToolButton;
    ToolButton12: TToolButton;
    TBInsCol: TToolButton;
    TBInsRow: TToolButton;
    TBDelCol: TToolButton;
    TBDelRow: TToolButton;
    TBUnderline: TToolButton;
    TBCut: TToolButton;
    TBFormat: TToolButton;
    ToolButton3: TToolButton;
    TBBarderMenu: TToolButton;
    TBVertAlignTop: TToolButton;
    TBVertAlignCenter: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    TBHorzAlignLeft: TToolButton;
    WorksheetGrid: TsWorksheetGrid;
    procedure AcNumFormatCustomGetNumberFormatString(Sender: TObject;
      AWorkbook: TsWorkbook; var ANumFormatStr: String);
    procedure BtnPageSUClick(Sender: TObject);
    procedure CBFunctionsChange(Sender: TObject);
    procedure ColorComboboxAddColors(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure MnuOpenClick(Sender: TObject);
    procedure MnuSaveClick(Sender: TObject);
    procedure MnuTestReportClick(Sender: TObject);
    procedure MnuNewClick(Sender: TObject);
    procedure Panel1Click(Sender: TObject);
    procedure TBInsColClick(Sender: TObject);
    procedure TBInsRowClick(Sender: TObject);
    procedure TBDelColClick(Sender: TObject);
    procedure TBDelRowClick(Sender: TObject);
    procedure WorksheetGridClick(Sender: TObject);

  private
    { private declarations }

  public
    procedure LoadFile(const AFileName: string);
    procedure UpdateFormatProperties(AFormatIndex: integer);
    procedure SaveSS;
    procedure SaveQMReportDef;
    procedure ShowPageInfo;
    function CalcSSWidth: single;
    function CalcSSHeight: single;
    function SetColPageBreak: boolean;

  end;
// define needed QMCLient functions not currently found in tisQMClient
function QMCallx(subrname: PAnsiChar; argc: integer;
  a1, a2, a3, a4, a5, a6, a7: PAnsiChar): integer;
  cdecl; external 'qmclilib.dll' Name 'QMCallx';
function QmGetArg(ArgNo: integer): PAnsiChar; cdecl;
  external 'qmclilib.dll' Name 'QMGetArg';
procedure QMFree(p: PAnsiChar); cdecl; external 'qmclilib.dll' Name 'QMFree';

const
  INT_QUERY_FUNCTION = 998;
  INT_LOOKUP_FUNCTION = 999;

  // for prtsettings
  constPageSu = 0;
  constNewSS  = 1;

  // paper sizes in mm
  PaperWidth: array [0..7] of double = (215.9, 215.9, 279.4, 841, 594, 420, 297, 210);
  PaperHeight: array [0..7] of double = (279.4, 355.6, 431.8, 1189, 841, 594, 420, 297);

  PAGE_CELL_BORDER: TsCellBorderStyle = (LineStyle: lsThick; Color: clBlue);
  PAGE_DOTTED_CELL_BORDER: TsCellBorderStyle = (LineStyle: lsDotted; Color: clBlue);


  // need to prompt for the following!!!
  QMAccount = 'NEXTJOB';
  VerInfo = '.01b';

  //QM_Client Stuff
  QM_server = '127.0.0.1';
  QM_port = -1;

  //  qm field delimiters

  QM_AM = #254;  // attribute or field mark
  QM_VM = #253;  // value or item mark
  QM_SVM = #252;  // subvalue mark

  //* QM Server error status values */
  SV_OK = 0;   //* Action successful                       */
  SV_ON_ERROR = 1;   //* Action took ON ERROR clause             */
  SV_ELSE = 2;   //* Action took ELSE clause                 */
  SV_ERROR = 3;   //* Action failed. Error text available     */
  SV_LOCKED = 4;   //* Action took LOCKED clause               */
  SV_PROMPT = 5;   //* Server requesting input                 */

var
  MainForm: TMainForm;
  DefinitionFile: string;  // Golbal var full path and name of Definition File
  BaseFileName: string;    // file name and path of opened file less extention
  TestData: string;       // Json String to send NxtReportGen when testing report
  prtsettingMode : integer; // Entry mode for prtsettings constPageSu - page settings only or constNewSS ;

implementation

uses
  fpcanvas, lazutf8, utestreport, uprtsettings, umvquerydlg, umvlookupdlg,
  umvreplacedlg, snumformatform;



// REMEMBER !!! WorksheetGrid.ReadFormulas := True!  and workbooksource option boAutoCalc true

// Dummy mvQuery Function to all save of query text in cell record
//RegisterFunction('mvQUERY', 'S', 'SSSS', INT_QUERY_FUNCTION, @fpsQUERY);
//Define mvQuery: 5 parameters =mvQuery(File_Name, Sort_By, Qualifiers, Data_Fields, Display_Field)
//  File_Name - QM File Name to generate report from  ie CUSTOMERS
//  Sort_By     - Sort Criteria    ie by cm.name
//  Qualifiers  - Selection Clause  ie with cm.status='ACTIVE'
//  Data_Fields - Field Names to report as defined in the QM File Dictionary ie  cm.name, cm.address, cm.phone
//   Display_Field - Field Name to report in this column ie cm.name

function MyQMGetArg(ArgNo: integer): ansistring;
  // retrieves the value of argument variables to QMCall(), QMCallx()
  // Rem use MyGetArg to free memory allocated by QMGetArg -- see QM Doc
var
  s: PAnsiChar;
begin
  s := QmGetArg(ArgNo);
  MyQMGetArg := PAnsiChar(s);
  QMFree(s);
end;

procedure fpsQUERY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY QUERY FUNCTION
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
  end;
  // "ArgToString" simplifies getting the string from a TsExpressionResult
  // rem arguments are parsed into Args Array ie
  // =mvquery(arg0,arg1,arg2.....)
  s := ExtractWord(1, ArgToString(Args[3]), [' ']);
  //  s := 'mvquery: ' + S;
  Result := StringResult(s);
  // "StringResult" stores the string s in the ResString field of the
  // TsExpressionResult and sets the ResultType to rtString.
  // There is such a function for each basic data type.
end;

procedure fpsREPORT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY Report FUNCTION
var
  s: string;

begin
  s := '';
  if Args[0].ResultType = rtError then
  begin
    Result := ErrorResult(Args[0].ResError);
    exit;
  end;
  s := ArgToString(Args[0]);

  Result := StringResult(s);

end;

// Dummy mvLOOKUP Function to all save of LOOKUP text in cell record
//RegisterFunction('mvLOOKUP', 'S', 'SSS', INT_LOOKUP_FUNCTION, @fpsLOOKUP);
procedure fpsLOOKUP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// DUMMY LOOKUP FUNCTION
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
    s := s + ArgToString(Args[i]) + ', ';
  end;
  Result := StringResult(s);

end;
// Dummy mvREPLACE Function to all save of REPLACE text in cell record
//RegisterFunction('mvReplace', 'S', 'SS', INT_LOOKUP_FUNCTION, @fpsREPLACE);
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

{ TMainForm }

function TMainForm.CalcSSWidth: single;
var
  i: integer;
begin
  CalcSSWidth := 0;
// WorkSheetGrid.ColCount gives us the actual grid col count including the row count col on the left border
// fpspreadsheet does not count this as a col, col 'A' is the first col ref by index 0
//  GetLastColIndex  Returns the 0-based index of the last column containing a cell  with a
//  CELL RECORD (due to content or formatting).
//  This means there can be cols visible in the worksheet grid that have no associated cells!
//  Resulting in this size calc being incorrect if we use WorkSheetGrid.Worksheet.GetLastColIndex
//  so use the count for the grid instead :  WorkSheetGrid.ColCount -2
  for i := 0 to WorkSheetGrid.ColCount -2 do
    CalcSSWidth := CalcSSWidth + WorkSheetGrid.Worksheet.GetColWidth(i, suMillimeters);
end;

function TMainForm.CalcSSHeight: single;
var
  i: integer;
begin
  CalcSSHeight := 0;
  for i := 0 to WorkSheetGrid.Worksheet.GetLastRowIndex do
    CalcSSHeight := CalcSSHeight + WorkSheetGrid.Worksheet.GetRowHeight(i, suMillimeters);
end;

function TMainForm.SetColPageBreak: boolean;
//
//   to do:
//     save last col location of pagebreak (to clear if it exits)
//     pagebreak is an option of the cells,
//     a better way to do this may be with cell formating
//     see
//

var
  i : integer;
  SSpageWidth, SSWidth: single;
  sheet: TsWorksheet;
  c: cardinal;
begin
  SSWidth := 0;
  SetColPageBreak := False;
  i :=  FrmPrtSettings.LBPaper.ItemIndex;
  SSpageWidth := PaperWidth[FrmPrtSettings.LBPaper.ItemIndex];

  // Search for and remove previous Page Breaks  (this works because if there is a page break set there is a Cell Reord)
  for i := 0 to WorkSheetGrid.Worksheet.GetLastColIndex  do
    If WorksheetGrid.Worksheet.IsPageBreakCol(i) then
      WorksheetGrid.Worksheet.RemovePageBreakFromCol(i);


  if SSpageWidth > 20 then     // need some sort of test to see if page width is assigned
//    for i := 0 to LastColidx  do
  for i:= 0 to WorkSheetGrid.ColCount -2 do
    begin
      SSWidth := SSWidth + WorkSheetGrid.Worksheet.GetColWidth(i, suMillimeters);
      if (SSWidth > SSpageWidth) then
      begin
        sheet := WorksheetGrid.Worksheet;
        if sheet <> nil then
        begin
          c := WorksheetGrid.GetWorksheetCol(i+1);
          sheet.AddPageBreakToCol(c);
          WorksheetGrid.Invalidate;
          break;
        end;
      end;
    end;
end;



procedure TMainForm.ShowPageInfo;
var
  size: single;
begin
  // ss width
  size := CalcSSWidth();
  lblSSWidth.Caption := 'SS Width: ' + FloatToStrF(Size, fffixed, 0, 2);
  lblSSWidth.Caption := lblSSWidth.Caption + 'mm ' + FloatToStrF(
    (size * 0.0393701), fffixed, 0, 2) + 'in';
  // ss height
  size := CalcSSHeight();
  lblSSHeight.Caption := 'SS Height: ' + FloatToStrF(Size, fffixed, 0, 2);
  lblSSHeight.Caption := lblSSHeight.Caption + 'mm ' + FloatToStrF(
    (size * 0.0393701), fffixed, 0, 2) + 'in';
  // page width
  size := PaperWidth[FrmPrtSettings.LBPaper.ItemIndex];
  lblPageWidth.Caption := 'Page Width: ' + FloatToStrF(Size, fffixed, 0, 2);
  lblPageWidth.Caption := lblPageWidth.Caption + 'mm ' + FloatToStrF(
    (size * 0.0393701), fffixed, 0, 2) + 'in';
  // page height
  size := PaperHeight[FrmPrtSettings.LBPaper.ItemIndex];
  lblPageHeight.Caption := 'Page Height: ' + FloatToStrF(Size, fffixed, 0, 2);
  lblPageHeight.Caption := lblPageHeight.Caption + 'mm ' + FloatToStrF(
    (size * 0.0393701), fffixed, 0, 2) + 'in';
  // col infor  rem  WorkSheetGrid.ColCount includes the row number col on the left boarder of the grid, fpspreadsheet does not
  // count this as a cell!
  lblColInfo.Caption :=   'LastCol idx/nbr: ' +  IntToStr(WorkSheetGrid.Worksheet.GetLastColIndex) + ' ' + IntToStr(WorkSheetGrid.ColCount -1);
end;

procedure TMainForm.BtnPageSUClick(Sender: TObject);
begin
  prtsettingMode := constPageSu;
  FrmPrtSettings.showmodal;
  ShowPageInfo;
  SetColPageBreak;
end;

procedure TMainForm.AcNumFormatCustomGetNumberFormatString(Sender: TObject;
  AWorkbook: TsWorkbook; var ANumFormatStr: String);
var
  F: TNumFormatForm;
  sample: Double;
begin
  Unused(AWorkbook);
  F := TNumFormatForm.Create(nil);
  try
    F.Position := poMainFormCenter;
    with sWorkbookSource1.Worksheet do
      sample := ReadAsNumber(ActiveCellRow, ActiveCellCol);
    F.SetData(ANumFormatStr, sWorkbookSource1.Workbook, sample);
    if F.ShowModal = mrOK then
      ANumFormatStr := F.NumFormatStr;
  finally
    F.Free;
  end;
end;

procedure TMainForm.CBFunctionsChange(Sender: TObject);
var
  i: integer;
begin
  i := CBFunctions.ItemIndex;
  case i of
    0: mvquerydlg.Show;
    //   0 : if mvquerydlg.showmodel <> MrOk then
    //         showmessage('Selection: Cancel');   // mvquery function selected
    1: mvreplacedlg.Show;
    2: mvlookupdlg.Show;
    else
      ShowMessage('Selection: ' + IntToStr(i) + ' ' + CBFunctions.Items[i] +
        ' Not Currently Supported');
  end;
end;

procedure TMainForm.ColorComboboxAddColors(Sender: TObject);
begin

  with TsCellCombobox(Sender) do
  begin
    // These are the Excel-8 palette colors, a bit rearranged and without the
    // duplicates.
    AddColor($000000, 'black');
    AddColor($333333, 'gray 80%');
    AddColor($808080, 'gray 50%');
    AddColor($969696, 'gray 40%');
    AddColor($C0C0C0, 'silver');
    AddColor($FFFFFF, 'white');
    AddColor($FF0000, 'red');
    AddColor($00FF00, 'green');
    AddColor($0000FF, 'blue');
    AddColor($FFFF00, 'yellow');
    AddColor($FF00FF, 'magenta');
    AddColor($00FFFF, 'cyan');

    AddColor($800000, 'dark red');
    AddColor($008000, 'dark green');
    AddColor($000080, 'dark blue');
    AddColor($808000, 'olive');
    AddColor($800080, 'purple');
    AddColor($008080, 'teal');
    AddColor($9999FF, 'periwinkle');
    AddColor($993366, 'plum');
    AddColor($FFFFCC, 'ivory');
    AddColor($CCFFFF, 'light turquoise');
    AddColor($660066, 'dark purple');
    AddColor($FF8080, 'coral');
    AddColor($0066CC, 'ocean blue');
    AddColor($CCCCFF, 'ice blue');

    AddColor($00CCFF, 'sky blue');
    AddColor($CCFFCC, 'light green');
    AddColor($FFFF99, 'light yellow');
    AddColor($99CCFF, 'pale blue');
    AddColor($FF99CC, 'rose');
    AddColor($CC99FF, 'lavander');
    AddColor($FFCC99, 'tan');

    AddColor($3366FF, 'light blue');
    AddColor($33CCCC, 'aqua');
    AddColor($99CC00, 'lime');
    AddColor($FFCC00, 'gold');
    AddColor($FF9900, 'light orange');
    AddColor($FF6600, 'orange');
    AddColor($666699, 'blue gray');
    AddColor($003366, 'dark teal');
    AddColor($339966, 'sea green');
    AddColor($003300, 'very dark green');
    AddColor($333300, 'olive green');
    AddColor($993300, 'brown');
    AddColor($333399, 'indigo');
  end;
end;

procedure TMainForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if QMConnected then
    QMDisconnect();
end;

procedure TMainForm.SaveQMReportDef;
// Parse Grid and create QmReportDef Record (as CSV file)

//  need to find easy way to write csv file
//  could use another grid to save info in and write but seems like overkill

var
  Acell: PCell;
  CellAddr: string;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  NxtReportGen: string;
  CellFormulaText, TokenText: string;
  r, i: integer;
begin
  TestData := '';   // init test json string
  // Attempt to find NxtReportGen exe
  // should be in same directory as NxtDesinger
  NxtReportGen := ExtractFilePath(ParamStr(0)) + 'NxtReportGen.exe';
  if not (FileExists(NxtReportGen)) then
    MessageDlg('NxtReportGen Not Found in Assumed Location: ' +
      NxtReportGen + ' QM Subroutine NxtReport Will Error If Not Corrected!',
      mtInformation, [mbOK], 0);
  NxtReportGen := ExtractFilePath(ParamStr(0)) + 'NxtRunCommand.exe';
  if not (FileExists(NxtReportGen)) then
    MessageDlg('NxtRunCommand Not Found in Assumed Location: ' +
      NxtReportGen + ' QM Subroutine NxtReport Will Error If Not Corrected!',
      mtInformation, [mbOK], 0);
  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Sheet 1');
    r := 0; // row counter
    MyWorksheet.WriteText(r, 0, SaveDialog.FileName);  // save name of format file
    MyWorksheet.WriteText(r, 1, QMAccount);
    MyWorksheet.WriteText(r, 2, ExtractFilePath(ParamStr(0)));
    Inc(r);
    for Acell in WorksheetGrid.Worksheet.Cells do
    begin
      CellAddr := GetCellString(Acell^.Row, Acell^.Col, []);
      // Is there a formula entered in the cell ?
      //     if Length(Acell^.FormulaValue) > 0 then
      if cfHasFormula in Acell^.Flags then
      begin
        MyWorksheet.WriteText(r, 0, CellAddr);
        MyWorksheet.WriteText(r, 1, 'Q');
        //        MyWorksheet.WriteText(r, 2, Acell^.FormulaValue);
        CellFormulaText := WorkSheetGrid.Worksheet.ReadFormulaAsString(Acell, True);
        MyWorkSheet.WriteText(r, 2, CellFormulaText);
        // if this is a MVREPLACE function, save TOKEN for use in report testing
        if Pos('MVREPLACE', CellFormulaText) > 0 then
        begin
          i := Pos('%%', CellFormulaText);
          if i > 0 then
          begin
            i := i + 2;
            TokenTExt := UpperCase(ExtractSubStr(CellFormulaText, i, ['%']));
            if length(TestData) > 0 then
              TestData := TestData + ',';
            TestData := TestData + '"%%' + TokenText + '%%":"' + TokenText + 'Value"';
          end;
        end;

        Inc(r);
      end
      // Is there text entered in the cell ?
      else if Length(Acell^.UTF8StringValue) > 0 then
      begin
        MyWorksheet.WriteText(r, 0, CellAddr);
        MyWorksheet.WriteText(r, 1, 'T');
        MyWorksheet.WriteText(r, 2, Acell^.UTF8StringValue);
        Inc(r);
      end;
    end;

    CSVParams.Delimiter := ',';
    //    CSVParams.Delimiter := #9;   // tab
    CSVParams.FormatSettings.DecimalSeparator := '.';
    CSVParams.NumberFormat := '%.9f';
    CSVParams.QuoteChar := '"';

    // Save the spreadsheet to a file
    DefinitionFile := ChangeFileExt(SaveDialog.FileName, '.dat');
    MyWorkbook.WriteToFile(DefinitionFile, sfCSV, True);
    //    MessageDlg('Rem QM Report Def File has QM Account Name Currently Hardcoded: ' + QMAccount, mtInformation, [mbOK], 0);
  finally
    MyWorkbook.Free;
  end;
end;

// Saves sheet in grid to file, overwriting existing file
procedure TMainForm.SaveSS;
var
  err, fn: string;
  Row1, Row2: integer;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  if WorksheetGrid.Workbook.Filename <> '' then
  begin
    fn := AnsiToUTF8(WorksheetGrid.Workbook.Filename);
    SaveDialog.InitialDir := ExtractFileDir(fn);
    SaveDialog.FileName := ChangeFileExt(ExtractFileName(fn), '');
  end;

  // Save Page Layout Settings
  if FrmPrtSettings.RBPortrait.Checked then
    WorksheetGrid.Worksheet.PageLayout.Orientation := spoPortrait
  else
    WorksheetGrid.Worksheet.PageLayout.Orientation := spoLandscape;
  // Paper Size
  WorksheetGrid.Worksheet.PageLayout.PageWidth :=
    PaperWidth[FrmPrtSettings.LBPaper.ItemIndex];
  WorksheetGrid.Worksheet.PageLayout.PageHeight :=
    PaperHeight[FrmPrtSettings.LBPaper.ItemIndex];

  // Repeated rows
  try
    Row1 := StrToInt(FrmPrtSettings.EdtRow1.Text);
    Row2 := StrToInt(FrmPrtSettings.EdtRow2.Text);
    if Row1 > 0 then
    begin
      Row1 := Row1 - 1;  // rem Arow is 0 based but grid start at 1

      if Row2 > 0 then
      begin
        Row2 := Row2 - 1;
        WorksheetGrid.Worksheet.PageLayout.SetRepeatedRows(Row1, Row2);
      end
      else
        WorksheetGrid.Worksheet.PageLayout.SetRepeatedRows(Row1, Row1);
    end
    else
      WorksheetGrid.Worksheet.PageLayout.SetRepeatedRows(
        UNASSIGNED_ROW_COL_INDEX, UNASSIGNED_ROW_COL_INDEX);

  except
    FrmPrtSettings.EdtRow1.Text := '0';
    FrmPrtSettings.EdtRow2.Text := '0';
  end;

  // set margins
  try
    WorksheetGrid.Worksheet.PageLayout.TopMargin :=
      StrToFloat(FrmPrtSettings.EdtTopMargin.Text);
  except
    WorksheetGrid.Worksheet.PageLayout.TopMargin := 20;
  end;
  try
    WorksheetGrid.Worksheet.PageLayout.BottomMargin :=
      StrToFloat(FrmPrtSettings.EdtBotMargin.Text);
  except
    WorksheetGrid.Worksheet.PageLayout.BottomMargin := 20;
  end;
  try
    WorksheetGrid.Worksheet.PageLayout.LeftMargin :=
      StrToFloat(FrmPrtSettings.EdtLeftMargin.Text);
  except
    WorksheetGrid.Worksheet.PageLayout.LeftMargin := 17.78;
  end;
  try
    WorksheetGrid.Worksheet.PageLayout.RightMargin :=
      StrToFloat(FrmPrtSettings.EdtRightMargin.Text);
  except
    WorksheetGrid.Worksheet.PageLayout.RightMargin := 17.78;
  end;

  // copies
  try
    WorksheetGrid.Worksheet.PageLayout.Copies := StrToInt(FrmPrtSettings.EdtCopies.Text);
  except
    WorksheetGrid.Worksheet.PageLayout.Copies := 1;
  end;

  // options
  WorksheetGrid.Worksheet.PageLayout.Options := [];

  if FrmPrtSettings.CBGridLines.Checked then
    WorksheetGrid.Worksheet.PageLayout.Options :=
      WorksheetGrid.Worksheet.PageLayout.Options + [poPrintGridLines];

  if FrmPrtSettings.CBHeaders.Checked then
    WorksheetGrid.Worksheet.PageLayout.Options :=
      WorksheetGrid.Worksheet.PageLayout.Options + [poPrintHeaders];

  if FrmPrtSettings.CBFitToPage.Checked then
  begin
    WorksheetGrid.Worksheet.PageLayout.Options :=
      WorksheetGrid.Worksheet.PageLayout.Options + [poFitPages];
    WorksheetGrid.Worksheet.PageLayout.FitWidthToPages := 1;
    // all columns on one page width
    WorksheetGrid.Worksheet.PageLayout.FitHeightToPages := 0;
    // use as many pages as needed
  end;
  // save it?

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      WorksheetGrid.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));
      //WorksheetGrid.WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));     // works as well
    finally
      Screen.Cursor := crDefault;
      // Show a message in case of error(s)
      err := WorksheetGrid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);

begin
  RegisterFunction('MVQUERY', 'S', 'SSSS', INT_QUERY_FUNCTION, @fpsQUERY);
  RegisterFunction('MVLOOKUP', 'S', 'SSS', INT_QUERY_FUNCTION, @fpsLOOKUP);
  RegisterFunction('MVREPORT', 'S', 'S', INT_QUERY_FUNCTION, @fpsREPORT);
  RegisterFunction('MVREPLACE', 'S', 'SS', INT_QUERY_FUNCTION, @fpsREPLACE);
  MainForm.Caption := 'NxtDesigner Ver: ' + VerInfo;
  TestData := ''; // Init Test Data String, used for testing report
  prtsettingMode := constPageSu;
end;

procedure TMainForm.MnuNewClick(Sender: TObject);

begin
  prtsettingMode := constNewSS;
  if FrmPrtSettings.ShowModal = mrOk then
    begin
      WorksheetGrid.NewWorkbook(FrmPrtSettings.edCols.Value, FrmPrtSettings.edRows.Value);
      ShowPageInfo;
      SetColPageBreak;
      WorksheetGrid.Options := WorksheetGrid.Options + [goEditing];
      WorksheetGrid.Visible := True;
    end;
 end;

procedure TMainForm.Panel1Click(Sender: TObject);
begin
    ShowPageInfo;
end;

procedure TMainForm.TBInsColClick(Sender: TObject);
begin
  WorksheetGrid.InsertCol(WorksheetGrid.Col);
  WorksheetGrid.Col := WorksheetGrid.Col + 1;
  ShowPageInfo;
end;

procedure TMainForm.TBInsRowClick(Sender: TObject);
begin
  WorksheetGrid.InsertRow(WorksheetGrid.Row);
  WorksheetGrid.Row := WorksheetGrid.Row + 1;
  ShowPageInfo;
end;

procedure TMainForm.TBDelColClick(Sender: TObject);
var
  c: integer;
begin
  c := WorksheetGrid.Col;
  WorksheetGrid.DeleteCol(c);
  WorksheetGrid.Col := c;
  ShowPageInfo;
end;

procedure TMainForm.TBDelRowClick(Sender: TObject);
var
  r: integer;
begin
  r := WorksheetGrid.Row;
  WorksheetGrid.DeleteRow(r);
  WorksheetGrid.Row := r;
  ShowPageInfo;
end;

procedure TMainForm.WorksheetGridClick(Sender: TObject);
  Var
//    Celltext : string;
    ACell : PCell;
begin

//  Celltext := FloatToStrF(WorkSheetGrid.Worksheet.GetColWidth(WorksheetGrid.Col - 1, suMillimeters),fffixed, 0, 2);
//  WorksheetGrid.Worksheet.WriteText(WorksheetGrid.Worksheet.ActiveCellRow,WorksheetGrid.Worksheet.ActiveCellCol,celltext);

  lblCellWidth.Caption := 'Cell Width: ' +
    FloatToStrF(WorkSheetGrid.Worksheet.GetColWidth(WorksheetGrid.Col - 1, suMillimeters),
    fffixed, 0, 2);
  lblCellWidth.Caption := lblCellWidth.Caption + 'mm ' +
    FloatToStrF(WorkSheetGrid.Worksheet.GetColWidth(WorksheetGrid.Col - 1, suInches),
    fffixed, 0, 2) + 'in';
  lblCellHeight.Caption := 'Cell Height: ' +
    FloatToStrF(WorkSheetGrid.Worksheet.GetRowHeight(WorksheetGrid.Row - 1, suMillimeters),
    fffixed, 0, 2);
  lblCellHeight.Caption := lblCellHeight.Caption + 'mm ' +
    FloatToStrF(WorkSheetGrid.Worksheet.GetRowHeight(WorksheetGrid.Row - 1, suInches),
    fffixed, 0, 2) + 'in';

 {
  -- example code on how to find format tpye and string as displayed in TsSpreadhseet Inspector
  -- from fpspreadsheetctrls  UpdateFormatProperties
  }
  ACell := WorkSheetGrid.Worksheet.FindCell(WorksheetGrid.Worksheet.ActiveCellRow,WorksheetGrid.Worksheet.ActiveCellCol);
  if ACell <> nil then
    UpdateFormatProperties(ACell^.FormatIndex)
  else
    UpdateFormatProperties(-1);
end;

procedure TMainForm.UpdateFormatProperties(AFormatIndex: integer);
var
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;

begin
  if AFormatIndex > -1 then
    fmt := WorkSheetGrid.Workbook.GetCellFormat(AFormatIndex)
  else
    InitFormatRecord(fmt);

  if (AFormatIndex = -1) or not (uffNumberFormat in fmt.UsedFormattingFields) then
  begin
    edFmtType.Text := '(default)';
    edFmtStr.Text := '(none)';
  end else
  begin
    numFmt := WorkSheetGrid.Workbook.GetNumberFormat(fmt.NumberFormatIndex);
    edFmtType.Text :=Format('%s', [
      GetEnumName(TypeInfo(TsNumberFormat), ord(numFmt.NumFormat))]);
    edFmtStr.Text := numFmt.NumFormatStr;
  end;
end;


// Loads first worksheet from file into grid
procedure TMainForm.LoadFile(const AFileName: string);
var
  err: string;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));
      //      WorksheetGrid.Workbook.ReadFromFile(AFileName, sfOpenDocument);
      // Update user interface
      Caption := Format('fpsGrid - %s (%s)', [AFilename,
        GetSpreadTechnicalName(WorksheetGrid.Workbook.FileFormatID)]);

      // Collect the sheet names in the combobox for switching sheets.
    except
      on E: Exception do
      begin
        // Empty worksheet instead of the loaded one
        WorksheetGrid.NewWorkbook(7, 10);
        Caption := 'fpsGrid - no name';
        // Grab the error message
        WorksheetGrid.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := WorksheetGrid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TMainForm.MnuOpenClick(Sender: TObject);
var
  i: integer;
  RepeatRange: TsRowColRange;

begin
  if OpenDialog.FileName <> '' then
  begin
    OpenDialog.InitialDir := ExtractFileDir(OpenDialog.FileName);
    OpenDialog.FileName := ChangeFileExt(ExtractFileName(OpenDialog.FileName), '');
  end;
  if OpenDialog.Execute then
  begin
    LoadFile(OpenDialog.FileName);
    // load Page Settings
    if WorksheetGrid.Worksheet.PageLayout.Orientation = spoPortrait then
      FrmPrtSettings.RBPortrait.Checked := True
    else
      FrmPrtSettings.RBLandScape.Checked := True;

    FrmPrtSettings.EdtCopies.Text := IntToStr(WorksheetGrid.Worksheet.PageLayout.Copies);

    // set options
    FrmPrtSettings.CBGridLines.Checked := False;
    FrmPrtSettings.CBHeaders.Checked := False;
    FrmPrtSettings.CBFitToPage.Checked := False;

    if [poPrintGridLines] <= WorksheetGrid.Worksheet.PageLayout.Options then
      FrmPrtSettings.CBGridLines.Checked := True;

    if [poPrintHeaders] <= WorksheetGrid.Worksheet.PageLayout.Options then
      FrmPrtSettings.CBHeaders.Checked := True;

    if [poFitPages] <= WorksheetGrid.Worksheet.PageLayout.Options then
      FrmPrtSettings.CBFitToPage.Checked := True;
    // find and set page size
    for i := 0 to 7 do
      if Trunc(WorksheetGrid.Worksheet.PageLayout.PageHeight) = Trunc(PaperHeight[i]) then
      begin
        FrmPrtSettings.LBPaper.ItemIndex := i;
        Break;
      end;

    // set margins
    FrmPrtSettings.EdtTopMargin.Text :=
      FloatToStr(WorksheetGrid.Worksheet.PageLayout.TopMargin);
    FrmPrtSettings.EdtBotMargin.Text :=
      FloatToStr(WorksheetGrid.Worksheet.PageLayout.BottomMargin);
    FrmPrtSettings.EdtLeftMargin.Text :=
      FloatToStr(WorksheetGrid.Worksheet.PageLayout.LeftMargin);
    FrmPrtSettings.EdtRightMargin.Text :=
      FloatToStr(WorksheetGrid.Worksheet.PageLayout.RightMargin);
    // set rows to repeat

    if WorksheetGrid.Worksheet.PageLayout.HasRepeatedRows then
    begin
      RepeatRange := WorksheetGrid.Worksheet.PageLayout.RepeatedRows;
      FrmPrtSettings.EdtRow1.Text := IntToStr(RepeatRange.FirstIndex + 1);
      FrmPrtSettings.EdtRow2.Text := IntToStr(RepeatRange.LastIndex + 1);
    end
    else
    begin
      FrmPrtSettings.EdtRow1.Text := '0';
      FrmPrtSettings.EdtRow2.Text := '0';
    end;
//    LastPageBorder := -1;
    WorksheetGrid.Options := WorksheetGrid.Options + [goEditing];
    WorksheetGrid.Visible := True;
    ShowPageInfo;
  end;
end;

procedure TMainForm.MnuSaveClick(Sender: TObject);
begin
  SaveSS;
  SaveQMReportDef;
end;

procedure TMainForm.MnuTestReportClick(Sender: TObject);
begin
  FrmTestReport.showmodal;
end;

initialization
  {$I umainform.lrs}

end.
