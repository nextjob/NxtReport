unit uprtsettings;
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
}
{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ExtCtrls, Spin, fpstypes;

type

  { TFrmPrtSettings }

  TFrmPrtSettings = class(TForm)
    Button1: TButton;
    CBGridLines: TCheckBox;
    CBHeaders: TCheckBox;
    CBFitToPage: TCheckBox;
    EdtRightMargin: TEdit;
    EdtLeftMargin: TEdit;
    EdtBotMargin: TEdit;
    EdtTopMargin: TEdit;
    EdtRow2: TEdit;
    EdtRow1: TEdit;
    EdtCopies: TEdit;
    GBOrientation: TGroupBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    gbNewReport: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    LBPaper: TListBox;
    RBLandScape: TRadioButton;
    RBPortrait: TRadioButton;
    edCols: TSpinEdit;
    edRows: TSpinEdit;
    procedure EdtKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
    procedure LBPaperClick(Sender: TObject);

  private
    { private declarations }
  public
    { public declarations }
    procedure setColRowsz;
  end;

var
  FrmPrtSettings: TFrmPrtSettings;

implementation
uses
    umainform;
{ TFrmPrtSettings }


procedure TFrmPrtSettings.EdtKeyPress(Sender: TObject; var Key: char);
begin
  if not(Key IN ['0'..'9', #8, #9, #13, #27, #127]) then key:= #0;
end;

procedure TFrmPrtSettings.FormShow(Sender: TObject);
begin
  If umainform.prtsettingMode = umainform.constNewSS then
    begin
    // calculate row and col count based on current page size settings
      setColRowsz;
      gbNewReport.Visible := True
    end
  else
    gbNewReport.Visible := False;
end;

procedure TFrmPrtSettings.LBPaperClick(Sender: TObject);
begin
  // calculate row and col count based on current page size settings
  setColRowsz;
end;

procedure TFrmPrtSettings.setColRowsz;
Var
  SSpageWidth, SSpageHeight : single;
begin
  // calculate row and col count based on current page size settings
  SSpageWidth := umainform.PaperWidth[LBPaper.ItemIndex];
  SSpageHeight := umainform.PaperHeight[LBPaper.ItemIndex];
  edCols.Value := Trunc((SSpageWidth / MainForm.WorkSheetGrid.Worksheet.ReadDefaultColWidth(suMillimeters)) + 0.5);
  edRows.Value := Trunc((SSpageHeight / MainForm.WorkSheetGrid.Worksheet.ReadDefaultRowHeight(suMillimeters)) + 0.5);
end;

initialization
  {$I uprtsettings.lrs}

end.

