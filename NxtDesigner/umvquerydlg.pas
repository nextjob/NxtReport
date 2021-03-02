unit umvquerydlg;
//
// routine performs mvquery function entry into spreadsheet grid cell
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
  StdCtrls, ComCtrls;

type

  { TMVQueryDlg }

  TMVQueryDlg = class(TForm)
    btnAddField: TButton;
    CancelBtn: TButton;
    AndOrCB: TComboBox;
    FieldsMemo: TMemo;
    SelectFieldEdt2: TEdit;
    SelectValueEdt1: TEdit;
    SelectValueEdt2: TEdit;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    withCB1: TComboBox;
    SortByCB: TComboBox;
    OkBtn: TButton;
    Label4: TLabel;
    SelectFieldEdt1: TEdit;
    Label3: TLabel;
    SortByEdt: TEdit;
    FileNameEdt: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    withCB2: TComboBox;
    procedure btnAddFieldClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure FileNameEdtClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure OkBtnClick(Sender: TObject);
    procedure SelectFieldEdt1Click(Sender: TObject);
    procedure SelectFieldEdt2Click(Sender: TObject);
    procedure SortByEdtClick(Sender: TObject);
  private
    procedure FillEditText(Sender: TEdit);
    { private declarations }
  public
    { public declarations }
  end;

var
  MVQueryDlg: TMVQueryDlg;

implementation

uses
  umainform, udictselect;

{ TMVQueryDlg }

 procedure TMVQueryDlg.FillEditText(Sender: TEdit);
 var Item : TListItem;
 begin
    if FrmDictSelect.Listview1.SelCount > 0 then
    begin
      Item := FrmDictSelect.Listview1.selected;
      Sender.text := Item.Caption;
    end;
 end;

procedure TMVQueryDlg.OkBtnClick(Sender: TObject);
var
  i : integer;
  errmsg : string;
  celltext : string;
  displayflds : string;
  CloseAction: TCloseAction;
begin
  If length(FileNameEdt.Text) = 0 then
    errmsg :='File Name is missing';
  If FieldsMemo.Lines.Count = 0 then
    errmsg :='No Display Fields';

  If length(errmsg) > 0 then begin
    errmsg := errmsg + ', Please Correct';
    MessageDlg(errmsg, mtError, [mbOK], 0);
    ModalResult := mrNone;
  end
  else begin
// parse the form controls to get text to enter into the active cell
    celltext :=  '"'+ FileNameEdt.Text + '"';
//  sort by?
    if length(SortByEdt.Text) > 0 then
      celltext := celltext + ',"' + SortByCB.Items[SortByCB.ItemIndex] + ' ' + SortByEdt.Text + '"'
    else
      celltext := celltext + ',';
//  selections
    if (length(SelectFieldEdt1.text) > 0) and (length(SelectValueEdt1.text) > 0) then
      If (length(SelectFieldEdt2.text) > 0) and (length(SelectValueEdt2.text) > 0) then
        celltext := celltext + '," WITH ' + SelectFieldEdt1.text + ' ' + withCB1.Items[withCB1.ItemIndex] + ' ' + SelectValueEdt1.text + ' ' + AndOrCB.Items[AndOrCB.ItemIndex] + ' ' + SelectFieldEdt2.text + ' ' + withCB2.Items[withCB2.ItemIndex] + ' ' + SelectValueEdt2.text + '"'
      else
        celltext := celltext + '," WITH ' + SelectFieldEdt1.text + ' ' + withCB1.Items[withCB1.ItemIndex] + ' ' + SelectValueEdt1.text + '"'
    else
      celltext := celltext + ',';
//  display fields
    displayflds := '';
    for i := 0 to FieldsMemo.Lines.Count -1 do
      displayflds := displayflds + FieldsMemo.Lines[i] + ' ';
    celltext := celltext + ',"' + displayflds + '"';
    //  enter the mvquery text into the currently active cell
    celltext := '=MVQUERY('+celltext+')';
    MainForm.WorksheetGrid.Worksheet.WriteFormula(MainForm.WorksheetGrid.Worksheet.ActiveCellRow,MainForm.WorksheetGrid.Worksheet.ActiveCellCol,celltext);
    // enter mvreport into adjacent cells
    for i := 1 to FieldsMemo.Lines.Count -1 do
      begin
        celltext := '=MVREPORT("' + FieldsMemo.Lines[i] + '")';
        MainForm.WorksheetGrid.Worksheet.WriteFormula(MainForm.WorksheetGrid.Worksheet.ActiveCellRow,MainForm.WorksheetGrid.Worksheet.ActiveCellCol + i,celltext);
      end;
    CloseAction := caHide;
    FormClose(self,CloseAction);
  end;



end;

procedure TMVQueryDlg.SelectFieldEdt1Click(Sender: TObject);
begin
  FillEditText(SelectFieldEdt1);
end;

procedure TMVQueryDlg.SelectFieldEdt2Click(Sender: TObject);
begin
  FillEditText(SelectFieldEdt2);
end;

procedure TMVQueryDlg.SortByEdtClick(Sender: TObject);

begin
   FillEditText(SortByEdt);
end;

procedure TMVQueryDlg.FormShow(Sender: TObject);
begin
  FrmDictSelect.Show;
end;

procedure TMVQueryDlg.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
   FrmDictSelect.Hide;
   mvquerydlg.Hide;
end;

procedure TMVQueryDlg.FileNameEdtClick(Sender: TObject);
begin
     If FrmDictSelect.LBFiles.ItemIndex >= 0 then
         FileNameEdt.text := FrmDictSelect.LBFiles.items[FrmDictSelect.LBFiles.ItemIndex]
end;

procedure TMVQueryDlg.CancelBtnClick(Sender: TObject);
Var
  CloseAction: TCloseAction;
begin
  CloseAction := caHide;
  FormClose(self,CloseAction);
end;

procedure TMVQueryDlg.btnAddFieldClick(Sender: TObject);
var
  Item : TListItem;

begin
   if FrmDictSelect.Listview1.SelCount > 0 then
   begin
     Item := FrmDictSelect.Listview1.selected;
//     FieldsMemo.lines.add(Item.Caption);
     FieldsMemo.SelLength:=0;
     FieldsMemo.SelText := Item.Caption + LineEnding

//    FieldsMemo.lines.Insert[FieldsMemo.ItemIndex, Item.Caption + LineEnding];
   end;
end;

initialization
  {$I umvquerydlg.lrs}

end.

