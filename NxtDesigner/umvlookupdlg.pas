unit umvlookupdlg;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls;

type

  { Tmvlookupdlg }

  Tmvlookupdlg = class(TForm)
    CancelBtn: TButton;
    FileNameEdt: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    OkBtn: TButton;
    FieldDictEdt: TEdit;
    RecordIdEdt: TEdit;
    procedure CancelBtnClick(Sender: TObject);
    procedure FileNameEdtClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure OkBtnClick(Sender: TObject);
    procedure FieldDictEdtClick(Sender: TObject);
  private
    procedure FillEditText(Sender: TEdit);

  public

  end;

var
  mvlookupdlg: Tmvlookupdlg;

implementation

uses
    umainform, udictselect;

{ Tmvlookupdlg }
procedure TMVLookupDlg.FillEditText(Sender: TEdit);
var
  Item: TListItem;
begin
  if FrmDictSelect.Listview1.SelCount > 0 then
  begin
    Item := FrmDictSelect.Listview1.selected;
    Sender.Text := Item.Caption;
  end;
end;

procedure TMVLookupDlg.FormShow(Sender: TObject);
begin
  FrmDictSelect.Show;
end;

procedure Tmvlookupdlg.OkBtnClick(Sender: TObject);
var
  i: integer;
  errmsg: string;
  celltext: string;
  displayflds: string;
  CloseAction: TCloseAction;
begin
  if length(FileNameEdt.Text) = 0 then
    errmsg := 'File Name is Missing';

  if length(RecordIdEdt.Text) = 0 then
    errmsg := 'Record Key is Missing';

  if length(FieldDictEdt.Text) = 0 then
    errmsg := 'Dictionary Name or Field Number Missing';

  if length(errmsg) > 0 then
  begin
    errmsg := errmsg + ', Please Correct';
    MessageDlg(errmsg, mtError, [mbOK], 0);
    ModalResult := mrNone;
  end
  else
  begin
    // parse the form controls to get text to enter into the active cell
    celltext := '"' + FileNameEdt.Text + '","' + RecordIdEdt.Text + '","' + FieldDictEdt.text + '"';
    //  enter the mvquery text into the currently active cell
    celltext := '=MVLOOKUP(' + celltext + ')';
    MainForm.WorksheetGrid.Worksheet.WriteFormula(
      MainForm.WorksheetGrid.Worksheet.ActiveCellRow,
      MainForm.WorksheetGrid.Worksheet.ActiveCellCol, celltext);
    CloseAction := caHide;
    FormClose(Sender,CloseAction);
  end;

end;

procedure TMVLookupDlg.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  FrmDictSelect.Hide;
  mvLookupdlg.Hide;
end;

procedure TMVLookupDlg.FileNameEdtClick(Sender: TObject);
begin
  if FrmDictSelect.LBFiles.ItemIndex >= 0 then
    FileNameEdt.Text :=
      FrmDictSelect.LBFiles.items[FrmDictSelect.LBFiles.ItemIndex];
end;

procedure Tmvlookupdlg.CancelBtnClick(Sender: TObject);
  Var
  CloseAction: TCloseAction;
begin
  CloseAction := caHide;
  FormClose(Sender,CloseAction);
end;


procedure TMVLookupDlg.FieldDictEdtClick(Sender: TObject);
begin
  FillEditText(FieldDictEdt);
end;

initialization
  {$I umvlookupdlg.lrs}
end.
