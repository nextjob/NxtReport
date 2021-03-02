unit umvreplacedlg;
{
  Insert user selected MVREPLACE function and tag information into active cell
}
{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs, StdCtrls, fpstypes;

type

  { Tmvreplacedlg }

  Tmvreplacedlg = class(TForm)
    CancelBtn: TButton;
    CBTag: TComboBox;
    DescEdt: TEdit;
    Label2: TLabel;
    OkBtn: TButton;
    Label1: TLabel;
    procedure CancelBtnClick(Sender: TObject);
    procedure OkBtnClick(Sender: TObject);
  private

  public

  end;

var
  mvreplacedlg: Tmvreplacedlg;

implementation

uses
    umainform;

{ Tmvreplacedlg }

procedure Tmvreplacedlg.OkBtnClick(Sender: TObject);
  var
    i: integer;
    errmsg: string;
    celltext: string;
    displayflds: string;

  begin
    if length(cbTag.Text) = 0 then
      errmsg := 'Tag Name is Missing';

    cbTag.Text := UpperCase(trim(cbTag.Text));

    if length(DescEdt.Text) = 0 then
      if cbTag.Text = 'RUNTIME' then
        DescEdt.Text := 'Report Generation Run Time'
      else if cbTag.Text = 'RUNDATE' then
        DescEdt.Text := 'Report Generation Run Date'
      else
        errmsg := 'Tag Desc is Missing';

    if length(errmsg) > 0 then
    begin
      errmsg := errmsg + ', Please Correct';
      MessageDlg(errmsg, mtError, [mbOK], 0);
      ModalResult := mrNone;
    end
    else
    begin
       // parse the form controls to get text to enter into the active cell
      celltext := '"%%' + trim(cbTag.Text) + '%%","' + DescEdt.Text + '"';
      //  enter the mvreplace text into the currently active cell
      celltext := '=MVREPLACE(' + celltext + ')';
      MainForm.WorksheetGrid.Worksheet.WriteFormula(
        MainForm.WorksheetGrid.Worksheet.ActiveCellRow,
        MainForm.WorksheetGrid.Worksheet.ActiveCellCol, celltext);
      // for known replace tags set default format
      If  cbTag.Text = 'RUNTIME' then
        MainForm.WorksheetGrid.Worksheet.WriteDateTimeFormat( MainForm.WorksheetGrid.Worksheet.ActiveCellRow,
        MainForm.WorksheetGrid.Worksheet.ActiveCellCol,nfShortTimeAM)
      Else If cbTag.TExt = 'RUNDATE' then
        MainForm.WorksheetGrid.Worksheet.WriteDateTimeFormat( MainForm.WorksheetGrid.Worksheet.ActiveCellRow,
        MainForm.WorksheetGrid.Worksheet.ActiveCellCol,nfShortDate);

    end;

  end;

procedure Tmvreplacedlg.CancelBtnClick(Sender: TObject);

begin
  Close;
end;

initialization
  {$I umvreplacedlg.lrs}

end.

