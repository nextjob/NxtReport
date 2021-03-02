unit utestreport;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, tisQMClient;

type

  { TFrmTestReport }

  TFrmTestReport = class(TForm)
    BtnDisconnect: TButton;
    BtnTestReport: TButton;
    EdtReportxml: TEdit;
    EdtPath: TEdit;
    EdtDefFile: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    ListBox1: TListBox;
    MemoStat: TMemo;
    procedure BtnDisconnectClick(Sender: TObject);
    procedure BtnTestReportClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

const
  //QM_Client Stuff
  QM_server = '127.0.0.1';
  QM_port = -1;

  //  qm field delimiters

  QM_AM = #254;  // attribute or field mark
  QM_VM = #253;  // value or item mark
  QM_SVM = #252;  // subvalue mark

var
  FrmTestReport: TFrmTestReport;

implementation

uses
  umainform, uqmsettings;

{ TFrmTestReport }

function MyQMGetArg(ArgNo: Integer): AnsiString;
// retrieves the value of argument variables to QMCall(), QMCallx()
// Rem use MyGetArg to free memory allocated by QMGetArg -- see QM Doc
var
  s: PAnsiChar;
begin
  s := QmGetArg(ArgNo);
  MyQMGetArg := PAnsiChar(s);
  QMFree(s);
end;

procedure TFrmTestReport.BtnDisconnectClick(Sender: TObject);
begin
  if QMConnected then
    QMDisconnect();
  MemoStat.Lines.Add('Disconnected');
end;

procedure TFrmTestReport.BtnTestReportClick(Sender: TObject);
var
  statcode: integer;
  err: string;
  Def_File_Path, Def_File_Name, Report_File_Path, Report_File_Name,
  Replace_Array, My_Options, My_Return_Status: ansistring;

begin
  if FileExists(EdtPath.text + PathDelim + EdtDefFile.Text) then
  begin
    if not QMconnected then
      FrmQmSettings.ShowModal;

    if QmConnected then
    begin

      MessageDlg('Generating Report Via Call to QM Subroutine NXTREPORT ',
        mtInformation, [mbOK], 0);
      Screen.Cursor := crHourglass;
      try

        Def_File_Path := EdtPath.Text;
        Def_File_Name := EdtDefFile.Text;
        Report_File_Path := EdtPath.Text;
        Report_File_Name := EdtReportxml.Text + '.xml';
        Replace_Array := '{' + umainform.TestData + '}';
        My_Options := ListBox1.Items[ListBox1.ItemIndex];
        My_Return_Status := '0';
        MemoStat.Lines.Add('Calling NXTREPORT ' + Def_File_Path + ', ' +  Def_File_Name + ', ' + Report_File_Path + ', ' + Report_File_Name);
        statcode := QMCallx('NXTREPORT', 7, PAnsiChar(Def_File_Path),PAnsiChar(Def_File_Name), PAnsiChar(Report_File_Path),PAnsiChar(Report_File_Name), PAnsiChar(Replace_Array),
        PAnsiChar(My_Options), PAnsiChar(My_Return_Status));
        err := MyQMGetArg(7);  // Rem use MyGetArg to free memory allocated by QMGetArg -- see QM Doc
        MemoStat.Lines.Add('Report call status QMCallx: ' + IntToStr(statcode) + ' NxtReport: ' + err);

      except
        on E: Exception do
        begin
          MessageDlg('Error in Execution:' + E.Message, mtInformation, [mbOK], 0);
        end;
      end;
      Screen.Cursor := crDefault;
    end
    else
      MessageDlg('QM Not Connected', mtInformation, [mbOK], 0);

  end
  else
    MessageDlg('Report Template File Not Found, Required for Report Testing',
        mtInformation, [mbOK], 0);
end;

procedure TFrmTestReport.FormActivate(Sender: TObject);
var
  NxtReportGen: string;
begin
  if QMConnected then
    MemoStat.Lines.Add('QMServer Currently Connected')
  else
    MemoStat.Lines.Add('QMServer Currently Disonnected');
  if length(MainForm.SaveDialog.FileName) = 0 then
    // user has not saved report def, assume opened file
  begin
    EdtPath.Text := ExtractFileDir(MainForm.OpenDialog.FileName);
    EdtDefFile.Text := ChangeFileExt(ExtractFileName(MainForm.OpenDialog.FileName), '.dat');
  end
  else
  begin
    EdtPath.Text := ExtractFileDir(MainForm.SaveDialog.FileName);
    EdtDefFile.Text := ChangeFileExt(ExtractFileName(MainForm.SaveDialog.FileName), '.dat');
  end;
  // Attempt to find NxtReportGen exe
  // should be in same directory as NxtDesinger
  NxtReportGen := ExtractFilePath(ParamStr(0)) + 'NxtReportGen.exe';
  if not (FileExists(NxtReportGen)) then
    MessageDlg('NxtReportGen Not Found in Assumed Location: ' +
      NxtReportGen + ' QM Subroutine NxtReport Will Error If Not Corrected!',
      mtInformation, [mbOK], 0);

end;

procedure TFrmTestReport.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin
  If QmConnected() then QMDisconnect();
end;

procedure TFrmTestReport.FormCreate(Sender: TObject);
begin
  //  MemoStat.Clear;
end;

initialization
  {$I utestreport.lrs}

end.
