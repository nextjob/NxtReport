program NxtDesigner;

{$mode objfpc}{$H+}

uses
  Interfaces, // this includes the LCL widgetset
  Forms, umainform, LResources, laz_fpspreadsheet_visual, utestreport,
  uprtsettings, umvquerydlg, uqmsettings, uDictSelect, umvreplacedlg,
  umvlookupdlg, LazSerialPort;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TFrmTestReport, FrmTestReport);
  Application.CreateForm(TFrmPrtSettings, FrmPrtSettings);
  Application.CreateForm(TMVQueryDlg, MVQueryDlg);
  Application.CreateForm(TFrmQmSettings, FrmQmSettings);
  Application.CreateForm(TFrmDictSelect, FrmDictSelect);
  Application.CreateForm(Tmvreplacedlg, mvreplacedlg);
  Application.CreateForm(Tmvlookupdlg, mvlookupdlg);
  Application.Run;
end.

