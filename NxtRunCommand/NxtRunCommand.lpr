program NxtRunCommand;
//
// Program used to execute another program.
// Created to execute program without waiting for completion.
// via QMBasic OS.EXECUTE
// usage NxtRunCommand fullyQuilifiedPathofExe param(s) to pass
// ie)
// NxtRunCommand C:\Laz_projects\NxtReport\NxtDesigner\NxtReportGen.exe -f C:\QMACCOUNTS\nxtReportDefs\MyTestReport.xml -s XLSX -d
// Notes
// To prevent "cmd" terminal from display:
// You can compile your app as {$apptype GUI} (or similar), or use -WG  (project options - custom options)
// Then windows thinks it is a gui app, but since it does not open a form, nothing will appear.
// However, in this case, if you want to use writeln to stdout or read stdin, you must first check they are open.
//
{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Classes,
  Process,
  SysUtils { you can add units after this };

var
  i: integer;
  executable: string;
  Result: boolean;
  MyProcess: TProcess;
begin
  if FileExists(ParamStr(1)) then
  begin

  executable := ParamStr(1);
//  writeln('execute: ' + executable);
  MyProcess := TProcess.Create(nil);
  try
    MyProcess.Executable := executable;
    for i := 2 to paramCount() do
       MyProcess.Parameters.Add(ParamStr(i));
// Note to hide the cmd window of this process -- see use of -wg compile option above
//    MyProcess.Options := MyProcess.Options + [poNoConsole];
//    MyProcess.StartupOptions := MyProcess.StartupOptions + [suoUseShowWindow];
//    MyProcess.ShowWindow := swoHIDE;
    MyProcess.Execute;
  finally
    MyProcess.Free;
  end;

  end;
end.
