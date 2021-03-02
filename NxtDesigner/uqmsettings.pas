unit uqmsettings;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, tisQMClient;

type

  { TFrmQmSettings }

  TFrmQmSettings = class(TForm)
    BtnConnect: TButton;
    BtnDisconnect: TButton;
    EdtAccount: TEdit;
    EdtPasswrd: TEdit;
    EdtUserName: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    MemoStat: TMemo;
    procedure BtnConnectClick(Sender: TObject);
    procedure BtnDisconnectClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;
const
  //QM_Client Stuff
  QM_server = '127.0.0.1';
  QM_port   = -1;

 //  qm field delimiters

  QM_AM        =   #254;  // attribute or field mark
  QM_VM        =   #253;  // value or item mark
  QM_SVM       =   #252;  // subvalue mark

var
  FrmQmSettings: TFrmQmSettings;

implementation

{ TFrmQmSettings }

procedure TFrmQmSettings.BtnConnectClick(Sender: TObject);
// Rem: If Security is enabled on QM,
//  and user not defined in Secuirty subsytem
//  connect will fail!
//  See QM ref manual
begin
  if QMconnected then
  begin
   MemoStat.Lines.Add('QM is currently Connected')
  end else
    begin
      MemoStat.Lines.Add( 'Connecting to QM Server');
 //  attempt connection with data entered on LoginFrm
          if QMConnect(QM_server, QM_port, EdtUserName.Text, EdtPasswrd.Text, EdtAccount.Text) then

 // or optionally user the connect local version
 //           if QMConnectLocal(EdtAccount.Text) then
             MemoStat.Lines.Add('Connected')
            else
             begin
              MessageDlg('Connection failed: ' + QMError(), mtInformation, [mbOk], 0);
              MemoStat.Lines.Add('Connection failed: ' + QMError())
             end;
    end;
end;

procedure TFrmQmSettings.BtnDisconnectClick(Sender: TObject);
begin
  if QMConnected then
    QMDisconnect();
  MemoStat.Lines.Add('Disconnected');
end;

procedure TFrmQmSettings.FormActivate(Sender: TObject);
begin
  if QMConnected then
    MemoStat.Lines.Add('Currently Connected')
  else
    MemoStat.Lines.Add('Currently Disonnected');
end;

procedure TFrmQmSettings.FormCreate(Sender: TObject);
begin
//  MemoStat.Clear;
end;

initialization
  {$I uqmsettings.lrs}

end.

