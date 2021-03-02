unit uDictSelect;

{
Routine reads QMFILES.JSON file and populates LBFiles- list box.
User selects qm file and the list view ListView1 is populated with directory items
Note Routine ASSUMES JSON file QMFILES.JSON is found in the same directory as executable
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Grids,
  ComCtrls, fpjson;

type

  { TFrmDictSelect }

  TFrmDictSelect = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    LBFiles: TListBox;
    ListView1: TListView;
    procedure FormShow(Sender: TObject);
    procedure LBFilesClick(Sender: TObject);
  private
    Myjson: TJSONObject;
    procedure ReadJSONFromFile(FFileName: TFileName);
    procedure ReadJSONFromStream(AStream: TStream);
  public

  end;

var
  FrmDictSelect: TFrmDictSelect;

implementation

uses
  LazFileUtils,
  jsonscanner, jsonparser;

{$R *.lfm}

{ TFrmDictSelect }

procedure TFrmDictSelect.FormShow(Sender: TObject);
var
  i, remainingLVSize, ColSiz, NbrOfCols: integer;
begin
  // size the listview columns (this is a pia that it needs to be done this way!)
  remainingLVSize := ListView1.Width;
  NbrOfCols := ListView1.Columns.Count;
  ColSiz := remainingLVSize div NbrOfCols;
  for i := 0 to NbrOfCols - 1 do
    //    ListView1.Columns[i].Width :=  Length(ListView1.Column[i].Caption);
    if i = (NbrOfCols - 1) then
      ListView1.Columns[i].Width := remainingLVSize
    else
    begin
      ListView1.Columns[i].Width := ColSiz;
      remaininglvSize := remaininglvSize - ColSiz;
    end;

  ListView1.Items.Clear;
  // have we already loaded the JSON data?
  if LBFiles.Count <= 0 then
    ReadJSONFromFile('QMFILES.JSON');   // nope, go get it
end;


procedure TFrmDictSelect.ReadJSONFromFile(FFileName: TFileName);
var
  stream: TFileStream;
  msgText: string;
begin
  if (FFileName = '') then
    MessageDlg(' No JSON File Name Specified', mtInformation, [mbOK], 0, mbOK)
  else if not FileExists(FFileName) then
  begin
    msgText := FFileName + ' Was Not Found in NxtDesigner Program Directory.';
    msgText := msgText + LineEnding +
      ' Use the QM Program CREATE_QMFILES_JSON to Generate This File,';
    msgText := msgText + LineEnding + ' And Copy into NxtDesigners Program Directory.';
    MessageDlg(msgText, mtInformation, [mbOK], 0, mbOK);

  end
  else
  begin
    stream := TFileStream.Create(FFilename, fmOpenRead + fmShareDenyWrite);
    try
      ReadJSONFromStream(stream);
    finally
      stream.Free;
    end;
  end;
end;

procedure TFrmDictSelect.ReadJSONFromStream(AStream: TStream);
var
  p: TJSONParser;

  FileObj: TJSONObject;
  FilesObjArray: TJSONArray;
  i: integer;
  QmFileName: string;
  options: TJSONOptions;

begin

  options := [];
  {
  if FUseUTF8 then options := options + [joUTF8];
  if FStrict then options := options + [joStrict];
  if FComments then options := options + [joComments];
  if FIgnoreTrailingComma then options := options + [joIgnoreTrailingComma];
  }
  p := TJSONParser.Create(AStream, options);

  try
    Myjson := p.Parse as TJSONObject;
    if Myjson = nil then
      exit;
    FilesObjArray := Myjson.Find('Files', jtArray) as TJSONArray;
    if Assigned(FilesObjArray) then
      for i := 0 to FilesObjArray.Count - 1 do
      begin
        FileObj := FilesObjArray.Objects[i];
        QmFileName := FileObj.Find('Filename').AsString;
        LBFiles.Items.Add(QmFileName);
      end;
  finally
    p.Free;
  end;
end;

procedure TFrmDictSelect.LBFilesClick(Sender: TObject);
var
  FileObj, DictObj: TJSONObject;
  FilesObjArray, DictObjArray: TJSONArray;
  i, j: integer;
  QmFileName, FileToFind: string;
  DictName, DictDesc: string;
  FileFound: boolean;
  LVItem: TListItem;

begin

  // get the selected file from the listbox

  if LBFiles.ItemIndex >= 0 then
    FileToFind := LBFiles.items[LBFiles.ItemIndex]
  else
    FileToFind := '';

  ListView1.Items.Clear;
  FileFound := False;

  if Myjson = nil then
    exit;
  if length(FileToFind) > 0 then
  begin
    try
      FilesObjArray := Myjson.Find('Files', jtArray) as TJSONArray;
      if Assigned(FilesObjArray) then
        for i := 0 to FilesObjArray.Count - 1 do
        begin
          FileObj := FilesObjArray.Objects[i];
          QmFileName := FileObj.Find('Filename').AsString;
          if QmFileName = FileToFind then  // found it
          begin
            FileFound := True;
            DictObjArray := FileObj.Find('Dictionary', jtArray) as TJSONArray;
            //             memo2.Lines.Add(DictObjArray.AsJSON);
            if Assigned(DictObjArray) then
              for j := 0 to DictObjArray.Count - 1 do
              begin
                DictObj := DictObjArray.Objects[j];
                DictName := DictObj.Strings['Dictname'];
                DictDesc := DictObj.Strings['4'];
                //                add dictionary info to listview
                LVItem := ListView1.Items.Add;
                LVItem.Caption := DictName;
                LVItem.Subitems.Add(DictDesc);

              end;
          end;

        end;
      if not FileFound then
        MessageDlg(FileToFind + ' Not Found in Supplied JSON Data', mtInformation,
          [mbOK], 0, mbOK);

    except
      on E: Exception do
        ShowMessage('Error: ' + E.ClassName + #13#10 + E.Message);
    end;
  end;
end;

end.
