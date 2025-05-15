unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, System.Zip,
  DragDrop, DropTarget, DragDropFile, System.IOUtils, System.Types;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    Edit2: TEdit;
    Label2: TLabel;
    Edit3: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    DropFileTarget1: TDropFileTarget;
    Button3: TButton;
    Label5: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DropFileTarget1Drop(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: LongInt);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    if FileExists(OpenDialog1.FileName) then
      Edit1.Text := (OpenDialog1.FileName)
    else
      raise Exception.Create('File does not exist.');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    if FileExists(SaveDialog1.FileName) then
      raise Exception.Create('File already exists. Cannot overwrite.')
    else
      Edit2.Text := (SaveDialog1.FileName);
end;

function ExecPNGQuant(sCommandLine, sWorkDir: string): Boolean;
var
  dwExitCode: DWORD;
  tpiProcess: TProcessInformation;
  tsiStartup: TStartupInfo;
begin
  Result := false;
  FillChar(tsiStartup, sizeof(TStartupInfo), 0);
  tsiStartup.cb := sizeof(TStartupInfo);
  tsiStartup.hStdError := 0;
  tsiStartup.wShowWindow := SW_HIDE;
  tsiStartup.dwFlags := StartF_UseStdHandles + STARTF_USESHOWWINDOW;
  if CreateProcess(nil, PChar(IncludeTrailingBackSlash(GetCurrentDir) +
    'pngquant.exe ' + sCommandLine), nil, nil, false, 0, nil, PChar(sWorkDir),
    tsiStartup, tpiProcess) then
  begin
    if WAIT_OBJECT_0 = WaitForSingleObject(tpiProcess.hProcess, INFINITE) then
    begin
      if GetExitCodeProcess(tpiProcess.hProcess, dwExitCode) then
      begin
        if dwExitCode = 0 then
          Result := True
        else
          SetLastError(dwExitCode + $2000);
      end;
    end;
    dwExitCode := GetLastError;
    CloseHandle(tpiProcess.hProcess);
    CloseHandle(tpiProcess.hThread);
    SetLastError(dwExitCode);
    Result := True;
  end;
end;

function GetTempFolder: string;
begin
  Randomize;
  Result := IncludeTrailingBackSlash(GetEnvironmentvariable('Temp')) +
    inttohex(Random(integer.MaxValue)) + '\';
end;
(*
  function ProcessFiles(SourceFile,DestinationFile:string): boolean;
  var
  dir: string;
  i: integer;
  ProcTask:TTask;
  begin
  ProcTask:=TTask.Create;
  Application.ProcessMessages;
  ProcTask.Perform(
  procedure()
  //  var
  //    x: integer;
  begin

  end);
  end; *)

procedure TForm1.Button3Click(Sender: TObject);
var
  dir, dir1: string;
  i: integer;
  PNGS: TStringDynArray;
begin
  if (FileExists(Edit1.Text)) and
    (ExtractFileExt(Edit1.Text) = ExtractFileExt(Edit2.Text)) then
  begin
    dir := GetTempFolder;
    if (ExtractFileExt(Edit1.Text) = '.docx') or
      (ExtractFileExt(Edit1.Text) = '.docm') then
      dir1 := dir + 'word\media\';
    if (ExtractFileExt(Edit1.Text) = '.xlsx') or
      (ExtractFileExt(Edit1.Text) = '.xlsm') then
      dir1 := dir + 'xl\media\';
    if (ExtractFileExt(Edit1.Text) = '.pptx') or
      (ExtractFileExt(Edit1.Text) = '.pptm') then
      dir1 := dir + 'ppt\media\';
    forcedirectories(dir);
    if TZipFile.IsValid(Edit1.Text) then
    begin
      TZipFile.ExtractZipFile(Edit1.Text, dir);
      PNGS := TDirectory.GetFiles(dir1, '*.png');
      for i := Low(PNGS) to High(PNGS) do
      begin
        ExecPNGQuant(Edit3.Text + ' --force "' + PNGS[i] + '" --output "' +
          PNGS[i] + '"', dir);
        // showmessage(Edit3.Text + ' --force "' + PNGS[i] + '" --output"' +PNGS[i] + '"');
      end;
      // ExecPNGQuant(Edit3.Text+ ' "'+Edit1.Text+'" --output"'+Edit2.Text+'"',dir);
      TZipFile.ZipDirectoryContents(Edit2.Text, dir);
      TDirectory.delete(dir, True);
      if TZipFile.IsValid(Edit2.Text) then
        ShowMessage('Optimization/Compression Ok')
      else
        ShowMessage('Failed!');
    end
    else
      ShowMessage('Not a valid file');
  end
  else
    raise Exception.Create('Please choose a valid file');
end;

procedure TForm1.DropFileTarget1Drop(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: LongInt);
begin

  if (ExtractFileExt(DropFileTarget1.Files[0]) = '.docx') or
    (ExtractFileExt(DropFileTarget1.Files[0]) = '.xlsx') or
    (ExtractFileExt(DropFileTarget1.Files[0]) = '.pptx') or
    (ExtractFileExt(DropFileTarget1.Files[0]) = '.docm') or
    (ExtractFileExt(DropFileTarget1.Files[0]) = '.xlsm') or
    (ExtractFileExt(DropFileTarget1.Files[0]) = '.pptm') then
    Edit1.Text := DropFileTarget1.Files[0];
  Edit2.Text := ChangeFileext(DropFileTarget1.Files[0], '') + '_optimized' +
    ExtractFileExt(DropFileTarget1.Files[0]);
  if (Effect = DROPEFFECT_MOVE) then
    Effect := DROPEFFECT_NONE;

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  if not FileExists('pngquant.exe') then
  begin
    raise Exception.Create('Cannot find pngquant');
    Halt(1);
  end;
  OpenDialog1.Filter := 'Word Documents (*.docm;*.docx)|*.docm;*.docx|' +
    'Excel Spreadsheets (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|' +
    'PowerPoint Presentations (*.pptm;*.pptx)|*.pptm;*.pptx|' +
    'All Office Files (*.docm;*.docx;*.xlsm;*.xlsx;*.pptm;*.pptx)|*.docm;*.docx;*.xlsm;*.xlsx;*.pptm;*.pptx|'
    + 'All Files (*.*)|*.*';
  SaveDialog1.Filter := 'Word Documents (*.docm;*.docx)|*.docm;*.docx|' +
    'Excel Spreadsheets (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|' +
    'PowerPoint Presentations (*.pptm;*.pptx)|*.pptm;*.pptx|' +
    'All Office Files (*.docm;*.docx;*.xlsm;*.xlsx;*.pptm;*.pptx)|*.docm;*.docx;*.xlsm;*.xlsx;*.pptm;*.pptx|'
    + 'All Files (*.*)|*.*';
  { --force --verbose --quality=45-85 }
end;

end.
