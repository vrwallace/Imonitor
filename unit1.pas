unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, ComCtrls, JwaWinSvc, JwaWinType, IniFiles, strutils, synaser, comobj, process;

type
  // Define a pointer to an array of ENUM_SERVICE_STATUS
  PEnumServiceStatusArray = ^TEnumServiceStatusArray;
  TEnumServiceStatusArray = array[0..0] of ENUM_SERVICE_STATUS;


type
  TFileTailThread = class(TThread)
  private
    FFileName: string;
    FLastSize: int64;
    FNewData: string;
    FErrorMessage: string;
    FTerminateRequested: boolean; // Flag to request termination
    procedure TailFile;
    procedure UpdateMemo;
    procedure UpdateGUIWithError;


  protected
    procedure Execute; override;
  public
    procedure RequestTerminate; // Method to set the termination flag
    constructor Create(CreateSuspended: boolean; const FileName: string);
  end;


  { TForm1 }

  TForm1 = class(TForm)
    Buttonrs1: TButton;
    Button2: TButton;
    Buttonrs2: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    ButtonSearchForward: TButton;
    CheckBoxspeechenabled: TCheckBox;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    comComboBox1: TComboBox;
    Edit1: TEdit;
    Edit2: TEdit;
    rs1s: TEdit;
    rs2s: TEdit;
    editsearch: TEdit;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Memo1: TMemo;
    Timer1: TTimer;
    procedure Buttonrs1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Buttonrs2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure ButtonSearchForwardClick(Sender: TObject);
    procedure CheckBoxspeechenabledChange(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure comComboBox1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    function GetIniFilePath: string;
    procedure ThreadTerminationHandler(Sender: TObject);
    procedure UpdateEditBoxWithError(const Msg: string);
    function GetServiceStatusAsString(ServiceName: string): string;
    function RebootComputer: boolean;
    function SpeakText(const TextToSpeak: WideString; Rate: integer): boolean;
    procedure PopulateServiceList(ComboBox: TComboBox);


    function RestartService(ServiceName: string): boolean;
    procedure Timer1Timer(Sender: TObject);
  private
    TailThread: TFileTailThread; // Reference to the tailing thread
  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

constructor TFileTailThread.Create(CreateSuspended: boolean; const FileName: string);
begin
  inherited Create(CreateSuspended);
  OnTerminate := @Form1.ThreadTerminationHandler;
  FFileName := FileName;
  FreeOnTerminate := True;
end;

function Tform1.GetServiceStatusAsString(ServiceName: string): string;
var
  hSCM, hService: SC_HANDLE;
  ServiceStatus: TServiceStatus;
begin
  Result := 'Unknown'; // Default result

  // Open the Service Control Manager
  hSCM := OpenSCManager(nil, nil, SC_MANAGER_CONNECT);
  if hSCM = 0 then Exit;

  try
    // Open the service
    hService := OpenService(hSCM, PChar(ServiceName), SERVICE_QUERY_STATUS);
    if hService = 0 then Exit;

    try
      // Query the service status
      if QueryServiceStatus(hService, ServiceStatus) then
      begin
        // Determine the status and set the result
        case ServiceStatus.dwCurrentState of
          SERVICE_RUNNING: Result := 'Running';
          SERVICE_STOPPED: Result := 'Stopped';
          SERVICE_PAUSED: Result := 'Paused';
          SERVICE_START_PENDING: Result := 'Starting';
          SERVICE_STOP_PENDING: Result := 'Stopping';
          // Add more cases if needed
        end;
      end;
    finally
      CloseServiceHandle(hService);
    end;
  finally
    CloseServiceHandle(hSCM);
  end;
end;


procedure Tform1.ThreadTerminationHandler(Sender: TObject);
begin
  // Check if the sender is indeed a TFileTailThread and not nil
  if (Sender is TFileTailThread) and (Sender <> nil) then
  begin
    // Update the user interface or perform other cleanup operations
    // For example, updating Edit1 to indicate that the thread has stopped
    Edit1.Text := 'Tail Stopped';
    speaktext(edit1.Text + ' on ' + ExtractFileName(edit2.Text), 0);
    Edit1.Color := clred; // Reset the color or set to a specific one

    // If you have any other operations to perform on thread termination, add them here
  end;
end;

function tform1.GetIniFilePath: string;
begin
  Result := IncludeTrailingPathDelimiter(GetUserDir) + 'imonitor\settings.ini';

end;

procedure TFileTailThread.RequestTerminate;
begin
  FTerminateRequested := True;
end;

procedure TFileTailThread.TailFile;
var
  FileStream: TFileStream;
  NewSize, BytesToRead: int64;
  Buffer: array of byte;
begin
  FileStream := TFileStream.Create(FFileName, fmOpenRead or fmShareDenyNone);
  try
    NewSize := FileStream.Size;
    if NewSize > FLastSize then
    begin
      BytesToRead := NewSize - FLastSize;
      SetLength(Buffer, BytesToRead);
      FileStream.Position := FLastSize;
      FileStream.Read(Buffer[0], BytesToRead);
      FLastSize := NewSize;

      // Convert the buffer to a string and store in FNewData
      SetString(FNewData, pansichar(@Buffer[0]), Length(Buffer));

      // Synchronize the call to UpdateMemo
      Synchronize(@UpdateMemo);
    end;
  finally
    FileStream.Free;
  end;
end;

procedure TFileTailThread.UpdateMemo;
begin

  while Form1.Memo1.Lines.Count > 10000 do
    Form1.Memo1.Lines.Delete(0);
  // Update the TMemo on Form1
  Form1.Memo1.Lines.Add(FNewData);
end;



procedure TFileTailThread.Execute;
var
  NotifyHandle: THandle;
  WaitStatus: DWORD;
  DirName: string;
begin
  try
    DirName := ExtractFilePath(FFileName);
    NotifyHandle := FindFirstChangeNotification(PChar(DirName), False,
      FILE_NOTIFY_CHANGE_SIZE);
    if NotifyHandle = INVALID_HANDLE_VALUE then
      RaiseLastOSError;
    while not FTerminateRequested do
    begin

      WaitStatus := WaitForSingleObject(NotifyHandle, 1000); // Timeout set to 1000 ms
      case WaitStatus of
        WAIT_OBJECT_0:
        begin
          Synchronize(@TailFile); // Use Synchronize to interact with the main thread
          if not FindNextChangeNotification(NotifyHandle) then
            RaiseLastOSError;
        end;
        WAIT_TIMEOUT:
          Continue; // Timeout occurred, loop again
        WAIT_FAILED:
          RaiseLastOSError;
      end;

    end;
    FindCloseChangeNotification(NotifyHandle);

  except
    on E: Exception do
    begin
      FErrorMessage := E.Message;
      Synchronize(@UpdateGUIWithError);
    end;
  end;
end;

procedure TFileTailThread.UpdateGUIWithError;
begin
  Form1.UpdateEditBoxWithError(FErrorMessage);
end;

procedure TForm1.Buttonrs1Click(Sender: TObject);
begin
  if (trim(combobox1.Text) <> 'N/A') then
  begin
    if not restartservice(combobox1.Text) then
    begin
      speaktext('Failed to restart the service', 0);
      ShowMessage('Failed to restart the service');
    end

    else
    begin
      speaktext('Service restarted successfully', 0);
      ShowMessage('Service restarted successfully');
    end;
    rs1s.Text := GetServiceStatusAsString(combobox1.Text);

  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  FileStream: TFileStream;
  filename: string;
begin
  if Assigned(TailThread) and not TailThread.Finished then
  begin
    speaktext('The file tailing is already in progress.', 0);
    ShowMessage('The file tailing is already in progress.');
    Exit;
  end;
  filename := edit2.Text;
  if fileexists(filename) then
  begin
    memo1.Clear;

    // Open the file to get its current size
    FileStream := TFileStream.Create(filename, fmOpenRead or fmShareDenyNone);
    try
      // Create and start the file tailing thread with the current file size
      TailThread := TFileTailThread.Create(True, filename);
      TailThread.FLastSize := FileStream.Size; // Set FLastSize to current file size
      edit1.Text := 'Tail Started';
      speaktext(edit1.Text + ' on ' + ExtractFileName(edit2.Text), 0);
      edit1.color := clgreen;
    finally
      FileStream.Free;
    end;

    TailThread.Start;

  end
  else
  begin
    speaktext(ExtractFileName(edit2.Text) + ' Not found!', 0);
    ShowMessage(filename + ' Not found!');
  end;
end;

procedure TForm1.Buttonrs2Click(Sender: TObject);
begin
  if (trim(combobox2.Text) <> 'N/A') then
  begin
    if not restartservice(combobox2.Text) then
    begin
      speaktext('Failed to restart the service', 0);
      ShowMessage('Failed to restart the service');
    end

    else
    begin
      speaktext('Service restarted successfully', 0);
      ShowMessage('Service restarted successfully');
    end;

    rs2s.Text := GetServiceStatusAsString(combobox2.Text);
  end;
end;


procedure TForm1.Button4Click(Sender: TObject);
begin
  if Assigned(TailThread) then
  begin
    TailThread.RequestTerminate; // Request the thread to terminate
    TailThread.WaitFor;          // Wait for the thread to finish execution
    // Do not explicitly free the thread here as FreeOnTerminate is set to True
    TailThread := nil;
    // edit1.Text := 'Tail Stopped';
    // speaktext(edit1.text+' on '+ExtractFileName(edit2.text),0);
    // edit1.color := clred;

  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var
  OpenDialog: TOpenDialog;
  IniFile: TIniFile;
begin
  OpenDialog := TOpenDialog.Create(nil);
  try
    OpenDialog.Filter := 'Log files (*.log)|*.log|All files (*.*)|*.*';
    OpenDialog.InitialDir := ExtractFileDir(Application.ExeName);
    if OpenDialog.Execute then
    begin
      Edit2.Text := OpenDialog.FileName;

      // Save to INI file in user's documents folder
      IniFile := TIniFile.Create(GetIniFilePath);
      try
        IniFile.WriteString('Settings', 'LastLogFile', OpenDialog.FileName);
      finally
        IniFile.Free;
      end;
    end;
  finally
    OpenDialog.Free;
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  memo1.Clear;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  ComComboBox1.Items.CommaText := GetSerialPortNames();
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  try
    if MessageDlg('Confirm', 'Are you sure you want to reboot the system?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if RebootComputer then
        speaktext('System reboot initiated successfully.', 0)
      // showmessage('System reboot initiated successfully.')
      else
      begin
        ShowMessage('Failed to initiate system reboot.');
        speaktext('Failed to initiate system reboot.', 0);
      end;

    end;
    //else
    //showmessage('Reboot cancelled by user.');

    speaktext('Reboot cancelled by user.', 0);
  except
    on E: Exception do
      WriteLn(E.ClassName, ': ', E.Message);
  end;
end;

function tform1.SpeakText(const TextToSpeak: WideString; Rate: integer): boolean;
var
  SpVoice: olevariant;
  SavedCW: word;
begin
  Result := False;

  if CheckBoxSpeechenabled.Checked then
  begin
    try
      SpVoice := CreateOleObject('SAPI.SpVoice');
      // Change FPU interrupt mask to avoid SIGFPE exceptions
      SavedCW := Get8087CW;

      try
        Set8087CW(SavedCW or $1 or $2 or $4 or $8 or $16 or $32);
        SpVoice.Rate := Rate;
        SpVoice.Speak(TextToSpeak, 1);

        repeat
          Application.ProcessMessages;
        until SpVoice.WaitUntilDone(10);

        Result := True;
      finally
        // Restore FPU mask
        Set8087CW(SavedCW);
        // VarClear(SpVoice); // Release the COM object
      end;
    except
      // Handle any exceptions here, if needed
    end;
  end;
end;



function tform1.RebootComputer: boolean;
var
  Token: THandle;
  TokenPrivileges: TTokenPrivileges;
  PrevTokenPrivileges: TTokenPrivileges;
  ReturnLength: DWord;
begin
  Result := False;

  // First, obtain the necessary security privileges to shut down the system
  if OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES or
    TOKEN_QUERY, Token) then
  begin
    // Get the LUID for shutdown privilege
    LookupPrivilegeValue(nil, 'SeShutdownPrivilege', TokenPrivileges.Privileges[0].Luid);

    TokenPrivileges.PrivilegeCount := 1;
    TokenPrivileges.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED;

    // Adjust token privileges to include shutdown privilege
    if AdjustTokenPrivileges(Token, False, TokenPrivileges,
      SizeOf(PrevTokenPrivileges), PrevTokenPrivileges, ReturnLength) then
    begin
      // Attempt to reboot the system
      Result := ExitWindowsEx(EWX_REBOOT or EWX_FORCE, 0);
    end;

    CloseHandle(Token);
  end;
end;




procedure TForm1.ButtonSearchForwardClick(Sender: TObject);
var
  SearchText: string;
  LowerMemoText: string;
  SearchResultPos, SearchStartPos: integer;
begin
  SearchText := LowerCase(EditSearch.Text);
  if SearchText = '' then Exit;

  // Convert Memo text to lowercase for case-insensitive search
  LowerMemoText := LowerCase(Memo1.Text);

  // Start search from the end of the last found text
  SearchStartPos := Memo1.SelStart + Memo1.SelLength;
  if SearchStartPos >= Length(LowerMemoText) then SearchStartPos := 0; // Wrap around

  SearchResultPos := PosEx(SearchText, LowerMemoText, SearchStartPos + 1);

  if SearchResultPos > 0 then
  begin
    Memo1.SetFocus; // Set focus to Memo1
    Memo1.SelStart := SearchResultPos - 1;
    Memo1.SelLength := Length(EditSearch.Text);
    // Use original length for correct selection
  end
  else if SearchStartPos > 0 then
  begin
    // If not found and search was not started from beginning, start from beginning
    SearchResultPos := Pos(SearchText, LowerMemoText);
    if SearchResultPos > 0 then
    begin
      Memo1.SetFocus; // Set focus to Memo1
      Memo1.SelStart := SearchResultPos - 1;
      Memo1.SelLength := Length(EditSearch.Text);
      // Use original length for correct selection
    end
    else
    begin
      speaktext('Text not found.', 0);
      ShowMessage('Text not found.');
    end;
  end
  else
  begin
    speaktext('Text not found.', 0);
    ShowMessage('Text not found.');
  end;
end;

procedure TForm1.CheckBoxspeechenabledChange(Sender: TObject);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(GetIniFilePath);
  try
    IniFile.WriteBool('Settings', 'SpeechEnabled', CheckBoxSpeechenabled.Checked);
  finally
    IniFile.Free;
  end;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(GetIniFilePath);
  try
    IniFile.Writestring('Settings', 'Service1', combobox1.Text);
  finally
    IniFile.Free;
  end;


  Buttonrs1.Caption := 'Restart ' + combobox1.Text;
  if combobox1.Text = 'N/A' then
  begin
    rs1s.Text := 'N/A';
    rs1s.color := clblue;
  end
  else rs1s.text:='';
end;

procedure TForm1.ComboBox2Change(Sender: TObject);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(GetIniFilePath);
  try
    IniFile.Writestring('Settings', 'Service2', combobox2.Text);
  finally
    IniFile.Free;
  end;

  Buttonrs2.Caption := 'Restart ' + combobox2.Text;
  if combobox2.Text = 'N/A' then
  begin
    rs2s.Text := 'N/A';
    rs2s.color := clblue;
  end
  else rs2s.text:='';

end;

procedure TForm1.comComboBox1Change(Sender: TObject);
begin

end;




procedure TForm1.FormCreate(Sender: TObject);
var
  IniFile: TIniFile;
  AProcess: TProcess;
  AStringList: TStringList;
  I: integer;
begin

  // Load from INI file in user's documents folder
  IniFile := TIniFile.Create(GetIniFilePath);
  try
    Edit2.Text := IniFile.ReadString('Settings', 'LastLogFile',
      'c:\logs\somefilename.log');
    CheckBoxSpeechenabled.Checked :=
      IniFile.ReadBool('Settings', 'SpeechEnabled', False);

    combobox2.Text := IniFile.ReadString('Settings', 'Service2', 'N/A');
    combobox1.Text := IniFile.ReadString('Settings', 'Service1', 'N/A');

  finally
    IniFile.Free;
  end;

  ButtonSearchforward.Default := True;

  if combobox1.Text <> 'N/A' then  Buttonrs1.Caption := 'Restart ' + combobox1.Text;
  populateservicelist(combobox1);
  if combobox2.Text <> 'N/A' then  Buttonrs2.Caption := 'Restart ' + combobox2.Text;
  populateservicelist(combobox2);
  if combobox1.Text = 'N/A' then
  begin
    rs1s.Text := 'N/A';
    rs1s.color := clblue;
  end;
  if combobox2.Text = 'N/A' then
  begin
    rs2s.Text := 'N/A';
    rs2s.color := clblue;
  end;


  form1.Refresh;

end;



procedure tform1.PopulateServiceList(ComboBox: TComboBox);
var
  SCManager: SC_HANDLE;
  ServiceStatusArray: PEnumServiceStatusArray;
  BytesNeeded, ServicesReturned, ResumeHandle: DWORD;
  I: integer;
begin
  SCManager := OpenSCManager(nil, nil, SC_MANAGER_ENUMERATE_SERVICE);
  if SCManager = 0 then
  begin
    ShowMessage('Failed to open Service Control Manager');
    Exit;
  end;

  try
    EnumServicesStatus(SCManager, SERVICE_WIN32, SERVICE_STATE_ALL, nil,
      0, BytesNeeded, ServicesReturned, ResumeHandle);

    if BytesNeeded = 0 then Exit; // No services or error

    GetMem(ServiceStatusArray, BytesNeeded);
    try
      ResumeHandle := 0;
      if EnumServicesStatus(SCManager, SERVICE_WIN32, SERVICE_STATE_ALL,
        @ServiceStatusArray^[0], BytesNeeded, BytesNeeded, ServicesReturned,
        ResumeHandle) then
      begin
        ComboBox.Items.BeginUpdate;
        try
          ComboBox.Items.Clear;
          ComboBox.Items.Add('N/A');
          for I := 0 to ServicesReturned - 1 do
          begin
            ComboBox.Items.Add(ServiceStatusArray^[I].lpServiceName);
          end;
        finally
          ComboBox.Items.EndUpdate;
        end;
      end
      else
      begin
        ShowMessage('Failed to enumerate services');
      end;
    finally
      FreeMem(ServiceStatusArray);
    end;
  finally
    CloseServiceHandle(SCManager);
  end;
end;



procedure TForm1.FormDestroy(Sender: TObject);
begin
  if Assigned(TailThread) then
  begin
    TailThread.RequestTerminate; // Request the thread to terminate
    TailThread.WaitFor;          // Wait for the thread to finish execution
    // Note: Since FreeOnTerminate is set to True, the thread will free itself
  end;
end;




function tform1.RestartService(ServiceName: string): boolean;
var
  hSCM, hService: SC_HANDLE;
  ServiceStatus: TServiceStatus;
begin
  Result := False;
  hSCM := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
  if hSCM = 0 then
    Exit;

  try
    hService := OpenService(hSCM, PChar(ServiceName), SERVICE_STOP or
      SERVICE_START or SERVICE_QUERY_STATUS);
    if hService = 0 then
      Exit;

    try
      // Check if the service is running and then stop it
      if QueryServiceStatus(hService, ServiceStatus) and
        (ServiceStatus.dwCurrentState = SERVICE_RUNNING) then
      begin
        if ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus) then
        begin
          // Wait for the service to stop
          while QueryServiceStatus(hService, ServiceStatus) and
            (ServiceStatus.dwCurrentState <> SERVICE_STOPPED) do
            Sleep(1000);
        end;
      end;

      // Start the service
      if not StartService(hService, 0, nil) then
        Exit;

      Result := True;
    finally
      CloseServiceHandle(hService);
    end;
  finally
    CloseServiceHandle(hSCM);
  end;
end;

procedure TForm1.UpdateEditBoxWithError(const Msg: string);
begin
  if Assigned(Edit1) then
  begin
    Edit1.Text := Msg;
    Edit1.Color := clRed;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin

  ComComboBox1.Items.CommaText := GetSerialPortNames();
  if (trim(combobox1.Text) <> 'N/A') then
  begin
    if GetServiceStatusAsString(combobox1.Text) <> rs1s.Text then
      speaktext(combobox1.Text + ' ' + GetServiceStatusAsString(combobox1.Text) +
        ' on ' + ExtractFileName(edit2.Text), 0);
    rs1s.Text := GetServiceStatusAsString(combobox1.Text);
    if rs1s.Text = 'Running' then rs1s.color := clgreen
    else
      rs1s.color := clred;

  end;
  if (trim(combobox2.Text) <> 'N/A') then
  begin
    if GetServiceStatusAsString(combobox2.Text) <> rs2s.Text then
      speaktext(combobox2.Text + ' ' + GetServiceStatusAsString(combobox2.Text) +
        ' on ' + ExtractFileName(edit2.Text), 0);
    rs2s.Text := GetServiceStatusAsString(combobox2.Text);
    if rs2s.Text = 'Running' then rs2s.color := clgreen
    else
      rs2s.color := clred;

  end;

end;

end.
