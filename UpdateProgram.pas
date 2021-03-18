unit UpdateProgram;

interface

uses
  Winapi.Windows, Winapi.Messages,

  System.SysUtils, System.Variants, System.Classes, System.UITypes,
  System.StrUtils,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ExtCtrls,

  JvAppInst,

  MSspecialFolders, SBPro, PJVersionInfo;



type
  TClosure = (tClosed, tNotFound, tNotClosed);

  TUpdaterInfo = record
    ProductName : string;
    InstallApp  : string;
    IconFile    : string;
    VersionInfo : string;

    IconFileOk   : Boolean;
    InstallAppOk : Boolean;
  end;

  Tfrm_MainUpdate = class(TForm)
    pnl_Main: TPanel;
    img_Logo: TImage;
    lbl_MainPanel: TLabel;
    lbl_UpdateMessage: TLabel;
    btn_Update: TBitBtn;
    stsbrpr_UpdateSpending: TStatusBarPro;
    btn_Exit: TSpeedButton;
    lbl_Versions: TLabel;
    jvpnstncs_1: TJvAppInstances;
    lbl_UpdateMessages: TLabel;
    mmo_UpdateInfo: TMemo;
    PJVersionInfo1: TPJVersionInfo;
    procedure FormCreate(Sender: TObject);
    procedure btn_UpdateClick(Sender: TObject);
    procedure btn_ExitClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SetStatusBar(aMsg: string; aColor: tColor);
  private
    { Private declarations }
    AppsToKill     : TStringList;
    UpdaterInfoPath: string;
    UpdaterInfo    : TUpdaterInfo;


    function GetAppVersionStr(aNumPlaces: Byte): string;
    procedure runUpdate;
  public
    { Public declarations }
  end;

var
  frm_MainUpdate: Tfrm_MainUpdate;

implementation
uses
  WinApi.ShellApi, Winapi.tlhelp32;
{$R *.dfm}

const
  cUpdaterCloseAppsFile = 'GEMTmpCloseApps.txt';
  cUpdaterInfo          = 'GEMTmpUpdateInfo.txt';
  cUpdaterDataPath       = '\SlickRockSoftwareDesign\SRSDUpdater';
//  cUpdaterCloseAppsFile  = '\CloseApps.txt';
//  cUpdaterInfo           = '\UpdateInfo.txt';


function GetContrastColor(ABGColor: TColor): TColor;
var
  ADouble: Double;
  R, G, B: Byte;
begin
  if ABGColor <= 0 then
  begin
    Result := clWhite;
    Exit; // *** EXIT RIGHT HERE ***
  end;

  if ABGColor = clWhite then
  begin
    Result := clBlack;
    Exit; // *** EXIT RIGHT HERE ***
  end;

  // Get RGB from Color
  R := GetRValue(ABGColor);
  G := GetGValue(ABGColor);
  B := GetBValue(ABGColor);

  // Counting the perceptive luminance - human eye favors green color...
  ADouble := 1 - ( 0.299 * R + 0.587 * G + 0.114 * B)/255;

  if (ADouble < 0.5) then
    Result := clBlack  // bright colors - black font
  else begin
    Result := clWhite;  // dark colors - white font
  end;
end;


function Killprocess(Name: String): TClosure;
// close process by exe name
var
  PEHandle,hproc: cardinal;
  PE: ProcessEntry32;
  found: boolean;
begin
  PEHandle := CreateTOOLHelp32Snapshot(TH32cs_Snapprocess,0);
  found := false;
  result :=  tNotFound;
  if PEHandle <> Invalid_Handle_Value then begin
    PE.dwSize := Sizeof(ProcessEntry32);
    Process32first(PEHandle,PE);
    repeat
      if Lowercase(PE.szExeFile) = Lowercase(Pchar(Name)) then begin
        try
          hproc := openprocess(Process_Terminate,false,pe.th32ProcessID);
          TerminateProcess(hproc,0);
          closehandle(hproc);
          Result := tClosed;
          found := true;
        except
          Result := tNotClosed;
        end;
      end
      else
        Result := tNotFound;
    until (Process32next(PEHandle,PE) = false) or Found;
  end;
  closehandle(PEHandle);
end;


procedure Tfrm_MainUpdate.SetStatusBar(aMsg: string; aColor: tColor);
begin
  stsbrpr_UpdateSpending.Panels[0].Color := aColor;
  stsbrpr_UpdateSpending.Panels[0].Font.color := GetContrastColor(aColor);
  stsbrpr_UpdateSpending.Panels[0].Text := aMsg;
end;


procedure Tfrm_MainUpdate.btn_ExitClick(Sender: TObject);
begin
  Close;
end;


procedure Tfrm_MainUpdate.btn_UpdateClick(Sender: TObject);
var
  vError: Boolean;
  i: Integer;
begin
  verror := False;
  for i := 0 to AppsToKill.Count -1 do begin
    case Killprocess(AppsToKill[i]) of
      tClosed:  begin
        SetStatusBar(' App Closed', clLime);
        mmo_UpdateInfo.Lines.Add(AppsToKill[i]+ ' App Closed');
      end;

      tNotFound: begin
        SetStatusBar(' App not Found To Close', clYellow);
        mmo_UpdateInfo.Lines.Add(AppsToKill[i]+ ' App not Found To Close');
      end;

      tNotClosed: begin
        SetStatusBar(' App NOT Closed', clRed);
        mmo_UpdateInfo.Lines.Add('===============================================');
        mmo_UpdateInfo.Lines.Add(AppsToKill[i]+ ' App NOT Closed.  Can not update' +
                                          'until the app is closed');
        mmo_UpdateInfo.Lines.Add('Close the app and try again by clicking the "UpDate" button');
        mmo_UpdateInfo.Lines.Add('===============================================');
        verror := true;
      end;
    end;
  end;

  if not vError then begin
    runUpdate;
  end
  else
    MessageDlg('Could not close all the apps!'+#13+#10+
               'Can NOT update program until all'+#13+#10+
               'apps are closed.', mtError, [mbOK], 0);
end;


procedure Tfrm_MainUpdate.FormCreate(Sender: TObject);
begin
  PJVersionInfo1.FileName := Application.ExeName;
  Caption := PJVersionInfo1.ProductVersionNumber + ' - SRSD Updater App: ' ;

  mmo_UpdateInfo.Clear;
  AppsToKill := TStringList.Create;

  UpdaterInfoPath := GetTempDirectory;
//  UpdaterInfoPath := getWinSpecialFolder(CSIDL_COMMON_APPDATA, false)+ cUpdaterDataPath;
  AppsToKill.LoadFromFile(UpdaterInfoPath + cUpdaterCloseAppsFile);

//  lbl_Versions.Caption := ParamStr(1) + '-->' + ParamStr(2);
//  Caption := 'Update Spending: ' + GetAppVersionStr(4);
//  SetProgramPaths;
end;


procedure Tfrm_MainUpdate.FormDestroy(Sender: TObject);
begin
  FreeAndNil(AppsToKill);
end;


procedure Tfrm_MainUpdate.FormShow(Sender: TObject);
var
  UdInfo    : TStringList;
  index     : Integer;
  rStr, lStr: string;
begin
  UdInfo := TStringList.Create;
  try
    UdInfo.LoadFromFile(UpdaterInfoPath + cUpdaterInfo);
    for index := 0 to UdInfo.Count -1 do begin
      lStr := AnsiLeftStr(UdInfo[index], Pos('|', UdInfo[index]) -1);
      rStr := AnsiRightStr(UdInfo[index], Length(UdInfo[index]) - Pos('|', UdInfo[index]));

//        Caption := 'SRSD Updater App - :'+ PJVersionInfo1.ProductVersionNumber;


      if lStr = 'ProductName' then begin
        UpdaterInfo.ProductName := rStr;
        Caption := PJVersionInfo1.ProductVersionNumber + '-SRSD Update App: ' + rStr +': ' + GetAppVersionStr(4);
        lbl_MainPanel.Caption := 'Update ' + rStr + ' App';
      end;

      if lStr = 'InstallApp' then  begin
        UpdaterInfo.InstallApp := rStr;
        UpdaterInfo.InstallAppOk := FileExists(rStr);
        if not UpdaterInfo.InstallAppOk then begin
          mmo_UpdateInfo.Lines.Add('ERROR: Install program not Found:');
          mmo_UpdateInfo.Lines.Add('  ' + QuotedStr(rStr));
          SetStatusBar('ERROR: Install program not Found!', clRed);
          btn_Update.Enabled := false;
        end
        else begin
          mmo_UpdateInfo.Lines.Add('Update Installer found:');
          mmo_UpdateInfo.Lines.Add('  ' + UpdaterInfo.InstallApp);
        end;
      end;

      if lStr = 'IconFile' then begin
        if FileExists(rStr) then begin
          img_Logo.Picture.LoadFromFile(rStr);
          mmo_UpdateInfo.Lines.Add('Logo file found:');
          mmo_UpdateInfo.Lines.Add('  ' + rStr);
        end;
      end;

      if lStr = 'VersionInfo' then begin
        lbl_Versions.Caption := rStr;
        UpdaterInfo.VersionInfo := rStr;
      end;

    end;
  finally
    FreeAndNil(UdInfo);
  end;
end;


function Tfrm_MainUpdate.GetAppVersionStr(aNumPlaces: Byte): string;
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then
    RaiseLastOSError;
  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then
    RaiseLastOSError;
  if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    RaiseLastOSError;
  case aNumPlaces of
    3: begin
      Result := Format('%d.%d.%d',
        [LongRec(FixedPtr.dwFileVersionMS).Hi,   //major
         LongRec(FixedPtr.dwFileVersionMS).Lo,   //minor
         LongRec(FixedPtr.dwFileVersionLS).Hi]); //release
    end;
    4: begin
      Result := Format('%d.%d.%d.%d',
        [LongRec(FixedPtr.dwFileVersionMS).Hi,  //major
         LongRec(FixedPtr.dwFileVersionMS).Lo,  //minor
         LongRec(FixedPtr.dwFileVersionLS).Hi,  //release
         LongRec(FixedPtr.dwFileVersionLS).Lo]); //build

    end;
    else begin
      Result := Format('%d.%d.%d.%d',
        [LongRec(FixedPtr.dwFileVersionMS).Hi,    //major
         LongRec(FixedPtr.dwFileVersionMS).Lo]);  //minor
    end;
  end;
end;


procedure Tfrm_MainUpdate.runUpdate;
//var
//    SEInfo: TShellExecuteInfo;
////    ExitCode: DWORD;
//    ExecuteFile{, ParamString, StartInString}: string;
begin
//  ExecuteFile:= UpdateInstallLocalPathFile;

  if not FileExists(UpdaterInfo.InstallApp) then  begin
    mmo_UpdateInfo.Lines.Add('Could NOT find the program installer');
    Exit;
  end;


  ShellExecute(handle,'open',PChar(UpdaterInfo.InstallApp), nil, nil, SW_SHOWNORMAL);
//  ShellExecute(handle,'open',PChar(UpdateInstallLocalPathFile), nil, nil, SW_SHOWNORMAL);
  Close;
//    FillChar(SEInfo, SizeOf(SEInfo), 0) ;
//    SEInfo.cbSize := SizeOf(TShellExecuteInfo) ;
//    //SEInfo.
//    with SEInfo do begin
////      fMask := SEE_MASK_NOCLOSEPROCESS;
//      Wnd := Application.Handle;
//      lpFile := PChar(ExecuteFile) ;
//      lpVerb := PWideChar('Open');
//      lpParameters := PWideChar(
// {
// ParamString can contain the
// application parameters.
// }
// // lpParameters := PChar(ParamString) ;
// {
// StartInString specifies the
// name of the working directory.
// If ommited, the current directory is used.
// }
// // lpDirectory := PChar(StartInString) ;
//      nShow := SW_SHOWNORMAL;
//    end;
//    if ShellExecuteEx(@SEInfo) then begin
////      repeat
////        Application.ProcessMessages;
////        GetExitCodeProcess(SEInfo.hProcess, ExitCode) ;
////      until (ExitCode <> STILL_ACTIVE) or Application.Terminated;
//      Close;
//    end
//    else begin
////      MessageDlg('Error updating Spending!', mtError, [mbOk],0);
//      Close;
//    end;
 end;

end.
