unit mainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  LCLTranslator, DateUtils, ComCtrls, Buttons, Menus, libusb, ComObj, ActiveX,
  Process, IniFiles, Variants, WinampControl, SynEdit, ListHighlighter,
  RegExpr, resource, versiontypes, versionresource, LCLIntf;

type
  TButtonEvent = procedure(Sender: TObject; ButtonMask: Byte) of object;

  // levels of logging events
  TLogLevel = (lvINFO, lvWARNING, lvERROR, lvCRITICAL);

  TMascotAction = (maALLOFF, maGREEN, maRED, maBEAK, maTILT, maFLAP);


  // a structure for saving application settings:
  TApplicationConfig = record
    SoxPath: String;                // Path to sox.exe
    Language: String;               // GUI language
    DoCommentSongChange: Boolean;   // comment when Winamp changes the song
    DoShowTaskbarPopups: Boolean;
    DoRandomTalk: Boolean;          // parrot will randmoly talk in intervals
    RandomTalkInterval: Integer;    // interval between random talk
    VoiceName: String;              // TTS voice name
    VoiceSpeed: Integer;            // TTS talking speed
    VoicePitch: Integer;            // TTS talking voice pitch
    VoiceLoudness: Integer;         // speaker loudness
    FlapDuration: Integer;          // default duration of wing flapping movement in ms
    TiltDuration: Integer;          // default duration of head tile movement in ms
    UntiltDuration: Integer;        // default duration of head un-tilt movement in ms
  end;

  { TForm1 }
  TForm1 = class(TForm)
    btnAllOff: TButton;
    btnClearLog: TSpeedButton;
    btnConnect: TButton;
    btnEnglish: TBitBtn;
    btnGerman: TBitBtn;
    btnGreen: TButton;
    btnPlayWave: TButton;
    btnPreviewNextSong: TSpeedButton;
    btnRed: TButton;
    btnSaveNoWinampLines: TSpeedButton;
    btnSaveNextSongLines: TSpeedButton;
    btnTestBeak: TButton;
    btnPreviewNoWinamp: TSpeedButton;
    btnTestFlaps: TButton;
    btnTestTilt: TButton;
    btnTTS: TSpeedButton;
    cbBalloonHints: TCheckBox;
    cbTalkEnabled: TCheckBox;
    cbCommentSongChange: TCheckBox;
    EditTTS: TEdit;
    GroupBoxVoiceInfo: TGroupBox;
    HideTimer: TTimer;
    SplashImage: TImage;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    lblAge: TLabel;
    lblGender: TLabel;
    lblLanguage: TLabel;
    lblLoudness: TLabel;
    lblPitch: TLabel;
    lblVoiceSpeed: TLabel;
    Panel4: TPanel;
    Panel5: TPanel;
    EditSoxPath: TEdit;
    GroupBox1: TGroupBox;
    GroupBox3: TGroupBox;
    Image1: TImage;
    ImageList2: TImageList;
    LabelAppName: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    LabelLink: TLabel;
    lblTalkTimer: TLabel;
    InfoMemo: TMemo;
    MenuItemClose: TMenuItem;
    MenuItemShowGui: TMenuItem;
    MenuItemSaySomething: TMenuItem;
    MenuItemToggleMute: TMenuItem;
    OpenDialogSox: TOpenDialog;
    Panel3: TPanel;
    btnTextPreview: TSpeedButton;
    btnSaveTalkLines: TSpeedButton;
    PanelNoWinamp: TPanel;
    PanelSongChange: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    PopupMenu1: TPopupMenu;
    SaveDialog1: TSaveDialog;
    Separator1: TMenuItem;
    btnSort: TSpeedButton;
    btnSaveLogFile: TSpeedButton;
    btnOpenSoXPath: TSpeedButton;
    SynEditTalk: TSynEdit;
    ImageList1: TImageList;
    Label1: TLabel;
    LoggingListView: TListView;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    Panel1: TPanel;
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    SynEditNoWinamp: TSynEdit;
    SynEditNewSong: TSynEdit;
    TabSheetInfo: TTabSheet;
    TabSheetTTS: TTabSheet;
    TabSheetParrot: TTabSheet;
    TabSheetWA: TTabSheet;
    TabSheetDbg: TTabSheet;
    TabSheetSettings: TTabSheet;
    TimerEventCheck: TTimer;
    TimerRandomTalk: TTimer;
    TrackBarTalkTimer: TTrackBar;
    TrackBarVoicePitch: TTrackBar;
    TrackBarVoiceSpeed: TTrackBar;
    TrackBarVoiceVolume: TTrackBar;
    TrayIcon1: TTrayIcon;
    VoiceComboBox: TComboBox;
    WinampControl: TWinampControl;
    procedure SplashImageClick(Sender: TObject);
    procedure btnClearLogClick(Sender: TObject);
    procedure btnEnglishClick(Sender: TObject);
    procedure btnGermanClick(Sender: TObject);
    procedure btnPlayWaveClick(Sender: TObject);
    procedure btnPreviewNextSongClick(Sender: TObject);
    procedure btnPreviewNoWinampClick(Sender: TObject);
    procedure btnRedClick(Sender: TObject);
    procedure btnSaveNextSongLinesClick(Sender: TObject);
    procedure btnSaveNoWinampLinesClick(Sender: TObject);
    procedure btnSaveTalkLinesClick(Sender: TObject);
    procedure btnTestBeakClick(Sender: TObject);
    procedure btnTestTiltClick(Sender: TObject);
    procedure btnTextPreviewClick(Sender: TObject);
    procedure btnTTSClick(Sender: TObject);
    procedure cbBalloonHintsChange(Sender: TObject);
    procedure cbCommentSongChangeChange(Sender: TObject);
    procedure cbTalkEnabledChange(Sender: TObject);
    procedure HideTimerTimer(Sender: TObject);
    procedure VoiceComboBoxChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure btnConnectClick(Sender: TObject);
    procedure btnGreenClick(Sender: TObject);
    procedure btnTestFlapsClick(Sender: TObject);
    procedure btnAllOffClick(Sender: TObject);
    procedure FormWindowStateChange(Sender: TObject);
    procedure LabelLinkClick(Sender: TObject);
    procedure MenuItemCloseClick(Sender: TObject);
    procedure MenuItemShowGuiClick(Sender: TObject);
    procedure MenuItemSaySomethingClick(Sender: TObject);
    procedure MenuItemToggleMuteClick(Sender: TObject);
    procedure btnSaveLogFileClick(Sender: TObject);
    procedure btnOpenSoXPathClick(Sender: TObject);
    procedure btnSortClick(Sender: TObject);
    procedure SynEditNewSongChange(Sender: TObject);
    procedure SynEditNewSongKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SynEditNoWinampKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SynEditTalkKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SynEditNoWinampChange(Sender: TObject);
    procedure SynEditTalkChange(Sender: TObject);
    procedure TimerEventCheckTimer(Sender: TObject);
    procedure TimerRandomTalkTimer(Sender: TObject);
    procedure TrackBarTalkTimerChange(Sender: TObject);
    procedure TrackBarVoiceSpeedChange(Sender: TObject);
    procedure TrackBarVoicePitchChange(Sender: TObject);
    procedure TrackBarVoiceVolumeChange(Sender: TObject);
    procedure TrayIcon1Click(Sender: TObject);
  private
    FOnButtonChange: TButtonEvent;
    SharedHighlighter: TListHighlighter;
    FLastWinampTitle: String;

    // USB stuff
    isUsbConnected: Boolean;
    pUsbContext: Plibusb_context;
    pUsbDevice: Plibusb_device_handle;

    // parrot stuff
    isParrotMuted: Boolean;
    isHeadTilted: Boolean;
    isParrotSpeaking: Boolean;
    mascotLedState: Byte; // set state of the LEDs (see MASCOT_FLAG_LEDRED and MASCOT_FLAG_LEDGREEN)

    // settings
    FApplicationConfig: TApplicationConfig;
    FIniFileName: String;

    // helper methods
    function ResourceVersionInfo: string;
    function PreProcessTTS(aText: string): string;
    function SplitCondition(const S: string; out LeftPart, OperatorPart, RightPart: string): Boolean;
    function SetDayMonthFromString(var dt: TDateTime; const s: string): Boolean;

    // custom event handlers
    procedure ParrotThreadFinished(Sender: TObject);
    procedure HandleButtonChange(Sender: TObject; ButtonMask: Byte); // called, when a button on the parrot is pressed
    procedure HandleSpeakingChange;

    // SAPI / text to speech methods
    procedure LoadSAPIVoices;
    procedure TextToWav(const aText, FileName: string);

    // parrot methods
    procedure ConnectToUSB; // establish a USB connection to the parrot
    procedure ParrotDoAction(mctAction: TMascotAction; duration: LongInt); // perform movement or set parrot LEDs
    procedure PlayWavOnParrot(const FileName: string);   // play a WAV file on the parrot speaker
    function ParrotSpeak(s: String): Boolean; // convert a string to speech and play it on the parrot speaker
    procedure ParrotSpeakNonBlocking(S: String); // same as ParrotSpeak, but running in a Thread

    // logging
    procedure LogEvent(const LogLevel: TLogLevel; const AText: String; const Timestamp: TDateTime);
    procedure SaveLogFile(FileName: TFileName);

    // settings
    procedure LoadConfig(const AFilePath: String);
    procedure SaveConfig(const AFilePath: String);

    // getter / setter methods
    procedure SetIsSpeaking(const Value: Boolean);
  public
    property OnButtonChange: TButtonEvent read FOnButtonChange write FOnButtonChange;
    property IsSpeaking: Boolean read isParrotSpeaking write SetIsSpeaking;
    procedure ReadButtons(out ButtonByte: Byte);
  end;

  { TButtonThread }
  TButtonThread = class(TThread)
  private
    procedure SyncReadButtons;
  public
    Form: TForm1;
  protected
    procedure Execute; override;
  end;

  { TParrotThread }
  TParrotThread = class(TThread)
  private
    FTextToSay: string;
  protected
    procedure Execute; override;
  public
    constructor Create(const TextToSay: string);
  end;


const
  // parrot robot protocol
  MASCOT_CMD_CTRL            = $BC00; // base command for the mascot
  MASCOT_FLAG_LEDRED         = $80;   // red LED on
  MASCOT_FLAG_LEDGREEN       = $40;   // green LED on
  MASCOT_CMD_BEAK            = $03;   // move beak
  MASCOT_CMD_TILT            = $04;   // tilt head (stays tilted until a "flap" is sent)
  MASCOT_CMD_FLAPS           = $05;   // flap wings
  MASCOT_CMD_ALLOFF          = $00;   // stop all movement and clear LEDs
  MASCOT_DEFAULT_TIMEOUT     = 5000;  // default timeout and movement duration in ms

  // USB stuff
  USB_MASCOT_VID = $03EE;  // vendor ID (Mitsumi)
  USB_MASCOT_PID = $FF01;  // product ID (Chatbird / PC-Mascot)


  // file names:
  PARROT_INIFILE = 'parrot.ini';  // where to load settings from
  TTS_WAV_FILE = 'tts.wav';       // where to save/load TTS output
  SOX_OUT_FILE = 'soxout.wav';    // where to save/load converted wave file from SoX
  TALK_FILE = 'talk.txt';         // where to load/save text, the parrot says randomly in intervals
  NEWSONG_FILE = 'newsong.txt';   // where to load/save text, the parrot says if Winamp changes the song
  NOWINAMP_FILE = 'nowinamp.txt'; // where to load/save text, the parrot says if Winamp is not started


var
  Form1: TForm1;
  ButtonThread: TButtonThread;

implementation

{$R *.lfm}

resourcestring
  { Day of the week }
  RsDoWSunday =    'Sunday';
  RsDoWMonday =    'Monday';
  RsDoWTuesday =   'Tuesday';
  RsDoWWednesday = 'Wednesday';
  RsDoWThursday =  'Thursday';
  RsDoWFriday =    'Friday';
  RsDoWSaturday =  'Saturday';


  { GUI texts }
  RsAppTitle =     'PC-Parrot';

  RsGender =      'Gender:';
  RsAge =         'Age:';
  RsLanguage =    'Language:';

  RsMute =        'Shut up!';
  RsUnMute =      'You may chatter again.';

  RsIdle =        'Parrot: idle';
  RsSpeaking =    'Parrot: active';

  RsUsbConnected =    'USB: connected';
  RsUsbDisconnected = 'USB: disconnected';


{ TButtonThread }

procedure TButtonThread.SyncReadButtons;
var
  b: Byte;
begin
  Form.ReadButtons(b);
  if Assigned(Form.OnButtonChange) then
    Form.OnButtonChange(Form, b);
end;

procedure TButtonThread.Execute;
begin
  while not Terminated do
  begin
    Synchronize(@SyncReadButtons);
    Sleep(1); // preserve some CPU time
  end;
end;


{ TParrotThread }

constructor TParrotThread.Create(const TextToSay: string);
begin
  inherited Create(True); // suspended
  FTextToSay := TextToSay;
  FreeOnTerminate := True; // does free its memory when completed
end;

procedure TParrotThread.Execute;
begin
  Form1.ParrotSpeak(FTextToSay);
end;



{ TForm1 }

// setter for isParrotSpeaking variable
procedure TForm1.SetIsSpeaking(const Value: Boolean);
begin
  if isParrotSpeaking = Value then
     Exit; // skip if isParrotSpeaking hasn't changed

  isParrotSpeaking := Value;
  HandleSpeakingChange;
end;


// event handler method for change of isParrotSpeaking variable
procedure TForm1.HandleSpeakingChange;
begin
  if isParrotSpeaking then begin
    // disable all buttons which can let the parrot speak
    StatusBar1.Panels[1].Text:= RsSpeaking;
    btnPlayWave.Enabled:= False;
    btnTextPreview.Enabled:= False;
    btnTTS.Enabled:= False;
    btnPreviewNextSong.Enabled:= False;
    btnPreviewNoWinamp.Enabled:= False;
  end else begin
    // re-enable all buttons which can let the parrot speak
    StatusBar1.Panels[1].Text:= RsIdle;
    btnPlayWave.Enabled:= True;
    btnTextPreview.Enabled:= True;
    btnTTS.Enabled:= True;
    btnPreviewNextSong.Enabled:= True;
    btnPreviewNoWinamp.Enabled:= True;
  end;
end;


// logs an event in the LoggingListView
procedure TForm1.LogEvent(const LogLevel: TLogLevel; const AText: String; const Timestamp: TDateTime);
var
  Item: TListItem;
  AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond: Word;
begin
  DecodeDateTime(Timestamp, AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond);

  Item := LoggingListView.Items.Add;
  Item.Caption := IntToStr(LoggingListView.Items.Count);
  Item.Subitems.Add(AText);
  case LogLevel of
    lvINFO: Item.Subitems.Add('INFO');
    lvWARNING: Item.Subitems.Add('WARNING');
    lvERROR: Item.Subitems.Add('ERROR');
    else
      Item.Subitems.Add('CRITICAL');
  end;
  Item.Subitems.Add(Format('%.04d-%.02d-%.02d %.02d:%.02d:%.02d.%.03d', [AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond]));
  Item.MakeVisible(False);  // scroll to new entry
end;


// helper function to set day and month in a TDateTime object
function TForm1.SetDayMonthFromString(var dt: TDateTime; const s: string): Boolean;
var
  d, m: Word;
  y: Word;
begin
  Result := False;
  // keep the year
  y := YearOf(dt);

  // try to get day and month from 'dd.mm'
  if Length(s) < 4 then Exit; // must be at least 'd.m'
  try
    d := StrToIntDef(Copy(s, 1, Pos('.', s)-1), 0);
    m := StrToIntDef(Copy(s, Pos('.', s)+1, 2), 0);
    if (d < 1) or (d > 31) or (m < 1) or (m > 12) then Exit;

    // construct new datetime value
    dt := EncodeDate(y, m, d);
    Result := True;
  except
    // refurn false on error
    Result := False;
  end;
end;


// helpert function to split a string into three parts: left of operator, operator and right of operator
function TForm1.SplitCondition(const S: string; out LeftPart, OperatorPart, RightPart: string): Boolean;
var
  op: string;
  p: Integer;
const
  OPERATORS: array[0..5] of string = ('<=', '>=', '!=', '<', '>', '=');
begin
  Result := False;
  LeftPart := '';
  OperatorPart := '';
  RightPart := '';

  for op in OPERATORS do
  begin
    p := Pos(op, S);
    if p > 0 then
    begin
      LeftPart := Trim(Copy(S, 1, p - 1));
      OperatorPart := op;
      RightPart := Trim(Copy(S, p + Length(op), MaxInt));
      Result := True;
      Exit;
    end;
  end;
end;

// Before a text is given to TTS we look for conditions and keywords in PreProcessTTS.
// This function can execute parrot commands (such as "flap the wings"), and play sound files on the parrots speaker.
// This function will remove known commands/keywords and return the remaining text to be converted to speech (TTS)
function TForm1.PreProcessTTS(aText: string): string;
var
  I: Integer;
  DoW: Word;
  Year, Month, Day: Word;
  Hour, Minute, Second, MilliSecond: Word;
  CurrentDateTime: TDateTime;
  s, hs: string;
  hi: Integer;
  Rex: TRegExpr;
  strCondition: String;
  isCondition: Boolean;
  L, Op, R: string;
  dt1, dt2: TDateTime;
  uptime: Longword;
begin
  S:= aText.Trim;

  CurrentDateTime := Now;
  DecodeDateTime(CurrentDateTime, Year, Month, Day, Hour, Minute, Second, MilliSecond);

  // get uptime in minutes
  uptime := GetTickCount64;
  uptime := trunc(uptime / 60000);

  // abort if it is a comment
  if s.StartsWith(';') then begin
    result:= '';
    Exit;
  end;


  // Handle keywords in [] braces
  isCondition:= True;
  Rex := TRegExpr.Create;
  strCondition:= '';
  try
    Rex.Expression := '^(\[.+\])'; // look for a [..] brace at the start of the string
    if Rex.Exec(S) then begin
      strCondition := Rex.Match[1];
    end;
  finally
    Rex.Free;
  end;

  if not strCondition.IsEmpty then begin
    s:= s.Replace(strCondition, '').Trim;
    strCondition:= strCondition.Replace('[', '');
    strCondition:= strCondition.Replace(']', '');

    strCondition:= strCondition.ToUpper;  // ok, because in Windows, it doesn't matter if its "fart.wav" or "FART.WAV"

    if SplitCondition(strCondition, L, Op, R) then begin
      // the condition contains an operator

      if (L = 'WAV') then begin // play sound file
        isCondition:= False;
        // R should be the file path + name
        if Op = '=' then begin   // only = operator is supported
          if FileExists(R) then
            PlayWavOnParrot(R);
            isCondition:= True;
        end;
      end else if (L = 'DOW') then begin  // compare with Day Of Week
        DoW:= DayOfWeek(CurrentDateTime);
        // R should be '1' to '7' (sunday to saturday)
        hi:= StrToInt(R);
        if Op = '<' then
           isCondition:= DoW < hi;
        if Op = '>' then
           isCondition:= DoW > hi;
        if Op = '<=' then
           isCondition:= DoW <= hi;
        if Op = '>=' then
           isCondition:= DoW >= hi;
        if Op = '!=' then
           isCondition:= DoW <> hi;
        if Op = '=' then
           isCondition:= DoW = hi;
      end else if (L = 'TIME') then begin // compare with time
        dt1:= CurrentDateTime;
        dt2:= StrToTime(R);
        if Op = '<' then
           isCondition:= Frac(dt1) < Frac(dt2);
        if Op = '>' then
           isCondition:= Frac(dt1) > Frac(dt2);
        if Op = '<=' then
           isCondition:= Frac(dt1) <= Frac(dt2);
        if Op = '>=' then
           isCondition:= Frac(dt1) >= Frac(dt2);
        if Op = '!=' then
           isCondition:= Frac(dt1) <> Frac(dt2);
        if Op = '=' then
           isCondition:= Frac(dt1) = Frac(dt2);
      end else if (L = 'DATE') then begin // compare with date
        dt1:= CurrentDateTime;
        dt2:= CurrentDateTime;
        // string R must be in the format 'dd.mm'
        if SetDayMonthFromString(dt2, R) then begin
          if Op = '<' then
             isCondition:= Int(dt1) < Int(dt2);
          if Op = '>' then
             isCondition:= Int(dt1) > Int(dt2);
          if Op = '<=' then
             isCondition:= Int(dt1) <= Int(dt2);
          if Op = '>=' then
             isCondition:= Int(dt1) >= Int(dt2);
          if Op = '!=' then
             isCondition:= Int(dt1) <> Int(dt2);
          if Op = '=' then
             isCondition:= Int(dt1) = Int(dt2);
        end else begin
          LogEvent(lvERROR, 'Could not parse date ' + R, now);
          isCondition:= False;
        end;
      end else if (L = 'UPTIME') then begin // compare with system uptime
        // string R must be an integer with the compare time in minutes (e.g. '60' for 1 hour)
        hi:= StrToInt(R);
        if Op = '<' then
           isCondition:= uptime < hi;
        if Op = '>' then
           isCondition:= uptime > hi;
        if Op = '<=' then
           isCondition:= uptime <= hi;
        if Op = '>=' then
           isCondition:= uptime >= hi;
        if Op = '!=' then
           isCondition:= uptime <> hi;
        if Op = '=' then
           isCondition:= uptime = hi;
      end;
    end // the condition contains an operator
    else begin
      // the condition does not contain an operator
      isCondition:= False; // default for the following checks
      if (strCondition = 'MUSIC') and (WinampControl.IsWinampRunning) and (WinampControl.IsPlaying) then
        isCondition:= True  // Winamp plays music
      else
      if (strCondition = 'WAPLAY') and (WinampControl.IsWinampRunning) and (WinampControl.IsPlaying) then
        isCondition:= True // Winamp plays music
      else
      if (strCondition = 'WASTOP') and ((not WinampControl.IsWinampRunning) or (not WinampControl.IsPlaying)) then
        isCondition:= True;  // Winamp not running or not playing music
    end; // the condition does not contain an operator

    // abort if the condition does not apply
    if not isCondition then begin
      LogEvent(lvINFO, 'Wont say that, because condition "'+ strCondition+ '" does not apply.', now);
      result:= '';
      Exit;
    end;

  end;  // if not strCondition.IsEmpty


  // Handle keywords in <>

  // day of week
  while (Pos('<DOW>', S.ToUpper) <> 0 ) do begin
    I:= Pos('<DOW>', S.ToUpper);
    Delete(S, I, Length('<DOW>'));
    DoW:= DayOfWeek(CurrentDateTime);
    case DoW of
     1: Insert(RsDoWSunday,    S, I);
     2: Insert(RsDoWMonday,    S, I);
     3: Insert(RsDoWTuesday,   S, I);
     4: Insert(RsDoWWednesday, S, I);
     5: Insert(RsDoWThursday,  S, I);
     6: Insert(RsDoWFriday,    S, I);
     7: Insert(RsDoWSaturday,  S, I);
    end;
  end;

  // time
  while (Pos('<TIME>', S.ToUpper) <> 0 ) do begin
    I:= Pos('<TIME>', S.ToUpper);
    Delete(S, I, Length('<TIME>'));
    if FApplicationConfig.Language = 'de' then begin
      hs:= FormatDateTime('h', CurrentDateTime) + ' Uhr ' + FormatDateTime('n', CurrentDateTime);
    end else begin
      hs:= FormatDateTime('h', CurrentDateTime) + ' ' + FormatDateTime('n', CurrentDateTime);
    end;
    Insert(hs, S, I);
  end;

  // date
  while (Pos('<DATE>', S.ToUpper) <> 0 ) do begin
    I:= Pos('<DATE>', S.ToUpper);
    Delete(S, I, Length('<DATE>'));
    hs:= inttostr(Day);
    // a date like e.g. "1 of August 2026" would be prononced by TTS like "one of August twothousandtwentysix"
    // therefore, we ourselves need to make a few adjustments.
    if FApplicationConfig.Language = 'de' then begin
      if (Day = 1) then
        hs:= 'erste '
      else if (Day = 3) then
        hs:= 'dritte '
      else if (Day = 6) then
        hs:= 'sechste '
      else if (Day = 7) then
        hs:= 'siebte '
      else if (Day = 8) then
        hs:= 'achte '
      else if (Day < 20) then
        hs:= hs +'te '
      else
        hs:= hs+'ste ';
      hs:= hs + FormatDateTime('mmmm yyyy', CurrentDateTime);
    end else begin
      if (Day = 1) then
        hs:= 'first '
      else if (Day = 2) then
        hs:= 'second '
      else if (Day = 3) then
        hs:= 'third '
      else if (Day = 21) then
        hs:= 'twenty-first '
      else if (Day = 22) then
        hs:= 'twenty-second '
      else if (Day = 23) then
        hs:= 'twenty-third '
      else if (Day = 31) then
        hs:= 'thirty-first '
      else
        hs:= hs+'ty ';
      hs:= hs + 'of '+ FormatDateTime('mmmm yyyy', CurrentDateTime);
    end;
    Insert(hs, S, I);
  end;

  // parrot head tilt
  while (Pos('<TILT>', S.ToUpper) <> 0 ) do begin
    I:= Pos('<TILT>', S.ToUpper);
    Delete(S, I, Length('<TILT>'));
    ParrotDoAction(maTILT, FApplicationConfig.TiltDuration);
    Sleep(FApplicationConfig.TiltDuration);
  end;

  // parrot wing flap
  while (Pos('<FLAP>', S.ToUpper) <> 0 ) do begin
    I:= Pos('<FLAP>', S.ToUpper);
    Delete(S, I, Length('<FLAP>'));
    ParrotDoAction(maFLAP, 2000);
    Sleep(1000);
  end;

  if not s.isEmpty then begin
    // create a string for ballon hints which is free of any TTS commands (such as <rate speed="2">faster</rate>)
    Rex := TRegExpr.Create;
    hs:= s;
    try
      // <[^>]*>  bedeutet:
      // <        = öffnende spitze Klammer
      // [^>]*    = beliebige Zeichen außer >
      // >        = schließende spitze Klammer
      Rex.Expression := '<[^>]*>';
      hs := Rex.Replace(hs, '', True);
    finally
      Rex.Free;
    end;
    caption:= hs;  // display text to be spoken in the window caption
    if FApplicationConfig.DoShowTaskbarPopups then begin
      TrayIcon1.BalloonHint:= hs;
      TrayIcon1.BalloonTimeout:= 3000;
      TrayIcon1.ShowBalloonHint; // display text to be spoken as a ballon hint
    end;
  end
  else
    if isHeadTilted then begin
      ParrotDoAction(maFLAP, FApplicationConfig.UntiltDuration); // un-tilt the head
      Sleep(FApplicationConfig.UntiltDuration);
    end;

  result:= S;
end;


// helper method to obtain a list of all available SAPI voices in VoiceComboBox
procedure TForm1.LoadSAPIVoices;
var
  TTS: OleVariant;
  Token: OleVariant;
  i: Integer;
  VoiceName: String;
  vIndex: Integer;
begin
  vIndex:= -1;
  VoiceComboBox.Items.Clear;
  TTS := CreateOleObject('SAPI.SpVoice');

  for i := 0 to TTS.GetVoices.Count - 1 do begin
    Token := TTS.GetVoices.Item(i);
    VoiceName:= Token.GetDescription;
    VoiceComboBox.Items.Add(VoiceName);
    if (vIndex < 0) and (VoiceName = FApplicationConfig.VoiceName) then
      vIndex:= i;
  end;

  if (vIndex >= 0) then
    VoiceComboBox.ItemIndex := vIndex // select voice as configured in settings
  else
  if VoiceComboBox.Items.Count > 0 then
    VoiceComboBox.ItemIndex := 0; // select first voice
end;


// load settings from file
procedure TForm1.LoadConfig(const AFilePath: String);
var
  IniFile: TIniFile;
begin
  try
    IniFile := TIniFile.Create(AFilePath);

    { application section }
    FApplicationConfig.SoxPath:=  IniFile.ReadString('APPLICATION', 'SoxPath', '%ProgramFiles%\sox-14-4-2\sox.exe');
    FApplicationConfig.Language := IniFile.ReadString('APPLICATION', 'Language', 'en');

    { voice section }
    FApplicationConfig.VoiceName:= IniFile.ReadString('VOICE', 'VoiceName', 'Microsoft Sam');
    FApplicationConfig.VoiceLoudness:= IniFile.ReadInteger('VOICE', 'VoiceLoudness', 30);
    FApplicationConfig.VoicePitch:= IniFile.ReadInteger('VOICE', 'VoicePitch', 0);
    FApplicationConfig.VoiceSpeed:= IniFile.ReadInteger('VOICE', 'VoiceSpeed', 0);

    { parrot section }
    FApplicationConfig.DoCommentSongChange := IniFile.ReadBool('PARROT', 'DoCommentSongChange', True);
    FApplicationConfig.DoShowTaskbarPopups := IniFile.ReadBool('PARROT', 'DoShowTaskbarPopups', True);
    FApplicationConfig.DoRandomTalk:= IniFile.ReadBool('PARROT', 'DoRandomTalk', False);
    FApplicationConfig.RandomTalkInterval:= IniFile.ReadInteger('PARROT', 'TalkInterval', 30);

    { robot section }
    FApplicationConfig.FlapDuration:= IniFile.ReadInteger('ROBOT', 'FlapDuration', 2000);
    FApplicationConfig.TiltDuration:= IniFile.ReadInteger('ROBOT', 'TiltDuration', 500);
    FApplicationConfig.UntiltDuration:= IniFile.ReadInteger('ROBOT', 'UntiltDuration', 500);

  finally
    IniFile.Free;
  end;
end;


// save settings to file
procedure TForm1.SaveConfig(const AFilePath: String);
var
  IniFile: TIniFile;
begin
  try
    IniFile := TIniFile.Create(AFilePath);

    { application section }
    IniFile.WriteString('APPLICATION', 'SoxPath', FApplicationConfig.SoxPath);
    IniFile.WriteString('APPLICATION', 'Language', FApplicationConfig.Language);

    { voice section }
    IniFile.WriteString('VOICE', 'VoiceName', FApplicationConfig.VoiceName);
    IniFile.WriteInteger('VOICE', 'VoiceLoudness', FApplicationConfig.VoiceLoudness);
    IniFile.WriteInteger('VOICE', 'VoicePitch', FApplicationConfig.VoicePitch);
    IniFile.WriteInteger('VOICE', 'VoiceSpeed', FApplicationConfig.VoiceSpeed);

    { parrot section }
    IniFile.WriteBool('PARROT', 'DoCommentSongChange', FApplicationConfig.DoCommentSongChange);
    IniFile.WriteBool('PARROT', 'DoShowTaskbarPopups', FApplicationConfig.DoShowTaskbarPopups);
    IniFile.WriteBool('PARROT', 'DoRandomTalk', FApplicationConfig.DoRandomTalk);
    IniFile.WriteInteger('PARROT', 'TalkInterval', FApplicationConfig.RandomTalkInterval);

    { robot section }
    IniFile.ReadInteger('ROBOT', 'FlapDuration', FApplicationConfig.FlapDuration);
    IniFile.ReadInteger('ROBOT', 'TiltDuration', FApplicationConfig.TiltDuration);
    IniFile.ReadInteger('ROBOT', 'UntiltDuration', FApplicationConfig.UntiltDuration);

  finally
    IniFile.Free;
  end;
end;


// helper method to obtain application version number as string
function TForm1.ResourceVersionInfo: string;
var
  Stream: TResourceStream;
  vr: TVersionResource;
  fi: TVersionFixedInfo;
begin
  Result := '';
  { This raises an exception if version info has not been incorporated into the
    binary (Lazarus Project -> Project Options -> Version Info -> Version numbering). }
  Stream := TResourceStream.CreateFromID(HINSTANCE, 1, PChar(RT_VERSION));
  try
    vr := TVersionResource.Create;
    try
      vr.SetCustomRawDataStream(Stream);
      fi := vr.FixedInfo;
      Result := Trim(Format('%d.%d.%d.%d', [fi.FileVersion[0],
        fi.FileVersion[1], fi.FileVersion[2], fi.FileVersion[3]]));
    finally
      vr.Free
    end;
  finally
    Stream.Free
  end;
end;


procedure TForm1.FormCreate(Sender: TObject);
var
  IniFilePath: String; // Inifile to load
begin
  // init FApplicationConfig in case on ini file exists
  FApplicationConfig.DoCommentSongChange:= False;
  FApplicationConfig.DoRandomTalk:= False;
  FApplicationConfig.DoShowTaskbarPopups:= True;
  FApplicationConfig.Language:= 'en';
  FApplicationConfig.RandomTalkInterval:= 30;
  FApplicationConfig.SoxPath:= '';
  FApplicationConfig.VoiceLoudness:= 20;
  FApplicationConfig.VoiceName:= 'Microsoft Sam';
  FApplicationConfig.VoicePitch:= 0;
  FApplicationConfig.VoiceSpeed:= 0;
  FApplicationConfig.FlapDuration:= 2000;
  FApplicationConfig.TiltDuration:= 500;
  FApplicationConfig.UntiltDuration:= 500;

  // init variables
  FIniFileName := ExtractFilePath(Application.ExeName) + PARROT_INIFILE; // default ini file
  pUsbContext:= nil;
  pUsbDevice:= nil;
  mascotLedState:= $00; // off
  isUsbConnected:= False;
  isParrotMuted:= False;
  IsSpeaking:= False;
  FLastWinampTitle:= '';
  OnButtonChange := @HandleButtonChange;

  // assign highlighter
  SharedHighlighter := TListHighlighter.Create(Self);
  SynEditTalk.Highlighter:= SharedHighlighter;
  SynEditNewSong.Highlighter:= SharedHighlighter;
  SynEditNoWinamp.Highlighter:= SharedHighlighter;

  // current directory shall always be where the exe is
  ChDir(ExtractFilePath(Application.ExeName));

  // --- Load settings ---
  IniFilePath := FIniFileName;
  if (ExtractFileName(IniFilePath) = IniFilePath) then
    IniFilePath := ExtractFilePath(Application.ExeName) + FIniFileName;
  if (FileExists(IniFilePath)) then begin
    LogEvent(lvINFO, 'Loading settings from "'+IniFilePath+'"', Now);
    LoadConfig(IniFilePath);
  end else
    LogEvent(lvERROR, 'Could not load settings from file "'+IniFilePath+'"', Now);

  Randomize;

  // do language related stuff
  SetDefaultLang(FApplicationConfig.Language);
  Caption:= RsAppTitle + ' v' + ResourceVersionInfo;
  TrayIcon1.Hint:= RsAppTitle;
  MenuItemToggleMute.Caption:= RsMute;
  if FApplicationConfig.Language = 'de' then
    InfoMemo.lines.LoadFromFile('info_de.txt', TEncoding.UTF8)
  else
    InfoMemo.lines.LoadFromFile('info_en.txt', TEncoding.UTF8);

  // get a list of SAPI voices
  LoadSAPIVoices;
  VoiceComboBoxChange(Self);


  // Set controls according to settings
  TrackBarVoiceVolume.Position:= FApplicationConfig.VoiceLoudness;
  TrackBarVoicePitch.Position:= FApplicationConfig.VoicePitch;
  TrackBarVoiceSpeed.Position:= FApplicationConfig.VoiceSpeed;
  EditSoxPath.Text:= FApplicationConfig.SoxPath;

  cbBalloonHints.Checked:= FApplicationConfig.DoShowTaskbarPopups;
  cbCommentSongChange.Checked:= FApplicationConfig.DoCommentSongChange;
  cbTalkEnabled.Checked:= FApplicationConfig.DoRandomTalk;
  TrackBarTalkTimer.Position:= FApplicationConfig.RandomTalkInterval;

  // try to connect to the PC-Mascot via USB
  ConnectToUSB;

  if (isUsbConnected) then begin
    // start button thread
    ButtonThread := TButtonThread.Create(True); // still suspended
    ButtonThread.FreeOnTerminate := False;
    ButtonThread.Form := Self;
    ButtonThread.Start;
    LogEvent(lvINFO, 'Button thread started.', now);
    ParrotDoAction(maGREEN, 0);  // switch on the green LED to indicate we are connected
  end;

  // load parrot texts
  SynEditTalk.Lines.LoadFromFile(TALK_FILE, TEncoding.UTF8);
  SynEditNewSong.Lines.LoadFromFile(NEWSONG_FILE, TEncoding.UTF8);
  SynEditNoWinamp.Lines.LoadFromFile(NOWINAMP_FILE, TEncoding.UTF8);

  HandleSpeakingChange;  // update status text in status bar

  if (not FileExists(FIniFileName)) then begin
    // no ini-file found -> display normally and show settings tab
    PageControl1.ActivePage:= TabSheetSettings;
  end;

  HideTimer.Enabled:= True;

end;



// user pressed save button for SynEditTalk
procedure TForm1.btnSaveTalkLinesClick(Sender: TObject);
begin
  SynEditTalk.Lines.SaveToFile(TALK_FILE,TEncoding.UTF8);
  btnSaveTalkLines.Enabled:= False;
end;


// handles the application reaction to button presses on the parrot robot
procedure TForm1.HandleButtonChange(Sender: TObject; ButtonMask: Byte);
var
  strNoWinamp: String;
  i: Integer;
begin
  if WinamPcontrol.IsWinampRunning then begin
    if ButtonMask and $01 <> 0 then begin
      if WinampControl.IsPlaying then begin
        if WinampControl.Stop then
          LogEvent(lvINFO, 'Winamp stops playing.', now)
        else
          LogEvent(lvERROR, 'Failed to stop playback.', now);
      end else begin
        if WinampControl.StartPlay then
          LogEvent(lvINFO, 'Winamp starts playing.', now)
        else
          LogEvent(lvERROR, 'Failed to start playback.', now);
      end;
    end;

    if ButtonMask and $02 <> 0 then begin
      if WinampControl.PrevTrack then
        LogEvent(lvINFO, 'Winamp plays previous song.', now)
      else
        LogEvent(lvINFO, 'Failed to go to previous song.', now);
    end;

    if ButtonMask and $04 <> 0 then begin
      if WinampControl.NextTrack then
        LogEvent(lvINFO, 'Winamp plays next song.', now)
      else
        LogEvent(lvERROR, 'Failed to go to next song.', now);
      end;
  end else begin
    if (ButtonMask and $07 <> 0) then begin

      // pick a line from the memo and say it
      if (SynEditNoWinamp.Lines.Count > 0) then begin
        i:= Random(SynEditNoWinamp.Lines.Count);
        strNoWinamp:= Trim(UTF8Decode(SynEditNoWinamp.Lines[i]));
        if (not strNoWinamp.IsEmpty) then begin
           if (not isParrotMuted)
             then  ParrotSpeakNonBlocking(strNoWinamp);
        end else
          LogEvent(lvWARNING, 'No message "Winamp is not running" found.', now);
      end;
    end;
  end;

  if ButtonMask and $08 <> 0 then begin
     MenuItemToggleMuteClick(Self);
    LogEvent(lvINFO, 'Button 4 pressed.', now)
  end;

  if ButtonMask = 0 then begin
    // no button pressed / all buttons released
  end;
end;


// saves the log entries in a file
procedure TForm1.SaveLogFile(FileName: TFileName);
var
  AStringList: TStringList;
  I: Integer;
begin
  AStringList := TStringList.Create;
  try
    AStringList.Add('#;Message;Level;Time');
    for I := 0 to (LoggingListView.Items.Count - 1) do
      AStringList.Add(LoggingListView.Items[I].Caption + ',' + // #
        '"' + LoggingListView.Items[i].SubItems[0] + '";' +    // Message
        LoggingListView.Items[i].SubItems[1] + ';' +           // Level
        LoggingListView.Items[i].SubItems[2]);                 // Time
    try
      AStringList.SaveToFile(FileName);
    except
      on E: Exception do
        LogEvent(lvERROR, 'Could not save log file. Error: ' + E.Message, Now);
    end;
  finally
    AStringList.Free;
  end;
end;


procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  SaveConfig(FIniFileName);
  SaveLogFile('parrot.log');

  if btnSaveTalkLines.Enabled then
    btnSaveTalkLinesClick(Self); // button is enabled > then text has changed; calling its click handler will save to file

  if btnSaveNoWinampLines.Enabled then
    btnSaveNoWinampLinesClick(Self); // button is enabled > then text has changed; calling its click handler will save to file

  if btnSaveNextSongLines.Enabled then
    btnSaveNextSongLinesClick(Self); // button is enabled > then text has changed; calling its click handler will save to file


  if (isUsbConnected) then
    ParrotDoAction(maALLOFF, 1000); // switch all LEDs off and stop any movement

  if Assigned(SharedHighlighter) then begin
    SharedHighlighter.Free;
    SharedHighlighter:= nil;
  end;

  if Assigned(ButtonThread) then begin
    ButtonThread.Terminate;
    ButtonThread.WaitFor;
    ButtonThread.Free;
    ButtonThread := nil;
  end;

  if (isUsbConnected) then begin
    libusb_release_interface(pUsbDevice, 0);
    libusb_close(pUsbDevice);
    libusb_exit(pUsbContext);
    pUsbDevice := nil;
    isUsbConnected := False;
    StatusBar1.Panels[0].Text:= RsUsbDisconnected;
  end;
end;

// read status of buttons on parrot robot via USB endpoint 0x81
procedure TForm1.ReadButtons(out ButtonByte: Byte);
var
  transferred: LongInt;
  rc: LongInt;
const
  BTN_TIMEOUT = 5; // timeout in ms
begin
  if (not isUsbConnected) then
    Exit;

  transferred := 0;
  rc := libusb_interrupt_transfer(pUsbDevice, $81, @ButtonByte, 1, transferred, BTN_TIMEOUT);
  if (rc <> 0) or (transferred <> 1) then
    ButtonByte := 0;
end;


// connect to the parrot robot via USB
procedure TForm1.ConnectToUSB;
begin
  if (isUsbConnected) then begin
    LogEvent(lvERROR, 'Cannot connect via USB - already connected.', now);
    Exit;
  end;

  if libusb_init(pUsbContext) <> 0 then begin
    LogEvent(lvERROR, 'libusb_init failed.', now);
    Exit;
  end;

  pUsbDevice := libusb_open_device_with_vid_pid(pUsbContext, USB_MASCOT_VID, USB_MASCOT_PID);
  if pUsbDevice = nil then begin
    LogEvent(lvERROR, 'USB Device NOT found.', now);
    libusb_exit(pUsbContext);
    Exit;
  end;

  LogEvent(lvINFO, 'USB device opened successfully.', now);

  if libusb_claim_interface(pUsbDevice, 0) <> 0 then begin
    LogEvent(lvERROR, 'Claiming the USB interface failed.', now);
    libusb_close(pUsbDevice);
    libusb_exit(pUsbContext);
    Exit;
  end;

  isUsbConnected := True;
  StatusBar1.Panels[0].Text:= RsUsbConnected;

  LogEvent(lvINFO, 'USB interface claimed. Device is connected.', now);

end;

// user wants to connect to the robot
procedure TForm1.btnConnectClick(Sender: TObject);
begin
  ConnectToUSB;
end;


// user wants to connect toggle the green LED
procedure TForm1.btnGreenClick(Sender: TObject);
var
  rc: Integer;
begin
  if (not isUsbConnected) then
    Exit;

  mascotLedState:= mascotLedState XOR MASCOT_FLAG_LEDGREEN;

  rc := libusb_control_transfer(pUsbDevice, $40, $01, MASCOT_CMD_CTRL + mascotLedState,
    0, nil, 0, 0);

  // for debugging: LogEvent(lvINFO, 'LED cmd rc=' + IntToStr(rc), now);
end;


// Perform parrot robot action such as flapping the wings or controlling the LEDs
procedure TForm1.ParrotDoAction(mctAction: TMascotAction; duration: LongInt);
var
  rc: Integer;
  cmd: Word;
begin
  if (not isUsbConnected) then
    Exit;

  cmd:= MASCOT_CMD_CTRL;

  case (mctAction) of
    maALLOFF:      mascotLedState:= $00;
    maRED:         mascotLedState:= MASCOT_FLAG_LEDRED;
    maGREEN:       mascotLedState:= MASCOT_FLAG_LEDGREEN;
    maBEAK:        cmd:= cmd + MASCOT_CMD_BEAK;
    maTILT:        cmd:= cmd + MASCOT_CMD_TILT;
    maFLAP:        cmd:= cmd + MASCOT_CMD_FLAPS;
  end;

  if mctAction = maTILT then
    isHeadTilted:= True; // remember that the head is tilted

  if mctAction = maFLAP then
    isHeadTilted:= False; // flapping always un-tilts the robot head

  cmd:= cmd + mascotLedState;

  rc := libusb_control_transfer(pUsbDevice, $40, $01, cmd, duration div 2, nil, 0, duration div 2);
  // for debugging: LogEvent(lvINFO, 'Mascot action rc=' + IntToStr(rc), now);
end;

procedure TForm1.btnAllOffClick(Sender: TObject);
begin
  ParrotDoAction(maALLOFF, 500);
end;

procedure TForm1.FormWindowStateChange(Sender: TObject);
begin
  if (wsMinimized = WindowState) then begin
    Hide;  // hide to taskbar tray
  end;
end;

procedure TForm1.LabelLinkClick(Sender: TObject);
begin
  OpenURL('https://cypax.net');
end;

procedure TForm1.MenuItemCloseClick(Sender: TObject);
begin
  Form1.Close;
end;

// user wants to sse TForm1
procedure TForm1.MenuItemShowGuiClick(Sender: TObject);
begin
  Show;
  Application.ShowMainForm := True;
  Application.Restore;
end;

// user wants the parrot to say something randomly (ignore isParrotMuted in this case)
procedure TForm1.MenuItemSaySomethingClick(Sender: TObject);
var
  S: string;
  I: Integer;
begin
  i:= Random(SynEditTalk.Lines.Count);

  S:= UTF8Decode(SynEditTalk.Lines[i]);
  ParrotSpeakNonBlocking(S);
end;

// mute/unmute the parrot
procedure TForm1.MenuItemToggleMuteClick(Sender: TObject);
var
  rc: Integer;
begin
  if isParrotMuted then begin
    isParrotMuted:= False;
    MenuItemToggleMute.Caption:= RsMute;
    MenuItemToggleMute.ImageIndex:= 7; // icon mute
    mascotLedState:= mascotLedState AND (not MASCOT_FLAG_LEDRED);
    rc := libusb_control_transfer(pUsbDevice, $40, $01, MASCOT_CMD_CTRL + mascotLedState, 0, nil, 0, 0);
  end else begin
    isParrotMuted:= True;
    MenuItemToggleMute.Caption:= RsUnMute;
    MenuItemToggleMute.ImageIndex:= 8; // icon unmute
    mascotLedState:= mascotLedState OR MASCOT_FLAG_LEDRED;
    rc := libusb_control_transfer(pUsbDevice, $40, $01, MASCOT_CMD_CTRL + mascotLedState, 0, nil, 0, 0);
  end;
end;

procedure TForm1.btnTestFlapsClick(Sender: TObject);
begin
  ParrotDoAction(maFLAP, 1000);
end;


// Plays a wave file on the parrot robot speaker.
// The wave file is passed to SoX and, converted to 12kHz PCM mono.
// The resulting output in SOX_OUT_FILE is then transmitted via USB to the parrots speaker.
// The parrot robot is instructed to move its beak for the duration of the sound file playing.
procedure TForm1.PlayWavOnParrot(const FileName: string);
var
  RawData, PCM12k: array of SmallInt;
  fs: TFileStream;
  riff, chunkID: array[0..3] of AnsiChar;
  chunkSize, fmtSize, dataSize: Cardinal;
  audioFormat, numChannels, bitsPerSample: Word;
  sampleRate: Cardinal;
  i, numSamples, outLen: Integer;
  MaxBlockSize: Integer = 1024;
  transferred: LongInt;
  rc: LongInt;
  block: array[0..1023] of Byte;
  volumeFactor: Double;
  P: TProcess;
  durationMS: LongInt;
begin
  if (not isUsbConnected) then begin
    LogEvent(lvERROR, 'Cannot play WAV - not connected.', now);
    Exit;
  end;

  // check if sox.exe is actually present
  if not FileExists(FApplicationConfig.SoxPath) then begin
    LogEvent(lvERROR, 'Could not find sox.exe', now);
    Exit;
  end;

  // convert the sound file to the data format, the parrot robot can handle
  P := TProcess.Create(nil);
  try
    P.Executable := FApplicationConfig.SoxPath;
    P.Parameters.Add(FileName);
    P.Parameters.Add('-r'); P.Parameters.Add('12000');   // 12 kHz
    P.Parameters.Add('-c'); P.Parameters.Add('1');       // Mono (1 channel)
    P.Parameters.Add('-b'); P.Parameters.Add('16');      // 16 bit
    P.Parameters.Add(SOX_OUT_FILE);
    P.Parameters.Add('pitch'); P.Parameters.Add(inttostr(TrackBarVoicePitch.Position*100));  // adjust pitch
    P.Options := [poWaitOnExit, poNoConsole];
    P.Execute;
  finally
    P.Free;
  end;

  // --- read wav file into a filestream ---
  fs := TFileStream.Create(SOX_OUT_FILE, fmOpenRead or fmShareDenyWrite);
  try
    fs.Read(riff, 4);
    fs.Read(chunkSize, 4);
    fs.Read(riff, 4); // "WAVE"

    fs.Read(chunkID, 4);
    fs.Read(fmtSize, 4);
    fs.Read(audioFormat, 2);
    fs.Read(numChannels, 2);
    fs.Read(sampleRate, 4);
    fs.Seek(6, soCurrent);
    fs.Read(bitsPerSample, 2);

    if (audioFormat <> 1) or (numChannels <> 1) or (bitsPerSample <> 16) then begin
      LogEvent(lvERROR, 'Cannot play WAV - only 16bit sound is supported.', now);
      Exit;
    end;

    fs.Read(chunkID, 4); // "data"
    fs.Read(dataSize, 4);

    numSamples := dataSize div 2;
    SetLength(RawData, numSamples);

    for i := 0 to numSamples - 1 do
      fs.Read(RawData[i], 2);

  finally
    fs.Free;
  end;

  // --- resample with desired loudness ---
  volumeFactor := TrackBarVoiceVolume.Position / 100.0;
  outLen := Length(RawData);
  SetLength(PCM12k, outLen);
  for i := 0 to outLen - 1 do
  begin
    PCM12k[i] := Round(RawData[i]  * volumeFactor);
  end;

  // instruct parrot to move its beak for the duration:
  durationMS := Round(outLen / 12); // calculate the sound duration in ms by dividing through 12kHz
  rc := libusb_control_transfer(pUsbDevice, $40, $01, MASCOT_CMD_CTRL + mascotLedState + MASCOT_CMD_BEAK,
    durationMS div 4, nil, 0, durationMS div 4);
  // for debugging: LogEvent(lvINFO, 'Motor cmd rc=' + Inttostr(rc), now);

  // --- send sound to robot ---
  i := 0;
  while i < Length(PCM12k)*2 do
  begin
    if i + MaxBlockSize > Length(PCM12k)*2 then
      MaxBlockSize := Length(PCM12k)*2 - i;

    Move(Pointer(@PCM12k[i div 2])^, block, MaxBlockSize);
    rc := libusb_interrupt_transfer(pUsbDevice, $02, @block[0], MaxBlockSize, transferred, 5000);
    if rc <> 0 then begin
      LogEvent(lvERROR, 'Cannot play WAV - failure during transfer via USB.', now);
      Exit;
    end;
    Inc(i, MaxBlockSize);
  end;

  // for debugging: LogEvent(lvINFO, 'WAV file successfully sent.', now);

end;

// user wants to save the log file
procedure TForm1.btnSaveLogFileClick(Sender: TObject);
begin
  if SaveDialog1.Execute then
    SaveLogFile(SaveDialog1.FileName);
end;

// user wants to specify, where SoX is located
procedure TForm1.btnOpenSoXPathClick(Sender: TObject);
begin
  //OpenDialogSox.FileName:= FApplicationConfig.SoxPath;
  if OpenDialogSox.Execute then begin
    FApplicationConfig.SoxPath:= OpenDialogSox.FileName;
    EditSoxPath.Text:= FApplicationConfig.SoxPath;
  end;
end;

// user wants to sort the list of random parrot sayings
procedure TForm1.btnSortClick(Sender: TObject);
var
  TempList: TStringList;
begin
  TempList := TStringList.Create;
  try
    TempList.Assign(SynEditTalk.Lines);
    TempList.Sort;
    SynEditTalk.Lines.Assign(TempList);
    btnSaveTalkLines.Enabled:= True;
  finally
    TempList.Free;
  end;
end;

procedure TForm1.SynEditNewSongChange(Sender: TObject);
begin
  btnSaveNextSongLines.Enabled:= True;
end;

// user wants to preview a parrot saying
procedure TForm1.SynEditNewSongKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key = 122) and (not IsSpeaking) then // 122 = F11
    btnPreviewNextSongClick(Self);
end;

procedure TForm1.SynEditNoWinampChange(Sender: TObject);
begin
  btnSaveNoWinampLines.Enabled:= True;
end;

// user wants to preview a parrot saying
procedure TForm1.SynEditNoWinampKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key = 122) and (not IsSpeaking) then // 122 = F11
    btnPreviewNoWinampClick(Self);
end;

procedure TForm1.SynEditTalkChange(Sender: TObject);
begin
   btnSaveTalkLines.Enabled:= True;
end;

// user wants to preview a parrot saying
procedure TForm1.SynEditTalkKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key = 122) and (not IsSpeaking) then  // 122 = F11
    btnTextPreviewClick(Self);
end;


// event method of TimerEventCheck, checks if Winamp has changed its title and
// lets the parrot say something if so
procedure TForm1.TimerEventCheckTimer(Sender: TObject);
var
  s: String;
  i: Integer;
begin

  if (isParrotMuted) then // skip if parrot is set to mute
    Exit;

  // check for events

  // Winamp title changed
  if (FApplicationConfig.DoCommentSongChange) and (WinampControl.IsWinampRunning) and (WinampControl.IsPlaying) then begin
    s:= WinampControl.GetTrackTitle;
    if (FLastWinampTitle = '') then
      FLastWinampTitle := s;
    if (FLastWinampTitle <> '') and (s <> FLastWinampTitle) then begin
      FLastWinampTitle:= s;
      i:= Random(SynEditNewSong.Lines.Count);
      S:= UTF8Decode(SynEditNewSong.Lines[i]);
      ParrotSpeakNonBlocking(S);
    end;
  end;


end;


// Event method of TimerRandomTalk, picks a line from SynEditTalk and lets
// the parrot say it.
// Note: it is not checked here whether the line has a condition which might not
// be fulfilled. - If the condition would not apply, the parrot won't say
// anything and the timer interval has to elapse again.
procedure TForm1.TimerRandomTalkTimer(Sender: TObject);
var
  S: string;
  I: Integer;
begin
  if (isParrotMuted) then
    Exit;

  i:= Random(SynEditTalk.Lines.Count);

  S:= UTF8Decode(SynEditTalk.Lines[i]);
  ParrotSpeakNonBlocking(S);
end;

// user changes the timer interval
procedure TForm1.TrackBarTalkTimerChange(Sender: TObject);
var
  timeStr: String;
  i: Integer;
  s, m: Integer;
begin
  FApplicationConfig.RandomTalkInterval:= TrackBarTalkTimer.Position;
  i:= FApplicationConfig.RandomTalkInterval * 10;  // i now interval in seconds
  m:= i div 60;
  s:= i mod 60;
  timeStr:= Format('%.2d:%.2d', [m, s]);
  lblTalkTimer.caption:= timeStr;
  TimerRandomTalk.Enabled:= False;
  TimerRandomTalk.Interval:= i * 1000;
  TimerRandomTalk.Enabled:= FApplicationConfig.DoRandomTalk;
end;


// user changes the TTS voice speed
procedure TForm1.TrackBarVoiceSpeedChange(Sender: TObject);
begin
  FApplicationConfig.VoiceSpeed:=TrackBarVoiceSpeed.Position;
  lblVoiceSpeed.Caption := inttostr(FApplicationConfig.VoiceSpeed);
end;


// user changes the TTS voice pitch
procedure TForm1.TrackBarVoicePitchChange(Sender: TObject);
begin
  FApplicationConfig.VoicePitch:= TrackBarVoicePitch.Position;
  lblPitch.caption:= inttostr(FApplicationConfig.VoicePitch);
end;


// user changes the speaker volume
procedure TForm1.TrackBarVoiceVolumeChange(Sender: TObject);
begin
  FApplicationConfig.VoiceLoudness:= TrackBarVoiceVolume.Position;
  lblLoudness.caption:= inttostr(FApplicationConfig.VoiceLoudness)+'%';
end;


// user wants to test audio on the parrot speaker
procedure TForm1.btnPlayWaveClick(Sender: TObject);
var
  S: String;
begin
  if OpenDialog1.Execute then begin
    S:= '[WAV='+OpenDialog1.FileName+']';
    ParrotSpeakNonBlocking(S);
  end;
end;

procedure TForm1.btnPreviewNextSongClick(Sender: TObject);
var
  S: string;
begin
  S:= UTF8Decode(SynEditNewSong.LineText);

  ParrotSpeakNonBlocking(S);
end;

procedure TForm1.btnPreviewNoWinampClick(Sender: TObject);
var
  S: string;
begin
  S:= UTF8Decode(SynEditNoWinamp.LineText);

  ParrotSpeakNonBlocking(S);
end;

procedure TForm1.btnRedClick(Sender: TObject);
var
  rc: Integer;
begin
  if (not isUsbConnected) then
    Exit;

  mascotLedState:= mascotLedState XOR MASCOT_FLAG_LEDRED;

  rc := libusb_control_transfer(pUsbDevice, $40, $01, MASCOT_CMD_CTRL + mascotLedState,
    0, nil, 0, 0);

  // for debugging: LogEvent(lvINFO, 'LED cmd rc=' + IntToStr(rc), now);
end;

procedure TForm1.btnSaveNextSongLinesClick(Sender: TObject);
begin
  SynEditNewSong.Lines.SaveToFile(NEWSONG_FILE, TEncoding.UTF8);
  btnSaveNextSongLines.Enabled:= False;
end;

procedure TForm1.btnSaveNoWinampLinesClick(Sender: TObject);
begin
  SynEditNoWinamp.Lines.SaveToFile(NOWINAMP_FILE, TEncoding.UTF8);
  btnSaveNoWinampLines.Enabled:= False;
end;



procedure TForm1.btnTestBeakClick(Sender: TObject);
begin
  ParrotDoAction(maBEAK, 1000);   // move beak for 1s
end;

procedure TForm1.btnTestTiltClick(Sender: TObject);
begin
  ParrotDoAction(maTILT, 300); // tilt robot head
end;


// Lets the parrot speak a text.
// This function will preprocess the string to check if any conditions or keywords apply (PreProcessTTS).
// Remaining text to be spoken is passed to TTS and stored in a file (TTS_WAV_FILE).
// The TTS_WAV_FILE is then transferred to the parrot speaker via USB.
function TForm1.ParrotSpeak(s: String): Boolean;
var
  TempFile: string;
begin

  // abort if text to speak is empty or an comment
  if ((s.IsEmpty) or (s.StartsWith(';'))) then begin
    Result:= False;
    Exit;
  end;

  // abort if busy
  if (IsSpeaking) then begin
    Result:= False;
    Exit;
  end;

  IsSpeaking:= True;
  HandleSpeakingChange; // make sure, all UI elements, related to parrot action are disabled

  S:= PreProcessTTS(S).Trim;  // pre-process the text (this may let the parrot play sounds and/or do movement)

  if (s.IsEmpty) then begin // nothing more to say
    IsSpeaking:= False;
    Result:= True;
    Exit;
  end;

  TempFile := TTS_WAV_FILE;
  TextToWav(s, TempFile);  // TTS to a wave file

  PlayWavOnParrot(TempFile);

  Result:= True;
end;

procedure TForm1.btnTextPreviewClick(Sender: TObject);
var
  S: string;
  ParrotThread: TParrotThread;
begin
  S:= UTF8Decode(SynEditTalk.LineText);

  ParrotSpeakNonBlocking(S);
end;


procedure TForm1.ParrotSpeakNonBlocking(S: String);
var
  ParrotThread: TParrotThread;
begin
  if (isParrotSpeaking) then
    Exit;

  ParrotThread:= TParrotThread.Create(S);
  ParrotThread.OnTerminate := @ParrotThreadFinished;
  ParrotThread.Start;
end;


// user wants to clear the event log
procedure TForm1.btnClearLogClick(Sender: TObject);
begin
  LoggingListView.Clear;
end;

procedure TForm1.SplashImageClick(Sender: TObject);
begin
  SplashImage.Visible:= False;
  PageControl1.Visible:= True;
end;

// user switches GUI language to English
procedure TForm1.btnEnglishClick(Sender: TObject);
begin
  FApplicationConfig.Language:= 'en';
  SetDefaultLang(FApplicationConfig.Language);
  VoiceComboBoxChange(Self);
  if FApplicationConfig.Language = 'de' then
    InfoMemo.lines.LoadFromFile('info_de.txt', TEncoding.UTF8)
  else
    InfoMemo.lines.LoadFromFile('info_en.txt', TEncoding.UTF8);
end;

// user switches GUI language to German
procedure TForm1.btnGermanClick(Sender: TObject);
begin
  FApplicationConfig.Language:= 'de';
  SetDefaultLang(FApplicationConfig.Language);
  VoiceComboBoxChange(Self);
  if FApplicationConfig.Language = 'de' then
    InfoMemo.lines.LoadFromFile('info_de.txt', TEncoding.UTF8)
  else
    InfoMemo.lines.LoadFromFile('info_en.txt', TEncoding.UTF8);
end;

// user wants to hear a TTS voice preview with custom text from EditTTS
procedure TForm1.btnTTSClick(Sender: TObject);
begin
  if EditTTS.Text = '' then
    Exit;

  ParrotSpeakNonBlocking(EditTTS.Text);
end;

// user checks/unchecks the checkbox for the balloon hint feature
procedure TForm1.cbBalloonHintsChange(Sender: TObject);
begin
  FApplicationConfig.DoShowTaskbarPopups:= cbBalloonHints.Checked;
end;

// user checks/unchecks the checkbox for the "parrot comments song change" feature
procedure TForm1.cbCommentSongChangeChange(Sender: TObject);
begin
  FApplicationConfig.DoCommentSongChange:= cbCommentSongChange.Checked;
end;

// user checks/unchecks the checkbox for the random parrot sayings feature
procedure TForm1.cbTalkEnabledChange(Sender: TObject);
begin
  FApplicationConfig.DoRandomTalk:= cbTalkEnabled.Checked;
  TimerRandomTalk.Enabled:= FApplicationConfig.DoRandomTalk;
end;

procedure TForm1.HideTimerTimer(Sender: TObject);
begin
  SplashImage.Visible:= False;
  PageControl1.Visible:= True;
  if (FileExists(FIniFileName)) then begin
    Hide;
  end;
  HideTimer.Enabled:= False;
end;

// Handler method to be called when the parrot thread finishes
procedure TForm1.ParrotThreadFinished(Sender: TObject);
begin
  IsSpeaking:= False;
  Caption:= RsAppTitle + ' v' + ResourceVersionInfo;
  if FApplicationConfig.DoShowTaskbarPopups then begin // clear balloon hint
    TrayIcon1.BalloonHint:= '';
    TrayIcon1.BalloonTimeout:= 1;
    TrayIcon1.ShowBalloonHint;
  end;
  if isHeadTilted then
    ParrotDoAction(maFLAP, 200); // un-tilt the head
end;

// user selects another TTS voice
procedure TForm1.VoiceComboBoxChange(Sender: TObject);
var
  Voices, Voice: Variant;
  i: Integer;
  vName: string;
begin

  GroupBoxVoiceInfo.Caption := '';
  lblGender.Caption:= RsGender;
  lblAge.Caption:= RsAge;
  lblLanguage.Caption:= RsLanguage;

  CoInitialize(nil);
  try
    Voice := CreateOleObject('SAPI.SpVoice');

    // get a list of all voices
    Voices := Voice.GetVoices;

    for i:= 0 to Voices.Count - 1 do begin
      Voice.Voice := Voices.Item(i);

      // get voice attributes
      vName     := Voice.Voice.GetAttribute('Name');
      if (vName = VoiceComboBox.Text) then begin
        GroupBoxVoiceInfo.Caption:= vName;
        lblGender.Caption:= lblGender.Caption + ' ' + VarToStr(Voice.Voice.GetAttribute('Gender'));   // Male/Female/Neutral
        lblAge.Caption:= lblAge.Caption + ' ' + VarToStr(Voice.Voice.GetAttribute('Age'));      // Child/Teen/Adult/Senior
        lblLanguage.Caption:= lblLanguage.Caption + ' ' + VarToStr(Voice.Voice.GetAttribute('Language'));
        Break;
      end;

    end;

    // store selected voice in settings
    FApplicationConfig.VoiceName:= vName;

  finally
    CoUninitialize;
  end;
end;


// convert a string to speech using SAPI TTS and store the resulting audio in a wav file
procedure TForm1.TextToWav(const aText, FileName: string);
var
  SpVoice, SpFileStream, Token: OleVariant;
  S: string;
begin
  S:= aText;

  if S.IsEmpty then // nothing to do, when just empty
    Exit;

  CoInitialize(nil); // init COM
  try
    SpVoice := CreateOleObject('SAPI.SpVoice');

    // select confuígured voice
    if VoiceComboBox.ItemIndex >= 0 then
    begin
      Token := SpVoice.GetVoices.Item(VoiceComboBox.ItemIndex);
      SpVoice.Voice := Token;
    end;

    // create FileStream
    SpFileStream := CreateOleObject('SAPI.SpFileStream');
    SpFileStream.Open(FileName, 3, False); // SSFMCreateForWrite = 3

    // redirect voice to stream
    SpVoice.AudioOutputStream := SpFileStream;

    // set talking speed
    SpVoice.Rate := TrackBarVoiceSpeed.Position;

    // convert text-to-speech into file
    SpVoice.Speak(S, 8);

    // close stream
    SpFileStream.Close;
  finally
    CoUninitialize;
  end;
end;

// show popup menu when click on tray icon
procedure TForm1.TrayIcon1Click(Sender: TObject);
begin
    PopupMenu1.PopUp;
end;


end.

