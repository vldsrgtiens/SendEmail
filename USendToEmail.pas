unit USendToEmail;

interface
uses
Wininet,


  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
   Vcl.StdCtrls, Vcl.ToolWin, Vcl.ActnMan, Vcl.ActnCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Buttons, Vcl.TabNotBk, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdComponent,
  IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient,
  IdSMTPBase, IdSMTP, IdBaseComponent, IdAttachment, IdMessageBuilder,
  FileCtrl,jpeg,
  System.IOUtils,
  DATEUTILS,
  IpHlpApi, IpTypes,
  ClipBrd,

  system.NetEncoding,
  IdAttachmentfile,
  System.Threading,
   System.StrUtils, System.RegularExpressions,
  IdMessage, Vcl.OleCtrls, Vcl.Samples.Spin, Vcl.Menus, Vcl.AppEvnts,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Grids, Vcl.DBGrids, IdAntiFreezeBase, IdAntiFreeze,
  Vcl.Imaging.pngimage, Vcl.ExtDlgs;

type
 TTypeMessge = (tInfo,tError,tSystem);

    TCard = class
    private
      Email : string;
      PhoneNumber: string;
      EmailLastSendDate  : TDateTime;
      PhoneNumberLastSendDate  : TDateTime;

    public
      // Коструктрор
      constructor Create(FEmail : string;
                         FPhoneNumber: string);

  end;

  TMyFindingThread = class(TThread)
  private
    FName: string;
    FEmail: string;
    FPhone: string;
    procedure Progress;
    procedure PrintLog(str: string);
    procedure FindProg;
    procedure DOC_Files;
    procedure PDF_Files;
    procedure EndThread;
  public
    procedure Execute; override;
  end;

  TSaveInDBaseThread = class(TThread)
  private
    total: integer;
    count: integer;
    procedure Progress;
    procedure SaveInDBase;
  public
    procedure Execute; override;
  end;

  TForm9 = class(TForm)
    Panel2: TPanel;
    SSLOpen: TIdSSLIOHandlerSocketOpenSSL;
    SMTP: TIdSMTP;
    TrayIcon1: TTrayIcon;
    PopupMenu1: TPopupMenu;
    BalloonHint1: TBalloonHint;
    ApplicationEvents1: TApplicationEvents;
    TimerDelay: TTimer;
    StatusBar1: TStatusBar;
    DataSource1: TDataSource;
    TimerSendEmail: TTimer;
    TimerExecSending: TTimer;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    TabSheet5: TTabSheet;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    MyEMailAddress: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label9: TLabel;
    Panel3: TPanel;
    MyGav: TLabel;
    MyPost: TComboBox;
    MyLogin: TEdit;
    MyPass: TEdit;
    NameSender: TEdit;
    SubjectMail: TEdit;
    BitBtn1: TBitBtn;
    GroupBox2: TGroupBox;
    MyTextMessage: TMemo;
    BitBtn5: TBitBtn;
    Settings: TTabSheet;
    Panel7: TPanel;
    Label1: TLabel;
    DateEndSending: TDateTimePicker;
    Panel8: TPanel;
    Label3: TLabel;
    Panel12: TPanel;
    CountEmailSendInDay: TSpinEdit;
    Panel9: TPanel;
    Label5: TLabel;
    Panel13: TPanel;
    MinDaysToSend: TSpinEdit;
    Panel10: TPanel;
    Label6: TLabel;
    Edit1: TEdit;
    Panel14: TPanel;
    TabSheet1: TTabSheet;
    GroupBox3: TGroupBox;
    Panel4: TPanel;
    LabelСhosenDirectory: TLabel;
    BitBtn2: TBitBtn;
    GroupBox4: TGroupBox;
    Panel5: TPanel;
    BitBtnScanFiles: TBitBtn;
    MemoEmailFromFiles: TMemo;
    PanelProgress: TPanel;
    LabelProgressFile: TLabel;
    LabelProgressTotal: TLabel;
    ProgressBarFile: TProgressBar;
    ProgressBarTotal: TProgressBar;
    Panel15: TPanel;
    CheckBoxDOC: TCheckBox;
    CheckBoxRTF: TCheckBox;
    TabSheet2: TTabSheet;
    BitBtn3: TBitBtn;
    RichLog: TRichEdit;
    TabSheet3: TTabSheet;
    MemoPDF: TMemo;
    TabSheet4: TTabSheet;
    BitBtn4: TBitBtn;
    BitBtn7: TBitBtn;
    Edit2: TEdit;
    GroupBox5: TGroupBox;
    LabelInfoSending: TLabel;
    BitBtnSendEmail: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn10: TBitBtn;
    IdAntiFreeze1: TIdAntiFreeze;
    Panel16: TPanel;
    ImageVisible: TImage;
    ImageUnVisible: TImage;
    BitBtn11: TBitBtn;
    BitBtn9: TBitBtn;
    SendPhones: TTabSheet;
    Panel17: TPanel;
    Edit3: TEdit;
    Label8: TLabel;
    BitBtn12: TBitBtn;
    Panel18: TPanel;
    Label10: TLabel;
    ProgressBar1: TProgressBar;
    Memo1: TMemo;
    Button1: TButton;
    OpenDialog2: TOpenDialog;
    CheckBoxTXT: TCheckBox;
    PageControlTextLetter: TPageControl;
    TextSheet: TTabSheet;
    ImageSheet: TTabSheet;
    ImageSend: TImage;
    OpenPictureDialog1: TOpenPictureDialog;
    DBGrid1: TDBGrid;
    SpeedButton1: TSpeedButton;
    Memo2: TMemo;
    Label11: TLabel;
    Label12: TLabel;
    NameSenderImage: TEdit;
    Label13: TLabel;
    SubjectMailImage: TEdit;
    Label14: TLabel;
    MessageTextImage: TEdit;
    ImageDefault: TImage;
    ImagePath: TLabel;
    GroupBox7: TGroupBox;
    BitBtn13: TBitBtn;
    Panel6: TPanel;
    TestPeriod: TTabSheet;
    Panel11: TPanel;
    Label7: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Panel21: TPanel;
    BitBtn14: TBitBtn;
    Panel22: TPanel;
    BitBtn15: TBitBtn;
    Panel19: TPanel;
    Memo3: TMemo;
    procedure MyLoginClick(Sender: TObject);
    procedure MyLoginExit(Sender: TObject);
    procedure MyPostChange(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtnScanFilesClick(Sender: TObject);
    procedure MemoEmailFromFilesChange(Sender: TObject);
    procedure BitBtnSendEmailClick(Sender: TObject);
    procedure ApplicationEvents1Minimize(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure StatusBar1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
      const Rect: TRect);
    procedure Edit1Change(Sender: TObject);
    procedure CheckBoxAutoRunClick(Sender: TObject);
    procedure CountEmailSendInDayChange(Sender: TObject);
    procedure DateEndSendingChange(Sender: TObject);
    procedure MinDaysToSendChange(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure TimerSendEmailTimer(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure TimerExecSendingTimer(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure MyLoginChange(Sender: TObject);
    procedure MyPassChange(Sender: TObject);
    procedure NameSenderChange(Sender: TObject);
    procedure SubjectMailChange(Sender: TObject);
    procedure MyTextMessageChange(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure ImageVisibleClick(Sender: TObject);
    procedure ImageUnVisibleClick(Sender: TObject);
    procedure BitBtn12Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ImageSendClick(Sender: TObject);
    procedure TrayIcon1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure MessageTextImageChange(Sender: TObject);
    procedure BitBtn13Click(Sender: TObject);
    procedure BitBtn15Click(Sender: TObject);
    procedure BitBtn14Click(Sender: TObject);

  private
    { Private declarations }
    function IsValidEmail(const Value: string): boolean;
    function CheckAllowed(const s: string): boolean;
    function CheckSend():boolean;
    procedure ErrorMessage(str: string);
    function SendMail(RecipientEmail,  MyMessage: string;SendFile: string): integer;
    function CreateMessageText(): string;
    procedure ScanFiles;
    Procedure SimuClavierCtrl_s(Const s : String);
    Procedure SimuClavierCtrl_(var C : Char);
    function ConnectAccount:boolean;
    procedure MakeTextFile;
    procedure Delay(msec:integer);
    function AddToBase:boolean;
    function ConnectBase: boolean;
    function DisconnectBase: boolean;

    function FindEmailInBase(FEmail: string): boolean;
    function AppendCardToBase(FEmail,FPhone: string):boolean;
    function GetEmailFromBase:string;
    function ChangeEmailStatus(CHEmail,NStatus: string):boolean;

    procedure Autorun(Flag:boolean; Path:String);
    procedure ExecSending;
    procedure RepeatSend;
    procedure SetExecTimer;
    function GetSysInfo(Mac:TStringList):TStringList;
    procedure Log(lType:TTypeMessge;lStr:string);
    procedure SaveCFG;
    procedure LoadConfig;
    procedure SaveAcnt;
    procedure LoadAcnt;
    procedure SaveMac;
    function CheckReg:boolean;
    function CheckInternet:boolean;
    procedure AbortAutoSending;
    procedure StartAutoSending;
    procedure ShowTrayBalloon(tm:TBalloonFlags;title,msg: string);
    procedure SaveInDataBase;
    function InternetConnected: boolean;
  public
    { Public declarations }
    procedure HardwareInfo;

  end;

var
  Form9: TForm9;
  FlagSender,FlagRecepiant,FlagMessage: boolean;
  SMTP_Host: string;
  SMTP_Port: Word;
  SMTP_AuthType: IdSMTP.TIdSMTPAuthenticationType;

  FindingThread: TMyFindingThread;
  SaveInDBaseThread: TSaveInDBaseThread;
  chosenDirectory: string;
  on_ : wordbool = false;

  PathExe,PathImages,PathAcnt: string;


  FlagSendImage:boolean=false;
  SendImageFullPath: string;

    CountFiles: integer;
    CountEmail: integer;
    ValueFiles: integer;
    ValueEmail: integer;
    ValuePhone: integer;
    ValueProgressScanFile: integer;


  ScanFilesList: TList;

  AccountConnect: boolean=false;
  TodaySendingEmail: integer;
  Registered: boolean;
  FlagChangeDataSender: boolean;
  SpamCountSend: integer;
  BeforEmail: string;
implementation

{$R *.dfm}

uses DataModule, UDm,Registry;

{ TForm9 }
function TForm9.InternetConnected: boolean;
 var
  lpdwConnectionTypes: word;
 begin
   lpdwConnectionTypes:= INTERNET_CONNECTION_MODEM +
                         INTERNET_CONNECTION_LAN +
                         INTERNET_CONNECTION_PROXY;
   Result:= InternetGetConnectedState(@lpdwConnectionTypes,0);
   if result then

   Log(tError,'InternetConnected: true')
   else
   Log(tError,'InternetConnected: false')
 end;


procedure TForm9.Autorun(Flag:boolean;  Path:String);
var Reg:TRegistry;
begin
  if Flag then
  begin
     Reg := TRegistry.Create;
     Reg.RootKey := HKEY_CURRENT_USER;
     Reg.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Run', false);
     Reg.WriteString('SendEmailProg', Path);
     Reg.Free;
  end
  else
  begin
     Reg := TRegistry.Create;
     Reg.RootKey := HKEY_CURRENT_USER;
     Reg.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',false);
     Reg.DeleteValue('SendEmailProg');
     Reg.Free;
  end;
end;

procedure TForm9.MakeTextFile;
var
  sl: TStringList;
  SCard: TCard;
begin
  sl := TStringList.Create;
  try
    for SCard in ScanFilesList do
     begin
      sl.Add('E-mail:'+SCard.Email+' Тел.:'+SCard.PhoneNumber);
     end;

    sl.SaveToFile(PathExe+'\Emails.txt');
    ErrorMessage('EmailFile:'+PathExe+'\Emails.txt');
  finally
    sl.Free
  end;
end;

procedure TForm9.AbortAutoSending;
begin
  BitBtn5.Caption:='Запустить авторассылку';
  MyTextMessage.Enabled:=true;
  MyLogin.Enabled:=true;
  MyPost.Enabled:=true;
  MyPass.Enabled:=true;
  NameSender.Enabled:=true;
  SubjectMail.Enabled:=true;
  BitBtn1.Enabled:=true;

  TimerExecSending.Enabled:=false;
  TimerSendEmail.Enabled:=false;
end;

procedure TForm9.StartAutoSending;
var
 buttonSelected:integer;
begin
   BitBtn5.Caption:='Остановить авторассылку';
   MyTextMessage.Enabled:=false;
   MyLogin.Enabled:=false;
   MyPost.Enabled:=false;
   MyPass.Enabled:=false;
   NameSender.Enabled:=false;
   SubjectMail.Enabled:=false;
   BitBtn1.Enabled:=false;

   buttonSelected := MessageDlg('Запустить рассылку сейчас?',mtCustom,
                              [mbYes,mbNo], 0);

  // Показ типа выбранной кнопки
  if buttonSelected = mrYes    then ExecSending;
  if buttonSelected = mrNo    then SetExecTimer;
end;

function TForm9.AddToBase: boolean;
begin

end;

function TForm9.FindEmailInBase(FEmail: string): boolean;
begin
 DM.FDQuery1.Active:=true;
 Result:= DM.FDQuery1.Locate('EMail',FEmail,[]);
 DM.FDQuery1.Active:=false;
end;

function TForm9.GetEmailFromBase: string;
var
 FDate: TDate;
 FlagNext: boolean;
 FStatus:string;
begin
 result:='empty';
 FlagNext:=false;
  DM.FDQuery1.IndexFieldNames:='LastDateEmail';
  DM.FDQuery1.Active:=true;
  DM.FDQuery1.First;
 repeat
  FDate:=DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime;
  FStatus:=DM.FDQuery1.FieldByName('StatusEmail').AsString;
  FDate:=IncDay(FDate, StrToInt(MinDaysToSend.Text));
  if ((FDate<=Now) and (FStatus<>'Error')) then
   begin
    result:=DM.FDQuery1.FieldByName('Email').AsString;
    FlagNext:=false;
   end
  else
   begin
    if not DM.FDQuery1.Eof then
     begin
      FlagNext:=true;
      DM.FDQuery1.Next;
     end
    else
     FlagNext:=false;
   end;
 until (FlagNext=false);
end;

function TForm9.GetSysInfo(Mac:TStringList): TStringList;
var
  pAdapterInfo, pTempAdapterInfo: PIP_ADAPTER_INFO;
  AdapterInfo: IP_ADAPTER_INFO;
  BufLen: DWORD;
  Status: DWORD;
  strMAC,s1,s2: String;
  i: Integer;
  Base64: TBase64Encoding;


begin

  Log(tError,'into GetSysInfo begin');
  Base64 := TBase64Encoding.Create(0);
  Mac.Clear;

  BufLen:= sizeof(AdapterInfo);
  pAdapterInfo:= @AdapterInfo;

  Status:= GetAdaptersInfo(nil, BufLen);
  pAdapterInfo:= AllocMem(BufLen);
  try
    Status:= GetAdaptersInfo(pAdapterInfo, BufLen);
    if Status<>0 then exit;


    while (pAdapterInfo <> nil) do
      begin
        //Log(tInfo,'Description: ' + pAdapterInfo^.Description);
        //Log(tInfo,'Name: ' + pAdapterInfo^.AdapterName);

        strMAC := '';
        for I := 0 to pAdapterInfo^.AddressLength - 1 do
            strMAC := strMAC + '-' + IntToHex(pAdapterInfo^.Address[I], 2);

        Delete(strMAC, 1, 1);
        //Log(tInfo,'MAC address: ' + strMAC);
        //Log(tInfo,'IP address: ' + pAdapterInfo^.IpAddressList.IpAddress.S);
        //Log(tInfo,'IP subnet mask: ' + pAdapterInfo^.IpAddressList.IpMask.S);
        //Log(tInfo,'Gateway: ' + pAdapterInfo^.GatewayList.IpAddress.S);
        //Log(tInfo,'DHCP enabled: ' + IntTOStr(pAdapterInfo^.DhcpEnabled));
        //Log(tInfo,'DHCP: ' + pAdapterInfo^.DhcpServer.IpAddress.S);
        //Log(tInfo,'Have WINS: ' + BoolToStr(pAdapterInfo^.HaveWins,True));
        //Log(tInfo,'Primary WINS: ' + pAdapterInfo^.PrimaryWinsServer.IpAddress.S);
        //Log(tInfo,'Secondary WINS: ' + pAdapterInfo^.SecondaryWinsServer.IpAddress.S);

        Mac.Add(strMAC);

        s1:=strMAC;
        s2:=Base64.Encode(s1);
        //Log(tInfo,'-----------');
        //Log(tInfo,s1);
        //Log(tInfo,s2);
        //Log(tInfo,Base64.Decode(s2));

        pTempAdapterInfo := pAdapterInfo;
        pAdapterInfo:= pAdapterInfo^.Next;

      if assigned(pAdapterInfo) then Dispose(pTempAdapterInfo);
    end;
    Result:=Mac;
  finally
    Dispose(pAdapterInfo);
    Base64.Free;
  end;

  Log(tError,'into GetSysInfo end');
end;

procedure TForm9.CountEmailSendInDayChange(Sender: TObject);
begin
 BitBtn11.Enabled:=true;
end;

procedure TForm9.MinDaysToSendChange(Sender: TObject);
begin
 BitBtn11.Enabled:=true;
end;



procedure TForm9.StatusBar1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
  const Rect: TRect);
begin
if Panel = StatusBar.Panels[0] then
  begin
   if StatusBar.Panels[0].Text=' Account Disconnect' then
    begin
     StatusBar.Canvas.Font.Color := clRed;
     StatusBar.Canvas.Font.Style := [fsBold];
     StatusBar.Canvas.TextOut(Rect.left, Rect.Top,StatusBar.Panels[0].Text);
    end
   else
   begin
    StatusBar.Canvas.Font.Color := clGreen;
    StatusBar.Canvas.Font.Style := [fsBold];
    StatusBar.Canvas.TextOut(Rect.left, Rect.Top,StatusBar.Panels[0].Text);
   end;
  end;

if Panel = StatusBar.Panels[2] then
  begin
   if StatusBar.Panels[2].Text='Unregistered' then
    begin
     StatusBar.Canvas.Font.Color := clRed;
     StatusBar.Canvas.Font.Style := [fsBold];
     StatusBar.Canvas.TextOut(Rect.left, Rect.Top,StatusBar.Panels[2].Text);
    end
   else
   begin
    StatusBar.Canvas.Font.Color := clGreen;
    StatusBar.Canvas.Font.Style := [fsBold];
    StatusBar.Canvas.TextOut(Rect.left, Rect.Top,StatusBar.Panels[2].Text);
   end;
  end;
end;

procedure TForm9.SubjectMailChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

function TForm9.AppendCardToBase(FEmail, FPhone: string): boolean;
begin
 DM.FDQuery1.Active:=true;
 DM.FDQuery1.Append;
 DM.FDQuery1.FieldByName('EMail').AsString := FEmail;
 DM.FDQuery1.FieldByName('Phone').AsString := FPhone;
 DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime := EncodeDate(2000, 1, 1);
 DM.FDQuery1.FieldByName('LastDatePhone').AsDateTime := EncodeDate(2000, 1, 1);
 DM.FDQuery1.FieldByName('StatusEmail').AsString := 'unknown';
 DM.FDQuery1.FieldByName('StatusPhone').AsString := 'unknown';
 DM.FDQuery1.Post;
end;

procedure TForm9.ApplicationEvents1Minimize(Sender: TObject);
begin
TrayIcon1.visible:=true;
//Убираем с панели задач
 ShowWindow(Handle,SW_HIDE);  // Скрываем программу
 ShowWindow(Application.Handle,SW_HIDE);  // Скрываем кнопку с TaskBar'а
 SetWindowLong(Application.Handle, GWL_EXSTYLE,
 GetWindowLong(Application.Handle, GWL_EXSTYLE) or (not WS_EX_APPWINDOW));
end;

procedure TForm9.BitBtn10Click(Sender: TObject);
begin
 Log(tInfo,'----');
 DM.FDQuery1.Active:=false;
 DM.FDQuery1.SQL.Clear;
 DM.FDQuery1.SQL.Text:='SELECT * FROM contact';
 DM.FDQuery1.Open;
 DM.FDQuery1.Last;
 Log(tInfo,'База данных: количество записей : ' + IntToStr(DM.FDQuery1.RecordCount));

 DM.FDQuery1.SQL.Clear;
 DM.FDQuery1.SQL.Text:='SELECT * FROM contact WHERE LastDateEmail='+ Chr(34)+'2000-01-01'+Chr(34);
 DM.FDQuery1.Open;
 DM.FDQuery1.Last;
 Log(tInfo,'База данных: кол-во Email еще не отправлялось письмо : ' + IntToStr(DM.FDQuery1.RecordCount));

 DM.FDQuery1.SQL.Clear;
 DM.FDQuery1.SQL.Text:='SELECT * FROM contact WHERE LastDateEmail LIKE '+ QuotedStr('%'+FormatDateTime('yyyy-mm-dd', Now) +'%')   ;
 DM.FDQuery1.Open;
 DM.FDQuery1.Last;
 Log(tInfo,'База данных: сегодня отправлено писем : ' + IntToStr(DM.FDQuery1.RecordCount));

 Log(tInfo,'----');
 DM.FDQuery1.Close;
end;

procedure TForm9.BitBtn11Click(Sender: TObject);
begin
 SaveCFG;
end;

procedure TForm9.BitBtn12Click(Sender: TObject);
var
 DBTotal,DBCount:integer;
 DBStringList: TStringList;
 DBPhone,str:string;
 RepFlag: TReplaceFlags;
 Base64: TBase64Encoding;
begin
 if AccountConnect=false then
  begin
   ErrorMessage('Аккоунт не подключен,подключитесь и повторите попытку');
   Exit;
  end;
 if not IsValidEmail(Edit3.Text) then
  begin
   ErrorMessage('Введите корректный Email');
   Exit;
  end;



 DM.FDQuery1.Active:=false;
 DM.FDQuery1.SQL.Clear;
 DM.FDQuery1.SQL.Text:='SELECT * FROM contact';
 DM.FDQuery1.Open;
 DM.FDQuery1.Last;
 DBTotal:=DM.FDQuery1.RecordCount;
 try
  if DBTotal>0 then
   begin
    DBStringList:=TStringList.Create;
    DBCount:=0;
    Panel18.Visible:=true;
    Base64:=TBase64Encoding.Create(0);
    try
     DM.FDQuery1.First;
     while not(DM.FDQuery1.Eof) do
      Begin
       DBPhone:=DM.FDQuery1.FieldByName('Phone').AsString;
       if DBPhone<>'' then
        begin
         DBStringList.Add(Base64.Encode(DBPhone));
         DBCount:=DBCount+1;
         ProgressBar1.Position:=round(DBCount/DBTotal*100);
         Label10.Caption:='формирование базы для отправки: '+IntToStr(DBCount)+'/'+IntToStr(DBTotal);
         Application.ProcessMessages;
        end;
      DM.FDQuery1.Next;
      End;
      str:='';
      RepFlag:=[rfReplaceAll];
      str:=DateTimeToStr(Now);
      str:=StringReplace(str, '.', '',RepFlag);
      str:=StringReplace(str, ':', '',RepFlag);
      str:=StringReplace(str, ' ', '',RepFlag);
      str:='\Phones'+str+'.txt';
      DBStringList.SaveToFile(PathExe+str);
      if SendMail(Edit3.Text,'База Телефонов',str)=0 then
       begin
        ErrorMessage('База отправлена');
        DeleteFile(PathExe+str);
       end
      else
      begin
       ErrorMessage('Возникли проблемы при отправке');
      end;

 //DM.FDQuery1.Active:=true;
   Log(tInfo,'----');
   Log(tInfo,'База данных: количество записей ' + IntToStr(DM.FDQuery1.RecordCount));
   Log(tInfo,'----');
    finally
     DBStringList.Free;
     Panel18.Visible:=false;
     Label10.Caption:='формирование базы для отправки ';
     Base64.Free;
    end;
   end;
 finally
   DM.FDQuery1.Close;
 end;


end;

procedure TForm9.BitBtn13Click(Sender: TObject);
begin
 Log(tInfo,intToStr(  DaysBetween(Date,DateEndSending.Date)));

 if (DateEndSending.Date<=IncDay(Date,3)) then
  begin

   Log(tInfo,'пора продлить');
   PageControl1.ActivePageIndex:=7;
  end
 else
  begin
   Log(tInfo,'тестовый период еще не истек') ;
   ShowMessage('Тестовый период еще не истек.');
  end;
end;

procedure TForm9.BitBtn14Click(Sender: TObject);
begin
 DateEndSending.Date:=IncDay(Now,30);
 //FlagChangeDataSender:=true;
 SaveAcnt;
 ShowMessage('');
end;

procedure TForm9.BitBtn15Click(Sender: TObject);
begin
 Panel19.Visible:= not Panel19.Visible;
end;

procedure TForm9.BitBtn1Click(Sender: TObject);
var
 buttonSelected:integer;
 resultSendMail:integer;
begin
 if BitBtn1.Caption='Отключиться' then
  begin
   BitBtn1.Caption:='Для корректной работы требеутся перезапуск программы';
   SMTP.Disconnect;
   AccountConnect:=false;
   StatusBar1.Panels[0].Text:=' Account Disconnect';
   BitBtn5.Enabled:=false;
  end
 else

 if BitBtn1.Caption='Подключиться' then
 if InternetConnected then
 begin
 if ConnectAccount then
  begin
       begin
    resultSendMail:= SendMail(MyEmailAddress.Caption,CreateMessageText,''); //MyEmailAddress.Caption
     if resultSendMail=0 then
      begin
       AccountConnect:=true;
       StatusBar1.Panels[0].Text:=' Account Connect';
       BitBtn1.Caption:='Отключиться';
       if FlagChangeDataSender then
        begin
         buttonSelected := MessageDlg('Подключение установлено, хотите установить данный аккоунт по умолчанию?',mtConfirmation,
                              [mbYes,mbNo], 0);
         if buttonSelected = mrYes    then SaveAcnt;
        //if buttonSelected = mrNo    then SetExecTimer;
        end
       else
       begin
        ShowTrayBalloon(bfInfo,'Information','Подключение установлено');
       end;

      // resultSendMail:= SendMail('send.programm@mail.ru','Подключение: '+MyEmailAddress.Caption+#13#10+CreateMessageText,'');
      // Log(tInfo,'send.programm@mail.ru: №' + IntToStr(resultSendMail));
       BitBtn5.Enabled:=true;
      end
     else
     begin

      AccountConnect:=false;
      StatusBar1.Panels[0].Text:=' Account Disconnect';
      if SMTP.Connected=true then SMTP.Disconnect;

      if resultSendMail=8 then
        ShowMessage ('Данный аккоунт временно заблокирован на отправку писем (1-2часа) из-за подозрения на СПАМ,'+#13#10+' с подозрением на СПАМ, попробуйте использовать другой аккаунт')
      else
       begin
        if resultSendMail=6 then
         ShowMessage ('Отсутствует подключение к интернету')
        else
         ShowMessage ('Не удалось подключиться, проверьте правильность ввода логина и пароля');
       end;
     end;
    end;


  end;
 end
 else
  ShowMessage ('Отсутствует подключение к интернету :(');
end;

procedure TForm9.BitBtn2Click(Sender: TObject);
var
 countFiles: integer;
begin
countFiles:=0;
MemoEmailFromFiles.Lines.Clear;
if SelectDirectory('Выберите папку с файлами', '',chosenDirectory ) then
 begin
    Log(tInfo,'Выбрана папка '+chosenDirectory) ;
    if CheckBoxDOC.Checked then countFiles:=countFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.doc'));
    if CheckBoxRTF.Checked then countFiles:=countFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.rtf'));
    if CheckBoxTXT.Checked then countFiles:=countFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.txt'));

    if countFiles>0 then
      BitBtnScanFiles.Enabled:=true
    else
      BitBtnScanFiles.Enabled:=false;
    LabelСhosenDirectory.Caption:= chosenDirectory+#13#10+' найдено файлов: '+IntToStr(countFiles)+' шт.';
 end;
end;

procedure TForm9.BitBtn3Click(Sender: TObject);
begin
 //SendMail('vldsrg.tomsk@yandex.ru','AttachFile',true);
end;

procedure TForm9.BitBtn4Click(Sender: TObject);
begin
 ConnectBase;

end;

procedure TForm9.BitBtn5Click(Sender: TObject);
var
 buttonSelected: integer;
begin
 //всякие проверки
 BitBtn5.Tag:=BitBtn5.Tag*-1;
 if BitBtn5.Tag=1 then
  begin
   StartAutoSending;
  end
 else
 begin
  AbortAutoSending;
 end;


end;

procedure TForm9.BitBtn6Click(Sender: TObject);
var
 FName,FEmail,FPhone,buf,str: string;
 f:textfile;
 StartPos,EndPos:integer;
begin
  if OpenDialog1.Execute then
  begin
   FName:=OpenDialog1.FileName;
   AssignFile(f, fName);
   try
    Reset(f);
    while not EOF(f) do
     begin
      readln(f, buf);
            if Pos('@',buf)<>0 then
              begin
               StartPos:=buf.LastIndexOf(' ',Pos('@',buf));
               EndPos:=buf.IndexOf(' ',Pos('@',buf));
               if EndPos<1 then EndPos:=Length(buf);

               str:=copy(buf,StartPos,EndPos-StartPos+1);
               str:=Trim(str);
               if Form9.IsValidEmail(str) then
                begin
                 FEmail:=str;
                end;
              end;

              if Pos('+7 (',buf)<>0 then    //+7 (913) 858-54-52
              begin
               StartPos:=Pos('+7 (',buf);
               str:=copy(buf,StartPos,18);
               Delete(str,16,1);
               Delete(str,13,1);
               Delete(str,8,2);
               Delete(str,3,2);
               if Length(str)=12 then
                begin
                 FPhone:=str;
                end;

              end;



     end;

     BitBtn6.Caption:=FEmail+' : '+FPhone;

   finally
    CloseFile(f);
   end;

  end;




end;

procedure TForm9.BitBtn7Click(Sender: TObject);
begin
 if FindEmailInBase(Edit2.Text) then
  BitBtn7.Caption:='Find TRUE'
 else
  BitBtn7.Caption:='Find False';

end;

procedure TForm9.BitBtn9Click(Sender: TObject);
begin

 DM.FDQuery1.Active:=false;
 DM.FDQuery1.SQL.Clear;
 DM.FDQuery1.SQL.Text:='SELECT * FROM contact';
 DM.FDQuery1.Open;
 DM.FDQuery1.Last;
 //DM.FDQuery1.Active:=true;
 Log(tInfo,'----');
 Log(tInfo,'База данных: количество записей ' + IntToStr(DM.FDQuery1.RecordCount));
 Log(tInfo,'----');
 DM.FDQuery1.Close;


end;

Procedure TForm9.SimuClavierCtrl_s(Const s : String);
var
 c_ : Char;
begin
   c_ := Char(s[1]) ;
   SimuClavierCtrl_(C_);
end;

procedure TForm9.SpeedButton1Click(Sender: TObject);
var
 FDQ: boolean;
begin
  FDQ:=DM.FDQuery1.Active;
  DM.FDQuery1.IndexFieldNames:='LastDateEmail';
  DM.FDQuery1.Active:=not FDQ;
  DM.FDQuery1.First;
end;

procedure TForm9.TimerExecSendingTimer(Sender: TObject);
begin
 TimerExecSending.Enabled:=false;
 ExecSending;
end;

procedure TForm9.TimerSendEmailTimer(Sender: TObject);
begin
 TimerSendEmail.Enabled:=false;
 RepeatSend;
end;

procedure TForm9.TrayIcon1Click(Sender: TObject);
begin
 TrayIcon1.ShowBalloonHint;
 ShowWindow(Handle,SW_RESTORE);
 SetForegroundWindow(Handle);
 TrayIcon1.Visible:=False;
end;

procedure TForm9.TrayIcon1DblClick(Sender: TObject);
begin
 TrayIcon1.ShowBalloonHint;
 ShowWindow(Handle,SW_RESTORE);
 SetForegroundWindow(Handle);
 TrayIcon1.Visible:=False;
end;

Procedure TForm9.SimuClavierCtrl_(var C : Char);
begin
// SIMULATION CALVIER Ctr A  ...    KEYEVENTF_KEYDOWN
keybd_event(VK_LCONTROL,0,0,0); //j'appuie sur la touche CONTROL GAUCHE
//Delay(50);
sleep(50);
//
// baisse et relиve la touche
//
keybd_event(Ord(C),0,0,0); //j'appuie sur la touche  "C"
//Delay(50);
sleep(50);
keybd_event(Ord(C),0,KEYEVENTF_KEYUP,0); //relever touche la touche  "C"
//Delay(50);
sleep(50);

keybd_event(VK_LCONTROL,0,KEYEVENTF_KEYUP,0); //relever  la touche  A
//Delay(50);
sleep(50);
//application.ProcessMessages;
//
end;

procedure TForm9.BitBtnScanFilesClick(Sender: TObject);
var
 Word: variant;
 task: ITask;
 f,t: integer;
begin
 ScanFiles;

  if not ConnectBase then
  begin
   ErrorMessage('Ошибка подключения к базе');
   exit;
  end;

 if ScanFilesList.Count=0 then
  begin
   ErrorMessage('Нечего добавлять в базу');
   exit;
  end;

  ProgressBarFile.Position:=0;
  ProgressBarFile.Max:=ScanFilesList.Count;
  LabelProgressFile.Caption:='';
  ProgressBarTotal.Position:=0;
  ProgressBarTotal.Max:=ScanFilesList.Count;
  LabelProgressTotal.Caption:='';
  ShowTrayBalloon(bfInfo,'Information','Добавление новых Email адресов в базу... Терпение!!');
   //Создаём задачу.

  SaveInDataBase;


   ErrorMessage('Добавлено новых Email  адресов: '+IntToStr(ProgressBarFile.Position)+'шт.');


   ScanFilesList.Clear;
   BitBtnScanFiles.Enabled:=false;
   LabelСhosenDirectory.Caption:='';
   chosenDirectory:='';
   CountFiles:=0;

  Form9.ProgressBarTotal.Position:=0;
  Form9.ProgressBarTotal.Position:=0;
  Form9.PanelProgress.Visible:=false;
end;

procedure TForm9.BitBtnSendEmailClick(Sender: TObject);
var
 SCard: TCard;
 i: integer;
 mess: string;
begin
 if not AccountConnect then
  begin
   ErrorMessage('Введите логин, пароль Вашего аккаунта и подключитесь');
   Exit;
  end;

  if ScanFilesList.Count=0 then
   begin
    ErrorMessage('Выберите папку с файлами резюме и просканируйте их');
    Exit;
   end;

  if NameSender.Text='' then
   begin
    ErrorMessage('Введите имя отправителя');
    Exit;
   end;

  if SubjectMail.Text='' then
   begin
    ErrorMessage('Введите тему письма');
    Exit;
   end;

  begin
   ProgressBarTotal.Position:=0;
   ProgressBarFile.Position:=0;
   PanelProgress.Visible:=true;

    CountEmail:=ScanFilesList.Count;
    ValueEmail:=0;
    mess:=CreateMessageText;


   for SCard in ScanFilesList do
    begin
     ProgressBarFile.Position:=33;
     Log(tInfo,SCard.Email+' : '+IntToStr(i)) ;
     if SendMail(SCard.Email,mess,'')=0 then
      begin
       ProgressBarFile.Position:=100;
       ValueEmail:=ValueEmail+1;
       ProgressBarTotal.Position:=round(ValueEmail*100/CountEmail);

      end;

      ProgressBarFile.Position:=0;
    end;
    LabelInfoSending.Caption:=' Отправлено писем: '+IntToStr(ValueEmail)+'шт.';;
    PanelProgress.Visible:=false;
  end;






end;

procedure TForm9.Button1Click(Sender: TObject);
var
 f: TextFile;
 StartPos,EndPos,GavPos: integer;
 buf,str: string;
 flagSymbol,FindLeftPart, FindRightPart: boolean;
begin
 if OpenDialog2.Execute then
  begin
   AssignFile(f, OpenDialog2.FileName);
   try
    Reset(f);
    while not EOF(f) do
     begin
      readln(f, buf);

      if Pos('@',buf)<>0 then
       begin
        GavPos:=Pos('@',buf);
        str:='';
        StartPos:=GavPos-1;
        flagSymbol:=false;//false - допустимый символ
        FindLeftPart:=false;
        while (StartPos>=1) and (FindLeftPart=False) do
         begin
          if not (buf[StartPos] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then
           begin
            StartPos:=StartPos+1;
            if StartPos<GavPos then
             begin
              FindLeftPart:=true;
              str:=Copy(buf,StartPos,GavPos-StartPos+1);
              Memo1.Lines.Add(str);
             end
            else
             begin
              StartPos:=0;
             end;
           end;
          StartPos:=StartPos-1;
         end;

        if FindLeftPart=True then
         begin
          EndPos:=GavPos+1;
          FindRightPart:=False;
           while (EndPos<=Length(buf)) and (FindRightPart=False) do
            begin
             if (not (buf[EndPos] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.'])) or (EndPos=Length(buf)) then
              begin
               if EndPos<>Length(buf) then EndPos:=EndPos-1;
               if EndPos>GavPos then
                begin
                 FindRightPart:=true;
                 str:=str+Copy(buf,GavPos+1,EndPos-GavPos);
                 str:=Trim(str);
                 Memo1.Lines.Add(str);
                 if IsValidEmail(str) then Memo1.Lines.Add('IsValidEmail');

                end
               else
                begin
                 EndPos:=Length(buf);
                end;
              end;
             EndPos:=EndPos+1;
            end;
         end;
       end;

      if Pos('+7',buf)<>0 then
       begin
          StartPos:=Pos('+7',buf);
          str:=Copy(buf,StartPos,2);
          EndPos:=StartPos+Length(str);
          FindRightPart:=False;
          while (EndPos<=Length(buf)) and (FindRightPart=False) do
           begin
            if (buf[EndPos] in ['0'..'9']) then
             begin
              str:=str+Copy(buf,EndPos,1);
              if Length(str)=12 then
               begin
                FindRightPart:=true;
                Memo1.Lines.Add(str);
               end;
             end;
            EndPos:=EndPos+1;
           end;
       end;
     end;
     Memo1.Lines.Add('EndFile');
   finally
     CloseFile(f);
   end;
  end;

end;

function TForm9.ChangeEmailStatus(CHEmail, NStatus: string): boolean;
begin

end;

function TForm9.CheckAllowed(const s: string): boolean;
  var
    i: integer;
  begin
    Result:= false;
    for i:= 1 to Length(s) do
    begin
      if not (s[i] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then
        Exit;
    end;
    Result:= true;
end;

procedure TForm9.CheckBox1Click(Sender: TObject);
begin
 BitBtn11.Enabled:=true;
end;

procedure TForm9.CheckBoxAutoRunClick(Sender: TObject);
begin
  BitBtn11.Enabled:=true;

end;

function TForm9.CheckInternet: boolean;
begin

end;

function TForm9.CheckReg: boolean;
var
 sl,sMac: TStringList;
 i,j: integer;
 Base64 :TBase64Encoding;
begin
 Result:=true;   //изменить на false
  {
 sl:=TStringList.Create;
 sMac:=TStringList.Create;
 Base64 := TBase64Encoding.Create(0);
 try
 GetDir(0,path);
 sl.LoadFromFile(path+'\acnt64.dll');
 Log(tError,'Befor GetSysInfo');
 try
    sMac:=GetSysInfo(sMac);
 except on E: Exception do
   Log(tError,'ERROR '+ e.Message);
 end;

  Log(tError,'After GetSysInfo');

 i:=0;
 while i<sMac.Count do
 begin
  j:=0;
  while j<sl.Count do
  begin
   if Base64.Decode(sl[j])=sMac[i] then
    begin
     Result:=true;
     exit;
    end;
   j:=j+1;
  end;
  i:=i+1;
 end;
 finally
   sl.Free;
   Base64.Free;
 end;
 }
end;

function TForm9.CheckSend: boolean;
begin

end;

function TForm9.ConnectAccount: boolean;
begin
 Result:=false;
 if ((MyLogin.Text='')or(MyPost.ItemIndex=-1)) then
  ErrorMessage('Проверьте правильность ввода логина и домена Вашей эл.почты')
 else
 begin

// SMTP := TIdSMTP.Create(Application);
  SMTP.Host := SMTP_Host;
  SMTP.Port := SMTP_Port;
  SMTP.AuthType := SMTP_AuthType;
  SMTP.Username := MyEmailAddress.Caption;{Должно совпадать с msg.From.Address}
  SMTP.Password := MyPass.Text;

  //это необходимо использовать для SSL
 // SSLOpen := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  SSLOpen.Destination := SMTP.Host+':'+IntToStr(SMTP.Port);
  SSLOpen.Host := SMTP.Host;
  SSLOpen.Port := SMTP.Port;
  SSLOpen.DefaultPort := 0;
  SSLOpen.SSLOptions.Method := sslvSSLv23;
  //SSLOpen.SSLOptions.SSLVersions:=[sslvSSLv23];
  SSLOpen.SSLOptions.Mode := sslmUnassigned;

  SMTP.IOHandler := SSLOpen;
  SMTP.UseTLS := utUseImplicitTLS;
  try

   SMTP.Connect;
   Result:=true;
  except on E: Exception do
   begin

    Log(tError,' ------ Error ConnectAccount: '+E.Message);
    if SMTP.Connected then SMTP.Disconnect;
    if Pos('Socket Error',E.Message)>0 then
     ShowMessage('Проверьте подключение к интернету  ')
    else
     ShowMessage('Проверьте данные аккоунта  ')
   end;
  end;
 end;

end;

function TForm9.ConnectBase: boolean;
begin
 Result:=false;
 DM.FDConnection1.Connected:= True;
   if DM.FDConnection1.Connected then
    begin
     Result:=true;
    end;
end;

function TForm9.CreateMessageText: string;
var
 str: string;
 ind: integer;
begin
 str:='';
 for ind := 0 to MyTextMessage.Lines.Count-1 do
  begin
   str:=str+MyTextMessage.Lines[ind]+'<br>';
  end;
 Result:=str;

end;

procedure TForm9.DateEndSendingChange(Sender: TObject);
begin
 FlagChangeDataSender:=true;
end;

procedure TForm9.Delay(msec: integer);
begin
 TimerDelay.Interval:=msec; // устанавливаем паузу в мс
 TimerDelay.Enabled:=True;// включаем таймер
 while TimerDelay.Enabled do  // пока таймер включен, держим паузу
 begin
  Application.ProcessMessages;
 end;
 TimerDelay.Enabled:=false;
end;

function TForm9.DisconnectBase: boolean;
begin
 DM.FDConnection1.Connected:= false;
 ShowMessage('FDConnection is disconnected.');
 Result:=true;
end;

procedure TForm9.Edit1Change(Sender: TObject);
begin
 if Edit1.Text='12356' then
  begin
    TabSheet4.TabVisible:=true;
  end;
end;

procedure TForm9.ErrorMessage(str: string);
begin
 ShowMessage(str);

end;

procedure TForm9.ExecSending;
var
 Sender:TObject;
begin

 if ( now>DateEndSending.Date) then
  begin
   ErrorMessage('Период рассылки окончен. Для возобновления установите новый период');
   Log(tError,DateTimeToStr(Now)+ ' : Период рассылки окончен');
   BitBtn5Click(Sender);
   Exit;
  end;


 if not ConnectBase then
  begin
   Log(tError,DateTimeToStr(Now)+ ' : не удалось подключиться к базе');
   BitBtn5Click(Sender);
   Exit;
  end;

 if (not AccountConnect) then
  begin
   ErrorMessage('Необходимо подключиться к Вашему Email аккоунту');
   Log(tError,DateTimeToStr(Now)+ ' : Email аккоунт не подключен');
   BitBtn5Click(Sender);
   Exit;
  end;

  TodaySendingEmail:=0;
  RepeatSend;

end;

procedure TForm9.FormCreate(Sender: TObject);
begin
Registered:=false;
FlagSender:=false;
FlagRecepiant:=false;
FlagMessage:=false;
BitBtnSendEmail.Enabled:=false;
PageControl1.TabIndex:=0;
PageControlTextLetter.TabIndex:=0;
TabSheet4.TabVisible:=false;
TabSheet3.TabVisible:=false;

PathExe:=ExtractFilePath(Application.ExeName);
 PathExe:=ExtractFileDir(PathExe);
PathImages:=PathExe+'\Images\';
PathAcnt:=TPath.GetCachePath;

Log(tError,DateTimeToStr(Now)+ ' PathAcnt: '+PathAcnt);


MyEMailAddress.StyleElements := [];
MyEMailAddress.Font.Color:=clRed;

HardwareInfo;

if not FileExists(PathAcnt+'\acnt64.dll') then //убрать NOT  для проверки аккоунта
 begin
  if CheckReg then
   begin

    Registered:=true;
    StatusBar1.Panels[2].Text:='Registered';
    if FileExists(PathAcnt+'\acnt32.dll') then
     begin
      LoadAcnt;
      MyPostChange(Sender);
      FlagChangeDataSender:=false;
     end;

    if FileExists(PathExe+'\config.cfg') then
     begin
      LoadConfig;
     end;

   end
  else
   begin
    Registered:=false;
    StatusBar1.Panels[2].Text:='Unregistered';
   end;

 end
else
 begin
  //SaveMac;  проверить аккаунт
  Registered:=true;
  StatusBar1.Panels[2].Text:='Unregistered';
 end;

 Panel18.Visible:=false;
end;

procedure TForm9.ImageSendClick(Sender: TObject);
var
 FName: string;
 FBitmap:TBitmap;
 bmp: TBitmap;
 jpg: TJPEGImage;
 png: TPngImage;
 Stream: TMemoryStream;
begin

 if not DirectoryExists(PathImages) then
  CreateDir(PathImages);


 OpenPictureDialog1.InitialDir:=PathImages;
 if OpenPictureDialog1.Execute then
  begin
   FlagChangeDataSender:=true;
   FName:=OpenPictureDialog1.FileName;
   ImageSend.Picture.LoadFromFile(FName);
     SendImageFullPath:= PathImages+ ExtractFileName(FName) ;

  FlagSendImage:=true;

  if (FName<>SendImageFullPath) then
   begin
    Log(tError, 'FullPath : '+SendImageFullPath);
    ImageSend.Picture.SaveToFile(SendImageFullPath);
    ImagePath.Caption:= SendImageFullPath;
    Log(tError,DateTimeToStr(Now)+ ' : '+'картинка добавлена');
   end;
  end;
end;

procedure TForm9.ImageUnVisibleClick(Sender: TObject);
begin
 ImageVisible.Visible:=true;
 MyPass.PasswordChar:='*';
 ImageUnVisible.Visible:=false;
end;

procedure TForm9.ImageVisibleClick(Sender: TObject);
begin
 ImageUnVisible.Visible:=true;
 MyPass.PasswordChar:=#0;
 ImageVisible.Visible:=false;
end;

function TForm9.IsValidEmail(const Value: string): boolean;
var
  i: integer;
  namePart, serverPart: string;
begin
  Result:= false;
  i:= Pos('@', Value);
  if i = 0 then
    Exit;
  if Pos('@', Value,i+1)<>0 then Exit;


  namePart:= Copy(Value, 1, i - 1);
  serverPart:= Copy(Value, i + 1, Length(Value));
  if (Length(namePart) = 0) or ((Length(serverPart) < 5)) then
    Exit;
  i:= Pos('.', serverPart);
  if (i = 0) or (i > (Length(serverPart) - 2)) then
    Exit;
  if Pos('.', serverPart,i+1)<>0 then
    Exit;
  Result:= CheckAllowed(namePart) and CheckAllowed(serverPart);

end;

procedure TForm9.LoadAcnt;
var
 sl,sMac: TStringList;
 fname: string;
 i:integer;
 Base64 : TBase64Encoding;
begin

  sl := TStringList.Create;
  sl.LoadFromFile(PathAcnt+'\acnt32.dll');

  Base64 := TBase64Encoding.Create(0);
  try
     begin
      MyLogin.Text:=Base64.Decode(sl[0]);
      MyPost.ItemIndex:=StrToInt(Base64.Decode(sl[1]));
      MyPass.Text:=Base64.Decode(sl[2]);
      NameSender.Text:=Base64.Decode(sl[3]);
      SubjectMail.Text:=Base64.Decode(sl[4]);
      NameSenderImage.Text:=Base64.Decode(sl[5]);
      SubjectMailImage.Text:=Base64.Decode(sl[6]);
      MessageTextImage.Text :=Base64.Decode(sl[7]);
      ImagePath.Caption:=Base64.Decode(sl[8]);
      DateEndSending.Date:=StrToDate(Base64.Decode(sl[9]));

      i:=10;
      MyTextMessage.Lines.Clear;
      while i<sl.Count do
       begin
        MyTextMessage.Lines.Add(Base64.Decode(sl[i]));
        i:=i+1;
       end;

      if ImagePath.Caption='default' then
       begin
        ImageSend.Picture.Assign(ImageDefault.Picture);
       end
      else
       begin
        if FileExists(ImagePath.Caption) then
         begin
          ImageSend.Picture.LoadFromFile(ImagePath.Caption);
          //FlagSendImage:=true;
         end
        else
         begin
          ImagePath.Caption:='default';
          ImageSend.Picture.Assign(ImageDefault.Picture);
          SaveAcnt;
         end;
       end;
     end;
  finally
    sl.Free;
    Base64.Free;
  end;

end;

procedure TForm9.LoadConfig;
var
 sl: TStringList;
 fname: string;
begin

  sl := TStringList.Create;
  sl.LoadFromFile(PathExe+'\config.cfg');
  try
     begin
      CountEmailSendInDay.Value:= StrToInt(sl[0]);
      MinDaysToSend.Value:=StrToInt(sl[1]);
     end;


    Log(tSystem,DateTimeToStr(Now)+' LoadConfig  : настройки загружены ');
    BitBtn9.Enabled:=false;
  finally
    sl.Free
  end;

end;

procedure TForm9.Log(lType: TTypeMessge; lStr: string);
begin
 case lType of
   tInfo: RichLog.SelAttributes.Color:=clBlack;
   tError: RichLog.SelAttributes.Color:=clRed;
   tSystem: RichLog.SelAttributes.Color:=clBlue;
 end;
 RichLog.Lines.Add(lStr);
 RichLog.SelStart:=Length(RichLog.text);
 RichLog.SelLength:=0;
 SendMessage(RichLog.Handle, EM_SCROLLCARET, 0, 0);
 //RichLog.perform(EM_LINESCROLL,0,RichLog.lines.count);
end;

procedure TForm9.MemoEmailFromFilesChange(Sender: TObject);
begin
 if MemoEmailFromFiles.Lines.Count=0 then
  BitBtnSendEmail.Enabled:=false
 else
  BitBtnSendEmail.Enabled:=true;
end;

procedure TForm9.MessageTextImageChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

procedure TForm9.MyLoginChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

procedure TForm9.MyLoginClick(Sender: TObject);
begin
MyLogin.Font.Color:=clBlack;
end;

procedure TForm9.MyLoginExit(Sender: TObject);
begin
if MyPost.ItemIndex>-1 then
 begin
   if not IsValidEmail(MyLogin.Text+MyGav.Caption+MyPost.Items.Strings[MyPost.ItemIndex]) then
    begin
     MyEMailAddress.StyleElements := [];
     MyEMailAddress.Font.Color:=clRed;
     MyLogin.Text:='';
     FlagSender:=false;
     ErrorMessage(' Ваш E-mail введен некорректно');
    end
   else
    begin
     MyEMailAddress.StyleElements := [];
     MyEMailAddress.Font.Color:=clGreen;
     MyEmailAddress.Caption:=MyLogin.Text+MyGav.Caption+MyPost.Items.Strings[MyPost.ItemIndex];
     FlagSender:=true;
    end;
 end
else
 FlagSender:=false;
end;

procedure TForm9.MyPassChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

procedure TForm9.MyPostChange(Sender: TObject);
begin
MyLoginExit(Sender);

FlagChangeDataSender:=true;

case MyPost.ItemIndex of
 0: begin
     SMTP_Host:='smtp.mail.ru';
     SMTP_Port := 465; //25
     SMTP_AuthType := satDefault;
    end;
 1: begin
     SMTP_Host:='smtp.yandex.ru';
     SMTP_Port := 465;
     SMTP_AuthType := satDefault;

    end;
 2: begin
     SMTP_Host:='mail.tut.by';
     SMTP_Port := 2525;  //587  465  995
     SMTP_AuthType := satDefault;
    end;
    3: begin
     SMTP_Host:='smtp.rambler.ru';
     SMTP_Port := 465;  //587  465  995
     SMTP_AuthType :=  satDefault;
    end;
// mail.ru
//yandex.ru
//gmail.com

end;
end;



procedure TForm9.MyTextMessageChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

procedure TForm9.NameSenderChange(Sender: TObject);
begin
FlagChangeDataSender:=true;
end;

procedure TForm9.RepeatSend;
var
 SendResult: integer;
 RRecipientEmail: string;
 sender: tobject;
begin
 RRecipientEmail:=GetEmailFromBase;
 if RRecipientEmail=BeforEmail then SpamCountSend:=SpamCountSend+1;

// MemoLog.Lines.Add('error   SMTP.Send : '+E.ClassName+' поднята ошибка, с сообщением : '+E.Message);


 if ((TodaySendingEmail<CountEmailSendInDay.Value)and(RRecipientEmail<>'empty')) then
  begin
   //Log(tInfo,DateTimeToStr(Now)+ '  :  '+'получен Email для отправки: '+RRecipientEmail);

   if IsValidEmail(RRecipientEmail) then
    SendResult:=SendMail(RRecipientEmail,CreateMessageText,'')
   else
    SendResult:=2;

   case SendResult of
    0: begin
         Log(tSystem,DateTimeToStr(Now)+ ' отправлено : '+RRecipientEmail);
         TodaySendingEmail:=TodaySendingEmail+1;
         StatusBar1.Panels[1].Text:=IntToStr(TodaySendingEmail)+'/'+IntToStr(CountEmailSendInDay.Value);
         DM.FDQuery1.Edit;
         DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=Now;
         DM.FDQuery1.FieldByName('StatusEmail').AsString:='Ok';
         DM.FDQuery1.Post;
         SpamCountSend:=0;
         TimerSendEmail.Interval:=300;
         TimerSendEmail.Enabled:=true;
       end;
    1: begin  //SPAM
        Log(tError,DateTimeToStr(Now)+ ' Error№1: Spam : '+RRecipientEmail);
        BeforEmail:=RRecipientEmail;
        if SpamCountSend=5 then
         begin
          DM.FDQuery1.Edit;
          DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=IncDay(Now,1-MinDaysToSend.Value);
          DM.FDQuery1.Post;
          Log(tSystem,RRecipientEmail+' - отправка невозможна, подозрение на спам, попробуем позже');
         end;
        TimerSendEmail.Interval:=5000*SpamCountSend;
        TimerSendEmail.Enabled:=true;
       end;
    2: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error: №2 : '+RRecipientEmail);

         DM.FDQuery1.Edit;
         DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=Now;
         DM.FDQuery1.FieldByName('StatusEmail').AsString:='Error';
         DM.FDQuery1.Post;
         TimerSendEmail.Interval:=300;
         TimerSendEmail.Enabled:=true;
       end;
    3: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error№3: Account not Connected');
        ErrorMessage('Обрыв подключения к аккаунту ');
        BitBtn5Click(sender);
       end;
    4: begin  //Error
        Log(tError,DateTimeToStr(Now)+' Error№4: unnown error');
        Log(tError,DateTimeToStr(Now)+ '  :  '+RRecipientEmail);

        DM.FDQuery1.Edit;
         DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=Now;
         DM.FDQuery1.FieldByName('StatusEmail').AsString:='Error';
         DM.FDQuery1.Post;
         TimerSendEmail.Interval:=300;
         TimerSendEmail.Enabled:=true;

        {ErrorMessage('Непредвиденный сбой... ');
        BitBtn5Click(sender);
        }
       end;
    5: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error№5: Message rejected under suspicion of SPAM');
        //ErrorMessage('Сообщение отклонено по подозрению в спаме ');
         DM.FDQuery1.Edit;
         DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=IncDay(Now,1-MinDaysToSend.Value);
         DM.FDQuery1.Post;
         TimerSendEmail.Interval:=5000;
         TimerSendEmail.Enabled:=true;

       end;
    6: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error №6: Socket Error');
        ErrorMessage('Отсутствует интернет ');
        BitBtn5Click(sender);
       end;
    7: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error: №7: '+RRecipientEmail+' данный адрес не существует');

         DM.FDQuery1.Edit;
         DM.FDQuery1.FieldByName('LastDateEmail').AsDateTime:=Now;
         DM.FDQuery1.FieldByName('StatusEmail').AsString:='Error';
         DM.FDQuery1.Post;
         TimerSendEmail.Interval:=200;
         TimerSendEmail.Enabled:=true;
       end;
    8: begin  //Error
        Log(tError,DateTimeToStr(Now)+ ' Error №8: Too many message from ... Try again later');
         ErrorMessage('Почтовый сервер приостановил отправку писем с данного аккаунта из-за подозрения на СПАМ. Перезапустите рассылку чере 1 час ');
        BitBtn5Click(sender);
        BitBtn1Click(sender);
       end;
   end;
  end
 else
 begin

  if(TodaySendingEmail=CountEmailSendInDay.Value)then
   begin
    Log(tSystem,DateTimeToStr(Now)+ 'отправка пакета завершена, отправлено писем '+IntToStr(TodaySendingEmail)+'шт.');
    ShowTrayBalloon(bfInfo,'Information','Отправлено '+IntToStr(TodaySendingEmail)+' сообщений');
    SetExecTimer;
    StatusBar1.Panels[1].Text:='';
   end;


  if(RRecipientEmail='empty') then
   begin
    Log(tSystem,DateTimeToStr(Now)+ 'Отправлено писем '+IntToStr(TodaySendingEmail)+'шт., нет подходящих Email');
    ShowTrayBalloon(bfWarning,'Warning','Отправлено '+IntToStr(TodaySendingEmail)+' сообщений, подходящих Email нет');
    SetExecTimer;
    StatusBar1.Panels[1].Text:='';
   end;
 end;

end;

procedure TForm9.SaveAcnt;
var
 sl,sMac: TStringList;
 fname: string;
 i:integer;
 Base64 : TBase64Encoding;
begin

  sl := TStringList.Create;
  Base64 := TBase64Encoding.Create(0);
  try
     begin
      sl.Add(Base64.Encode(MyLogin.Text));
      sl.Add(Base64.Encode(IntToStr(MyPost.ItemIndex)));
      sl.Add(Base64.Encode(MyPass.Text));
      sl.Add(Base64.Encode(NameSender.Text));
      sl.Add(Base64.Encode(SubjectMail.Text));
      sl.Add(Base64.Encode(NameSenderImage.Text));
      sl.Add(Base64.Encode(SubjectMailImage.Text));
      sl.Add(Base64.Encode(MessageTextImage.Text));
      sl.Add(Base64.Encode(ImagePath.Caption));
      sl.Add(Base64.Encode(DateToStr(DateEndSending.Date)));

      i:=0;
      while i<MyTextMessage.Lines.Count do
       begin
        sl.Add(Base64.Encode(MyTextMessage.Lines[i]));
        i:=i+1;
       end;

      sl.SaveToFile(PathAcnt+'\acnt32.dll');
      Log(tSystem,DateTimeToStr(Now)+' SaveAcnt  : настройки сохранены ');
      Log(tSystem,DateTimeToStr(Now)+' : '+PathAcnt+'\acnt32.dll');
     end;
  finally
    sl.Free;
    Base64.Free;
  end;
end;

procedure TForm9.SaveCFG;
var
 sl: TStringList;
 fname: string;
begin

  sl := TStringList.Create;
  try
     begin
      sl.Add(inttostr(CountEmailSendInDay.Value));
      sl.Add(inttostr(MinDaysToSend.Value));
     end;

    sl.SaveToFile(PathExe+'\config.cfg');
    Log(tSystem,DateTimeToStr(Now)+' SaveCFG  : настройки сохранены ');
    BitBtn9.Enabled:=false;
  finally
    sl.Free
  end;

end;

procedure TForm9.SaveInDataBase;
begin
   SaveInDBaseThread := TSaveInDBaseThread.create(true);
   SaveInDBaseThread.freeonterminate := false;
   SaveInDBaseThread.priority := tpNormal;

   SaveInDBaseThread.Start;

   SaveInDBaseThread.WaitFor;
   FreeAndNil(SaveInDBaseThread);
end;

// Информация о компьютере.
procedure TForm9.HardwareInfo;
var Size : cardinal;
PRes : PChar;
BRes : boolean;
lpSystemInfo : TSystemInfo;
laCompName_,laUserName_ ,laCPU_ :string;
begin
// Имя компьютера
Size := MAX_COMPUTERNAME_LENGTH + 1;
PRes := StrAlloc(Size);
BRes := GetComputerName(PRes, Size);
if BRes then
 begin
  laCompName_ := StrPas(PRes);
  Log(tSystem,DateTimeToStr(Now)+' CompName  : '+laCompName_);
 end;
// Имя пользователя
Size := MAX_COMPUTERNAME_LENGTH + 1;
PRes := StrAlloc(Size);
BRes := GetUserName(PRes, Size);
if BRes then
 begin
  laUserName_ := StrPas(PRes);
  Log(tSystem,DateTimeToStr(Now)+' UserName  : '+laUserName_);
 end;
// Процессор
GetSystemInfo(lpSystemInfo);
laCPU_ := 'класса x' + IntToStr
(lpSystemInfo.dwProcessorType);
  Log(tSystem,DateTimeToStr(Now)+' Processor  : '+laCPU_ );

end;

procedure TForm9.SaveMac;
var
 sl,sMac :TStringList;
 Base64 :TBase64Encoding;
 i:integer;
begin
 HardwareInfo;

{
 sl := TStringList.Create;
 sMac := TStringList.Create;
  Base64 := TBase64Encoding.Create(0);
   try
     begin
      sMac:=GetSystemInfo(sMac);
      i:=0;

      while i<sMac.Count do
      begin
       sl.Add(Base64.Encode(sMac[i]));
       i:=i+1;
      end;
      sl.SaveToFile(PathAcnt+'\acnt64.dll');
     end;
   finally
     sl.Free;
     sMac.Free;
     Base64.Free;
   end;
   }
end;

procedure TForm9.ScanFiles;
var
 i: integer;
 s,str: string;
 C: char ;
 FName,FEmail,FPhone: string;
begin


 MemoEmailFromFiles.Lines.Clear;
 if chosenDirectory='' then
  begin
   ErrorMessage('Выбирете папку с файлами');
   Exit;
  end;

  ScanFilesList:= TList.Create;
  ValueFiles:=0;
  ValueEmail:=0;
  ValuePhone:=0;
  CountFiles:=0;
  if CheckBoxDOC.Checked then
    CountFiles:=CountFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.doc'));
  if CheckBoxRTF.Checked then
    CountFiles:=CountFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.rtf'));
  if CheckBoxTXT.Checked then
    CountFiles:=CountFiles+Length(TDirectory.GetFiles(chosenDirectory, '*.txt'));


  ProgressBarTotal.Position:=0;
  ProgressBarFile.Position:=0;
  PanelProgress.Visible:=true;


 if CountFiles>0 then
  begin

   FindingThread := TMyFindingThread.create(true);
   FindingThread.freeonterminate := false;
   FindingThread.priority := tpNormal;

   FindingThread.Start;

   FindingThread.WaitFor;
   FreeAndNil(FindingThread);
  end;


end;



function TForm9.SendMail(RecipientEmail, MyMessage: string;SendFile: string): integer;
var
  msg: TIdMessage;
  att: TIdAttachmentFile;
  cid : WideString;
begin
 Result:=-1; //error number1

 try

  msg := TIdMessage.Create(Application);

  if PageControlTextLetter.TabIndex=0 then
   begin

  if SendFile='' then
   begin
    msg.ContentType:='text/html; charset=windows-1251';
    msg.Subject := SubjectMail.Text+ '               :';//+DateTimeToStr(Now);
    msg.From.Name := NameSender.Text;
    msg.Body.Text:=MyMessage+'<br>'+DateTimeToStr(Now);
  end
  else
   begin
    msg.ContentType := 'multipart/mixed; charset=windows-1251' ;
    msg.Subject :='База телефонов';
    msg.From.Name := NameSender.Text;
    msg.Body.Text:='';
     if (FileExists(PathExe+SendFile)) then
      begin
        TIdAttachmentFile.Create(msg.MessageParts, PathExe+SendFile);
      end;
   end;
   end
  else
   begin
    if FlagSendImage then
     begin
      msg.ContentType := 'html; charset=windows-1251' ;
      msg.Subject := SubjectMailImage.Text+ '               :';//+DateTimeToStr(Now);
      msg.From.Name := NameSenderImage.Text;

       with TIdMessageBuilderHtml.Create do
        begin
         PlainTextContentTransfer:='html';
         PlainTextCharset:='windows-1251';
         HTML.Text:='<html><body> <br><br> <img src="cid:'+ExtractFileName(SendImageFullPath)+'" height=100% width=100% >  </body></html>';
         HTMLCharset:='windows-1251';
         cid:=ExtractFileName(SendImageFullPath);
         HtmlFiles.Add(SendImageFullPath,cid);
         FillMessage(msg);
         //Memo3.Lines.Text:=HTML.Text;
        end;
       end
      else
       begin
        ErrorMessage(' Выберите картинку для отправки');
        Exit;
       end;
   end;


  msg.From.Address := MyEmailAddress.Caption; {&lt;&lt;Должно совпадать с SMTP.UserName}
  msg.Recipients.EMailAddresses :=RecipientEmail;

  if not LoadOpenSSLLibrary then ShowMessage('not load ssl');

  if SMTP.Connected then
    begin
     try
      SMTP.Send(msg);
      Result:=0; //without error
     // Log(tInfo,'info   Send EMAIL to '+RecipientEmail);
     except on E: Exception do
      begin
      result:=4; //another error
       if Pos('timeout exceeded',E.Message)>0 then Result:=2;//Error: timeout exceeded
       if Pos('SPAM',E.Message)>0 then Result:=1;//Error: SPAM
       if Pos('Socket Error',E.Message)>0 then Result:=6;//internet non found
       if Pos('invalid mailbox',E.Message)>0 then Result:=7;//user not found
       if Pos('Message rejected under ',E.Message)>0 then Result:=5;//Error: timeout exceeded
       if Pos('Try again later',E.Message)>0 then Result:=8;//Too many message from

      // Log(tError,'error   SMTP.Send № '+IntToStr(Result));//
      // if Result=4 then
        Log(tError,'error   SMTP.Send № '+IntToStr(Result)+' : '+E.ClassName+' поднята ошибка, с сообщением : '+E.Message);
      end;
     end;

    end
  else
   begin
    Result:=3;//error not connected
   end;

 // SMTP.Disconnect();
 // SMTP.Free;
 finally
  msg.Free;
 end;

end;

procedure TForm9.SetExecTimer;
var
 sDate,sTime: string;
 MinuttenCount: Double;
begin

 ShowTrayBalloon(bfInfo,'Information','Рассылка завершена, отправлено  '+ IntToStr(TodaySendingEmail) + ' писем');
 Log(tSystem,DateTimeToStr(Now)+ '  SetExecTimer  '+'Рассылка завершена, отправлено  '+ IntToStr(TodaySendingEmail) + ' писем');



end;

procedure TForm9.ShowTrayBalloon(tm: TBalloonFlags; title, msg: string);
begin
TrayIcon1.visible:=true; // делаем значок в трее видимым
TrayIcon1.BalloonFlags:=tm;
trayicon1.balloontitle:=(title);
trayicon1.balloonhint:=(msg);
trayicon1.showballoonHint;// показываем наше уведомление
end;

{ TCard }

constructor TCard.Create(FEmail, FPhoneNumber: string);
begin
 Self.Email:=FEmail;
 Self.PhoneNumber:=FPhoneNumber;
 Self.EmailLastSendDate:= EncodeDate(2000, 1, 1);
 Self.PhoneNumberLastSendDate:= EncodeDate(2000, 1, 1);
end;

{ TMyFindingThread }

procedure TMyFindingThread.DOC_Files;
var
 f: TextFile; // файл
 buf: string; // буфер для чтения из файла
 str, ExeFile: string;
 fCard: TCard;
 StartPos,EndPos,GavPos, ExtFiles: integer;
 s: string;
 flagSymbol,FindLeftPart,FindRightPart: boolean;
begin
 ValueFiles:=0;
 for ExtFiles := 0 to 2 do
  begin
   ExeFile:='';
   case ExtFiles of
    0: begin
        if Form9.CheckBoxDOC.Checked then
         ExeFile:='*.doc';
       end;
    1: begin
        if Form9.CheckBoxRTF.Checked then
         ExeFile:='*.rtf';
       end;
    2: begin
        if Form9.CheckBoxtxt.Checked then
         ExeFile:='*.txt';
       end;
   end;

if ExeFile<>'' then

 for s in TDirectory.GetFiles(chosenDirectory, ExeFile) do
  begin
   FName:=s;
   AssignFile(f, fName);
   try
    Reset(f);
    ValueProgressScanFile:=0;
    FEmail:='';
    FPhone:='';
    Synchronize(progress);
    while not EOF(f) do
     begin
      readln(f, buf);
      if Pos('@',buf)<>0 then
       begin
        GavPos:=Pos('@',buf);
        str:='';
        StartPos:=GavPos-1;
        flagSymbol:=false;//false - допустимый символ
        FindLeftPart:=false;
        while (StartPos>=1) and (FindLeftPart=False) do
         begin
          if not (buf[StartPos] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then
           begin
            StartPos:=StartPos+1;
            if StartPos<GavPos then
             begin
              FindLeftPart:=true;
              str:=Copy(buf,StartPos,GavPos-StartPos+1);
             end
            else
             begin
              StartPos:=0;
             end;
           end;
          StartPos:=StartPos-1;
         end;

        if FindLeftPart=True then
         begin
          EndPos:=GavPos+1;
          FindRightPart:=False;
           while (EndPos<=Length(buf)) and (FindRightPart=False) do
            begin
             if (not (buf[EndPos] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.'])) or (EndPos=Length(buf)) then
              begin
               if EndPos<>Length(buf) then EndPos:=EndPos-1;
               if EndPos>GavPos then
                begin
                 FindRightPart:=true;
                 str:=str+Copy(buf,GavPos+1,EndPos-GavPos);
                 str:=Trim(str);
                 if Form9.IsValidEmail(str) then
                  begin
                   FEmail:=str;
                   ValueEmail:=ValueEmail+1;
                  end;
                 ValueProgressScanFile:=33;
                 Synchronize(progress);
                end
               else
                begin
                 EndPos:=Length(buf);
                end;
              end;
             EndPos:=EndPos+1;
            end;
         end;
       end;

      if Pos('+7',buf)<>0 then
       begin
         FPhone:='';
          StartPos:=Pos('+7',buf);
          str:=Copy(buf,StartPos,2);
          EndPos:=StartPos+Length(str);
          FindRightPart:=False;
          while (EndPos<=Length(buf)) and (FindRightPart=False) do
           begin
            if (buf[EndPos] in ['0'..'9']) then
             begin
              str:=str+Copy(buf,EndPos,1);
              if Length(str)=12 then
               begin
                FindRightPart:=true;
                FPhone:=str;
               end;
             end;
            EndPos:=EndPos+1;
           end;
       end;

     end;
      ValueProgressScanFile:=100;
        if ((FEmail<>'') or (FPhone<>'')) then
         begin
          ScanFilesList.Add(TCard.Create(FEmail,FPhone));
         end;
      ValueFiles:=ValueFiles+1;
      Synchronize(Progress);
   finally
     CloseFile(f);
   end;
  end;
 end;
end;

procedure TMyFindingThread.EndThread;
begin
begin

end;
end;

procedure TMyFindingThread.Execute;
begin
 inherited;
 FindProg;

 Synchronize(EndThread);

end;

procedure TMyFindingThread.FindProg;
begin
     DOC_Files;
end;

procedure TMyFindingThread.PDF_Files;
begin

end;

procedure TMyFindingThread.PrintLog(str: string);
begin

end;

procedure TMyFindingThread.Progress;
var
 str: string;
begin
 Form9.ProgressBarTotal.Position:=round(ValueFiles*100/CountFiles);
 Form9.ProgressBarFile.Position:=ValueProgressScanFile;
 str:='';
 if ValueProgressScanFile=100 then
  begin
   str:=FName;
   delete(str,1,length(chosenDirectory)+1);
   str:=str+' : '+FEmail+' : '+FPhone;
   Form9.MemoEmailFromFiles.Lines.Add(str);
  end;
 //Form1.MemoFindingEMail.Lines.Add(TCard(FindingFiles.Last).Email);
 Form9.LabelProgressFile.Caption:=FName;
 Form9.LabelProgressTotal.Caption:=IntToStr(ValueFiles)+' / '+IntToStr(CountFiles);
 Application.ProcessMessages;
end;





{ TSaveInDBaseThread }

procedure TSaveInDBaseThread.Execute;
begin
  inherited;
  SaveInDBase;
end;

procedure TSaveInDBaseThread.Progress;
begin
 Form9.ProgressBarFile.Position:=count;
 Form9.LabelProgressFile.Caption:='Добавлено новых Email адресов - '+IntToStr(count);
 Form9.ProgressBarTotal.Position:=total;
 Form9.LabelProgressTotal.Caption:='Обработано Email адресов - '+IntToStr(total)+' / '+IntToStr(Form9.ProgressBarTotal.Max);
 Application.ProcessMessages;

end;

procedure TSaveInDBaseThread.SaveInDBase;
 var
  SCard: TCard;
begin
        count:=0;
        total:=0;
           for SCard in  ScanFilesList do
            begin
             if not Form9.FindEmailInBase(SCard.Email) then
              begin
               Form9.AppendCardToBase(SCard.Email,SCard.PhoneNumber);
               count:=count+1;
              end;
             total:=total+1;
             Synchronize(Progress);
            end;

end;

end.
