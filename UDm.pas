unit UDm;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.SQLite,
  FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs, FireDAC.VCLUI.Wait,
  FireDAC.Comp.UI, Data.DB, FireDAC.Comp.Client,Dialogs, FireDAC.Stan.Param,
  FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet;

type
  TDM = class(TDataModule)
    FDConnection1: TFDConnection;
    FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDQuery1: TFDQuery;
    DataSource1: TDataSource;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;

implementation

uses
  Winapi.Windows;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

procedure TDM.DataModuleCreate(Sender: TObject);
var
 path: string;
begin
 GetDir(0,path);
 if not FileExists(path+'\Contacts.db') then
  begin
   if FileExists(path+'\Contacts empty.db') then
    begin
     CopyFile(PChar(path+'\Contacts empty.db'),PChar(path+'\Contacts.db'),false);
    end;
  end;
 FDConnection1.Params.Add('Database='+path+'\Contacts.db');
 //FDConnection1.Params.Database := path+'\Contacts.db';
 try
  FDConnection1.Connected:=true;
 except on E: Exception do
   ShowMessage('error ConnectBase: '+ E.ClassName+' поднята ошибка, с сообщением : '+E.Message);
 end;

end;

end.
