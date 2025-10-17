unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellAPI, Registry;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    EditPorta: TEdit;
    ButtonAplicar: TButton;
    MemoLog: TMemo;
    procedure ButtonAplicarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure Log(const Msg: string);
    procedure ExecutarPowerShell(const Comando: string);
    function LerPortaRDP: Integer;
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Log(const Msg: string);
begin
  MemoLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + ' - ' + Msg);
end;

procedure TForm1.ExecutarPowerShell(const Comando: string);
begin
  // Executa um comando PowerShell elevado (modo administrador)
  ShellExecute(0, 'runas', 'powershell.exe',
    PChar('-Command "' + Comando + '"'), nil, SW_HIDE);
end;

function TForm1.LerPortaRDP: Integer;
var
  Reg: TRegistry;
begin
  Result := 3389; // padrão
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly('SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp') then
    begin
      try
        Result := Reg.ReadInteger('PortNumber');
      except
        on E: Exception do
          Log('Erro ao ler porta: ' + E.Message);
      end;
    end
    else
      Log('Chave do Registro não encontrada.');
  finally
    Reg.Free;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  PortaAtual: Integer;
begin
  PortaAtual := LerPortaRDP;
  EditPorta.Text := IntToStr(PortaAtual);
  Log('Porta atual do RDP: ' + IntToStr(PortaAtual));
end;

procedure TForm1.ButtonAplicarClick(Sender: TObject);
var
  Porta: string;
  Cmd: string;
begin
  Porta := Trim(EditPorta.Text);

  if Porta = '' then
  begin
    ShowMessage('Informe o número da porta.');
    Exit;
  end;

  try
    // 1️⃣ Altera a porta do RDP no registro
    Cmd := Format('Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "PortNumber" -Value %s', [Porta]);
    ExecutarPowerShell(Cmd);
    Log('Porta alterada para ' + Porta);

    // 2️⃣ Adiciona exceção no firewall
    Cmd := Format('New-NetFirewallRule -DisplayName "Acesso RDP Porta %s" -Direction Inbound -Protocol TCP -LocalPort %s -Action Allow -Profile Any', [Porta, Porta]);
    ExecutarPowerShell(Cmd);
    Log('Regra de firewall criada para a porta ' + Porta);

    // 3️⃣ Reinicia o serviço de Terminal Server
    ExecutarPowerShell('Restart-Service -Name TermService -Force');
    Log('Serviço de Terminal Server reiniciado.');
    Log('Tipo Bílis! TELORME Reconfigurado');
    Log('Mais fácil que levantar uma ponte pra arrumar.');
    ShowMessage('Alterações aplicadas com sucesso!');

  except
    on E: Exception do
      Log('Erro: ' + E.Message);
  end;
end;

end.
