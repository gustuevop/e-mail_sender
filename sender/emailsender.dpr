program emailsender;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient,
  IdSMTPBase, IdSMTP, Vcl.StdCtrls, idMessage, idSSLOpenSSL;

  var
    t: Boolean;
    LSMTP : TIdSMTP;
    LMessage : TIdMessage;
    LSocketSSL : TIdSSLIOHandlerSocketOpenSSL;

begin
  try
    { TODO -oUser -cConsole Main : Insert code here }

        begin
          LSMTP := TIdSMTP.Create(nil);
          LMessage := TIdMessage.Create(nil);
          LSocketSSL := TIdSSLIOHandlerSocketOpenSSL.Create(nil);

          with LSocketSSL do
          begin
            with SSLOptions do
            begin
              Mode := sslmClient;
              Method := sslvTLSv1_2;
            end;
            Host := 'smtp.office365.com';
            Port := 587;
          end;

          with LSMTP do
          begin
            IOHandler := LSocketSSL;
            Host := 'smtp.office365.com';
            Port := 587;
            AuthType := satDefault;
            Username := '';
            Password := '';
            UseTLS := utUseExplicitTLS;
          end;

          with LMessage do
          begin
            From.Address := '';
            From.Name := '';
            Recipients.Add;
            Recipients.Items[0].Address := '';
            Subject := '';
            Body.Add('');
          end;

          try
            LSMTP.Connect;
            LSMTP.Send(LMessage);
            ShowMessage('Envio concluído');
          Except
            ON E: Exception do
            ShowMessage(E.Message);
          end;
        end;

      Read;

  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
