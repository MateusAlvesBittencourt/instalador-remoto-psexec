# Instalador Remoto via PsExec — versão Python

Interface para Windows feita em Python com `tkinter`. O script `RemoteInstallerParallel.ps1` controla a fila, a cópia direta, o PsExec e a instalação remota.

A interface possui layout corporativo responsivo, painel de andamento com destaque por cores e status permanente da execução. Em janelas menores, os formulários e botões são reorganizados automaticamente; a área principal e a tabela possuem barras de rolagem.

## Forma mais rápida de usar

1. Instale Python 3 para Windows e marque **Add Python to PATH**.
2. Extraia todos os arquivos do ZIP.
3. Execute `executar.bat`.
4. Aceite a solicitação de administrador do Windows.

Não é necessário Visual Studio, `cl.exe` nem compilador C.

## Gerar um EXE

Execute `gerar-exe.bat`. Ele instala o PyInstaller, se necessário, e cria:

```text
dist\InstaladorRemoto.exe
```

## Uso

1. Selecione a pasta completa que contém o pacote na unidade `J:`. A ferramenta converte automaticamente `J:` para `\\POA01FLS05\Software$`.
2. Selecione dentro dela o instalador `.exe` ou `.msi`.
3. Informe os parâmetros silenciosos. Para MSI, o padrão é `/qn /norestart`. Ao trocar de MSI para EXE, os parâmetros padrão de MSI são removidos automaticamente; informe apenas os parâmetros próprios do instalador EXE, se necessários.
4. Selecione a lista TXT ou CSV de patrimônios. O formato esperado é `W` seguido por números.
5. O campo do PsExec vem preenchido com `C:\Windows\System32\PsExec64.exe`. Use **Procurar** apenas se ele estiver em outro local.
6. Informe o nome da pasta que será criada dentro de `C:\Temp`.
7. Escolha entre 1 e 5 instalações simultâneas. O padrão recomendado é 4.
8. Escolha a validação: código de saída, arquivo instalado ou serviço do Windows.
9. Teste o acesso às máquinas antes de iniciar a instalação.
10. Use **Copiar relatório** para copiar a tabela de resultados. O conteúdo é tabulado e pode ser colado diretamente no Excel, e-mail ou Bloco de Notas.

As mensagens de conexão gravadas pelo PsExec no canal de erro são capturadas sem interromper a operação. Depois da tentativa de cópia, a ferramenta verifica diretamente se o instalador EXE ou MSI existe e possui conteúdo no destino. Quando o arquivo é confirmado, a instalação continua mesmo que o PsExec ou o Robocopy tenha retornado um código inesperado.

## Validação após instalar

- **Código de saída:** aceita os códigos de sucesso do instalador.
- **Arquivo instalado:** informe um caminho como `C:\Program Files\Fabricante\Aplicativo.exe`.
- **Serviço do Windows:** informe o nome interno do serviço, não o nome de exibição.

Quando a validação falha, a pasta permanece em `C:\Temp` para diagnóstico.

## Fila e reexecução

A ferramenta executa até quatro máquinas em paralelo por padrão e inicia as demais automaticamente conforme surgem vagas. A tabela mostra acesso, cópia, instalação, validação, limpeza, tempo e detalhes. Depois da conclusão, **Reexecutar falhas** cria uma nova fila somente com as máquinas que falharam.

## Requisitos de rede

- A conta atual deve ser administradora local dos computadores remotos.
- O compartilhamento administrativo `\\PATRIMONIO\C$` deve estar acessível.
- PsExec deve ser permitido pela política de segurança e pelo antivírus.
- O Robocopy é executado pelo PsExec dentro de cada máquina como SYSTEM. Os dados seguem diretamente de `\\POA01FLS05\Software$` para `C:\Temp`, sem passar pela máquina do técnico.

## Instalações

EXE:

```text
"C:\Temp\NomeDoPacote\setup.exe" <parâmetros informados>
```

MSI:

```text
msiexec.exe /i "C:\Temp\NomeDoPacote\aplicativo.msi" /qn /norestart
```

Os resultados permanecem na tabela da interface e podem ser copiados pelo botão “Copiar relatório”. Nenhum relatório CSV é gerado automaticamente após a execução. Códigos `0`, `1641` e `3010` são tratados como sucesso; os dois últimos indicam necessidade de reinicialização.

Depois que o instalador retorna sucesso, a pasta copiada é excluída de `C:\Temp`. Se a instalação falhar, os arquivos são mantidos para diagnóstico. O resultado aparece na coluna `Limpeza` da tabela.
