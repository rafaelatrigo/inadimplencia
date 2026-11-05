# Acompanhamento de inadimplência · Grupo Carozzo

Painel interno de inadimplência (evolução semanal, % sobre o recebimento esperado, aging, top 10, recuperação de crédito e Tratados). Fica publicado no **GitHub Pages** e é protegido por **senha**. Segue o mesmo modelo do painel de despesas.

- **Senha:** fica em `scripts/senha.txt`, um arquivo local que não vai para o GitHub. Precisa ter pelo menos 12 caracteres. Para trocar a senha, edite o arquivo e rode `atualizar_painel.bat`.
- **Dados de clientes:** nunca vão para o GitHub sem criptografia. Os relatórios do Sienge (`entrada/`) e o histórico (`dados/`) ficam só no seu computador. No GitHub vai apenas `site/data/painel.enc.json`, que é criptografado.

## De onde vêm os dados: duas bases, lado a lado

O painel monta **duas bases**. Você escolhe qual ver em **Base**, no menu lateral. O card **Conferência Sienge × planilha** mostra as diferenças título a título.

| Base | De onde vem |
|---|---|
| **Automática (Sienge)** | CSV do Contas a Receber em `entrada`, com a regra dos 10 dias e sem os Tratados |
| **Planilha** | Planilha `Inadimplência_Carozzo_DDMMAAAA.xlsx` com a data mais recente no nome, na pasta do OneDrive `2. Carozzo\INADIMPLENCIA` (abas Base Atual e Base Anterior) |

As informações abaixo valem para as duas bases:
- **Tratados:** aba Tratados da planilha mais os classificados pelo painel (o do painel vale por cima).
- **Sub judice:** PDF de Títulos em Cobrança, ou "Sim" na coluna AA da planilha.
- **Status da cobrança:** o lançamento mais recente entre o PDF e a Base Posição Clientes.
- **Variação Semanal:** histórico da planilha.

Quando alguém salva a planilha, o OneDrive sincroniza, o serviço percebe e publica sozinho.

## Rotina semanal (modo relatórios do Sienge)

1. No Sienge, exporte os dois relatórios e salve na pasta **`entrada`**:
   - **Contas a Receber** em **CSV**:
     - parâmetros: dez/25 até D-0;
     - flags: Parcelas SPC/Serasa, Inadimplentes, Sub judice e Parcelas de resíduo;
     - plano financeiro: 1.01.01.02.
   - **Relatório de Títulos em Cobrança** em **PDF**:
     - flags: Em cobrança, Inadimplentes e Sub judice;
     - situação: Todas.
2. Pronto. O serviço percebe os arquivos novos e publica sozinho em cerca de 2 minutos. Para publicar na hora, use **Publicar agora** no painel.
3. No painel, em **Títulos inadimplentes**, marque o que precisa sair da base (distrato, pagamento em trânsito, renegociação…) e clique em **Tratar selecionados**.

O script sempre usa o arquivo **mais recente** de cada tipo. Os arquivos antigos podem ficar na pasta.

## Rotina mensal

- **Recebimento esperado:** salve os PDFs **Contas a Receber (por Data de Vencimento)** que você envia ao BPO em **`entrada/boletagem`** (um por empresa e mês; subpastas valem). O painel usa a coluna **Total** do "Total da empresa".
  - Clientes em `esperado_excluir_clientes` no `config.json` saem desse total. Hoje é o JDL no Feira.
  - Se houver mais de um PDF da mesma empresa e mês (v2, v3…), vale o mais recente.
- **Fechamento para a Rita:** no card **% Inadimplência**, escolha o mês de referência e clique em **Registrar fechamento**. Os números ficam congelados em "Fechamentos registrados".

## Regras de cálculo

| Item | Regra |
|---|---|
| **Base** | Títulos vencidos até a data do relatório, sem os Tratados. O cliente entra se tiver **algum título com mais de 10 dias** de atraso, e aí entram **todos** os títulos vencidos dele (`regra_dias` no config). |
| **Dias ajustado** | O maior atraso do cliente. É ele que define a faixa do aging e o "acima de 90 dias". |
| **Sub judice** | Situação "Sub judice" no relatório de Títulos em Cobrança, por documento. Fica fora das faixas do aging e do "acima de 90 dias". |
| **Status da cobrança** | O último lançamento do cliente no relatório de Títulos em Cobrança. Lançamentos com data futura não contam e aparecem em Avisos, para corrigir no Sienge. |
| **Semana anterior** | O retrato salvo pelo painel na semana anterior (`dados/retratos`). |
| **Recuperado** | Estava na semana anterior, saiu da base e **não** foi tratado. |
| **Tratados** | Distrato vai para a caixinha Distrato. Os outros status ficam como **aguardando compensação** enquanto o título está no Contas a Receber do Sienge, e passam a **compensado** quando saem dele. |
| **% acumulado** | Saldo inadimplente da base ÷ recebimento esperado somado do mês do título vencido mais antigo de cada empreendimento até o mês de referência. Meses sem esperado registrado aparecem em aviso. |
| **% do mês** | Títulos com vencimento no mês ÷ recebimento esperado do mês. |
| **Chave do título** | Documento + parcela, sem as marcas (C), (S) e (*) do Sienge. A parcela é sempre lida como texto, e as parcelas que o Excel transformou em data voltam ao formato original. |

## Instalação (uma vez)

1. Instale as bibliotecas: `python -m pip install pdfplumber cryptography`
2. No GitHub, crie o repositório **rafaelatrigo/inadimplencia**. Depois, em **Settings > Pages > Source**, escolha **GitHub Actions**.
3. Na pasta do painel, abra o terminal e rode:
   ```
   git init -b main
   git remote add origin https://github.com/rafaelatrigo/inadimplencia.git
   ```
   Depois, dê duplo clique em `atualizar_painel.bat` para fazer a primeira publicação.
4. Dê duplo clique em `instalar_servico.bat`. O serviço inicia junto com o Windows e usa a porta 8733 (o de despesas usa a 8732; os dois convivem).
5. Abra **https://rafaelatrigo.github.io/inadimplencia/**. Na primeira vez, o navegador pode perguntar se o site pode "acessar dispositivos na rede local". Clique em **Permitir**. É assim que o painel conversa com o serviço.

Para parar e remover o serviço, use `desinstalar_servico.bat`. Ele não afeta o serviço do painel de despesas.

### Histórico inicial

A pasta `dados/` já vem com o histórico importado da planilha *Inadimplência Carozzo 05112026*:
- os 52 Tratados;
- a Variação Semanal desde jul/25;
- a Base Anterior de 29/10;
- o esperado de mar a out/26, com o Feira sem o JDL.

Para importar de novo, use `scripts/importar_planilha.py`. Ele não sobrescreve Tratados sem `--forcar`.

## Pastas (scripts/config.json)

| Chave | Padrão | O que é |
|---|---|---|
| `pasta_entrada` | `entrada` | Onde salvar o CSV do Contas a Receber e o PDF de Títulos em Cobrança |
| `pasta_boletagem` | `entrada/boletagem` | PDFs enviados ao BPO. Pode apontar para a pasta do OneDrive onde eles já ficam. |
| `pasta_dados` | `dados` | Tratados, fechamentos, retratos semanais e esperado manual. **Recomendado:** apontar para uma pasta do OneDrive, para ter backup. |

Os caminhos podem ser absolutos (ex.: `C:\\Users\\...\\OneDrive - Carozzo\\...`) ou relativos à pasta do painel.
