# Painel de Propostas — Vertis · BMM · Mannare

Painel em um único arquivo (`index.html`), código aberto e sem dependências. Abre em qualquer navegador (Chrome, Edge, Safari), no computador ou no celular, sem conta no Claude. Logos, cores e a fonte Urbanist já estão embutidos no arquivo, então funciona até sem internet (exceto a sincronização).

## Como funciona

**1. Elaborar** — Em *+ Nova proposta*, preencha os dados do cliente (razão social e CNPJ). Na aba **Proposta**, escreva o objeto, os itens com valores e ajuste as cláusulas. Use `**negrito**` e `__sublinhado__` no texto. Escolha quem assina pela contratada.

**2. Enviar para aprovação** — Na aba **Aprovação**, clique em *Enviar para aprovação*. Os outros sócios recebem um e-mail de aviso (na versão compartilhada) e a proposta aparece para eles com a marca “seu voto”.

**3. Aprovar** — Cada sócio abre a proposta, clica em *Ver a proposta*, e depois em *Aprovar* ou *Solicitar ajustes*, confirmando com o **PIN pessoal**.
- Um pedido de ajuste devolve a proposta para elaboração. Quem elaborou corrige e reenvia.
- Se alguém alterar o conteúdo depois do envio, as aprovações são anuladas automaticamente e é preciso aprovar de novo.

**4. Gerar o PDF** — Quando todos aprovam, o status vira *Aprovada p/ envio*. Clique em *Gerar PDF final* e, na janela de impressão, escolha **Salvar como PDF**. O PDF sai no papel timbrado **Layout B**: faixa com BMM, Vertis e Mannare no topo de todas as páginas, rodapé com endereço, telefones e número da página, seções numeradas, tabela de valores, bloco de assinaturas e “De acordo”.
- Antes da aprovação, a pré-visualização sai com a marca d'água **MINUTA**.
- No Chrome/Edge, em “Mais definições”, desmarque *Cabeçalhos e rodapés*. O número da página é colocado pela própria proposta. Use Chrome ou Edge para gerar o PDF: o Firefox não imprime o rodapé.

**5. Parcelas e pagamento** — Na aba **Proposta**, defina as parcelas: percentual e condição, por exemplo "no aceite da proposta" e "na conclusão do processo".
- O texto da forma de pagamento é gerado sozinho a partir das parcelas.
- Na aba **Financeiro** da proposta, informe o serviço (nome curto), a captação (quem trouxe o cliente), a competência (quem emite a NF), o imposto, a previsão de cada parcela e eventuais desembolsos.
- A divisão do líquido segue a regra da captação. Em caso de exceção, marque "Exceção" e informe o motivo. A captação e a divisão entram na aprovação dos sócios.

**6. Financeiro** — Quando a proposta é marcada como *Aceita pelo cliente*, cada parcela vira um **recebível** na visão **Financeiro**, já com valor, previsão e divisão.
- Os status são: *A faturar → NF emitida → NF enviada → NF atrasada → Pago* (ou *Cancelada*).
- Clique numa parcela para registrar o número e a data da NF e o pagamento. O botão "Marcar pago hoje" faz isso de uma vez.
- O painel mostra: quanto falta faturar, faturado em aberto, atrasado (previsão vencida), recebido no mês, previsto para 30 dias, o líquido a receber por sócio e o **repasse por mês** (BMM, Mannare e Rafael).
- *+ Recebível avulso* serve para cobranças sem proposta. A opção "repetir mensalmente" é para contratos como o consultivo mensal.
- *Exportar (.csv)* gera a planilha para o contador.
- Cálculo: líquido = bruto − imposto − desembolso. A parte de cada sócio é o líquido × o percentual da divisão.

**7. Contas a pagar** — Fica em Financeiro › *Contas a pagar*. Serve para as despesas da operação: contador, plataformas, telefone, veículos, salários etc.
- Cada conta tem categoria, centro de custo (BMM/Mannare), vencimento e valor.
- O status muda sozinho pelo vencimento: *A pagar → Vencendo* (até 3 dias) *→ Atrasado*.
- Marque **Recorrente** para contas mensais: ao marcar como paga, a do mês seguinte é lançada automaticamente. Para lançar vários meses de uma vez, use "Lançar N vezes".
- Marque **Reembolso** quando um sócio pagou do próprio bolso e informe quem pagou. O painel mostra o reembolso individual (metade do valor), como no ClickUp.

**8. Fornecedores** — Fica em Financeiro › *Fornecedores*. Registra os pedidos a terceiros (mapas, georreferenciamento, topografia…): fornecedor, contato, empreendimento, parcela, data do pedido, vencimento e valor.
- Os status são: *Pedido recebido → Próximo ao vencimento → Vencido → Pago*, e também *Agendado*.
- Vincule o pedido à **proposta** e à **parcela a receber** do serviço e marque "Descontar como desembolso". O valor entra no desembolso daquela parcela e reduz o líquido a dividir entre os sócios.
- Na ficha da proposta (aba Financeiro), os fornecedores vinculados aparecem com o total.

**9. Fluxo de caixa** — Mostra, mês a mês (3 meses atrás até 6 à frente): recebido, a receber, pago, a pagar, o saldo do mês e o acumulado.
- Dá para filtrar por empresa.
- Entradas usam o valor bruto dos recebíveis (sem impostos). Saídas são as contas a pagar mais os fornecedores.

**10. Acompanhar** — Só propostas aprovadas podem ser marcadas como *Enviada ao cliente*, *Em negociação* ou *Aceita pelo cliente*. A assinatura com o cliente continua no seu fluxo atual (por exemplo, a assinatura eletrônica do Google Docs).

## Instalação compartilhada (Google Sheets) — uma vez, ~15 min

1. Crie uma Planilha Google nova (ex.: “Propostas – Banco de dados”).
2. Abra **Extensões → Apps Script**, apague o conteúdo e cole todo o `Code.gs`. Salve.
3. Vá em **⚙ Configurações do projeto → Propriedades do script** e adicione:
   - `TOKEN` = uma senha comum de acesso ao painel;
   - `PINS` = o PIN de cada sócio, com o nome **exatamente** como cadastrado no painel. Por exemplo:
     `{"Letícia Libério":"4821","Bruno Montenegro":"7390","Rafael Lucas Russo Daniel":"1155"}`
   - Cada sócio deve escolher o próprio PIN e informá-lo só a quem administra a planilha.
4. Clique em **Implantar → Nova implantação → App da Web**:
   - Executar como: **Eu**;
   - Quem pode acessar: **Qualquer pessoa**.
   - Autorize o acesso à planilha e ao envio de e-mails.
5. Copie a URL que termina em `/exec`.
6. Abra o `index.html`, vá em **⚙ Configurações**, preencha a URL e o TOKEN, escolha quem é você e salve.
7. Envie a cada sócio o `index.html`, a URL e o TOKEN. Cada um faz o passo 6 no próprio computador.

Para ter um link fixo (inclusive no celular), publique a pasta gratuitamente no GitHub Pages ou no Netlify Drop. O arquivo não contém dados: eles ficam na planilha e só são acessíveis com o TOKEN.

## Segurança
- **O que o servidor garante:** ninguém registra aprovação sem o PIN correto; quem elaborou não aprova a própria proposta; um voto só vale para a versão que o sócio revisou.
- **Planilha:** não edite manualmente a coluna `json` da planilha.
- **Modo local** (sem planilha): serve só para testar. Não pede PIN e os dados ficam apenas naquele navegador.

## Personalização
- **Centros de custo e categorias de despesa:** em ⚙ Configurações › Financeiro.
- **Regra de divisão e imposto:** em ⚙ Configurações › Financeiro. O padrão é:
  - captação Mannare: Mannare 66,5%, BMM 28,5%, Rafael 5%;
  - captação BMM: BMM 47,5%, Mannare 47,5%, Rafael 5%;
  - imposto: 12,5%.
- **Sócios e empresas:** em ⚙ Configurações. A lista de sócios define quem precisa aprovar (todos, menos quem elaborou).
- **Cláusulas padrão:** ajuste numa proposta e clique em *Salvar cláusulas como padrão*.
- **Papel timbrado:** o modelo original está em `marca/papel-timbrado-layout-B.pdf`. Dele saem três recortes:
  - `cabecalho-proposta.jpg` (faixa do topo);
  - `rodape-texto.jpg` (endereço e telefones, impresso no rodapé);
  - `rodape-tela.jpg` (rodapé na pré-visualização).

  O endereço e os telefones fazem parte da arte. Para mudá-los, atualize o timbrado e peça para regerar o `index.html`, ou substitua as constantes `HEADER_IMG`, `FOOTER_TXT` e `FOOTER_SCREEN` no arquivo.
- **Código:** as regras de aprovação ficam no bloco “Regras de aprovação”, idêntico no `index.html` e no `Code.gs`. Se alterar o `Code.gs`, use **Implantar → Gerenciar implantações → editar → Nova versão**.

## Atualizando de uma versão anterior
1. Substitua o `index.html` em todos os computadores (ou no link publicado).
2. No Apps Script, apague o código antigo, cole o `Code.gs` novo e **salve**.
3. Ainda no editor, escolha a função **configurarPlanilha** no menu ao lado de "Depurar" e clique em **Executar**. Isso cria as abas que faltam (Recebiveis, Pagamentos…) e pede autorização, se precisar.
4. **Publique a nova versão:** *Implantar → Gerenciar implantações → ✎ editar → Versão: **Nova versão** → Implantar*. Sem esse passo, o link `/exec` continua rodando o código antigo, mesmo com o novo salvo. Não use "Nova implantação": ela gera outro link.
5. No painel, clique em ⟳ Atualizar. Se aparecer o aviso vermelho de "versão antiga do código", o passo 4 não foi concluído.

## Limite
Cada proposta (e cada recebível) ocupa uma linha da planilha, com os dados completos em uma célula (cerca de 50 mil caracteres). Isso é suficiente para propostas como a do modelo, mesmo com bastante discussão nos comentários.
