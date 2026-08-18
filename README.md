# projeto-planilha-influenciadores-betesporte
# Projeto Planilha Influenciadores Betesporte

Sistema de gestão, auditoria, relatório e cobrança de postagens diárias de influenciadores, construído em **Google Apps Script (GAS)** integrado a uma planilha do **Google Sheets**.

## Funcionalidades

- Painel lateral (sidebar) para busca de influenciador e registro de status com avanço automático
- Dashboard de desempenho com KPIs e gráfico de distribuição de formatos
- Ranking de assiduidade (melhores e piores)
- Relatório diário por e-mail
- Relatórios por período, mês e por influenciador + data
- Checagem de inadimplentes (quem não postou hoje)
- Cobrança individual formatada para WhatsApp
- Exportação de relatórios em PDF (salvo no Google Drive)
- Arquivo do dia no histórico (aba HISTORICO)

## Estrutura de arquivos

| Arquivo | Responsabilidade |
|---------|------------------|
| `Menu.gs` | Cria o menu do painel e o `onEdit` (despacho por aba) |
| `Code.gs` | Linha selecionada, geração de PDF, formatação de status |
| `Utils.gs` | Abas, ordenação, dashboard, salvar status, busca |
| `Setup.gs` | Criação das abas (ACOMPANHAMENTO, BANCO_DE_DADOS, RANKING) |
| `SidebarBackend.gs` | Backend da sidebar: navegação, busca, histórico, ranking |
| `Ranking.gs` | Cálculo do ranking de assiduidade |
| `RelatorioDiario.gs` | Envio do relatório diário por e-mail |
| `Historico.gs` | Arquiva status no histórico e normaliza datas/textos |
| `Verificacao.gs` | Módulo de verificação (em desenvolvimento) |
| `Sidebar.html` | Interface lateral (busca, status, navegação) |
| `FiltroDataModal.html` | Modal de relatório por período e formato |
| `FiltroMesModal.html` | Modal de relatório mensal |
| `FiltroInfluenciadorDataModal.html` | Modal de consulta por influenciador + período |
| `CobrancaSemanalModal` | Modal de cobrança semanal |

## Abas da planilha

- **ACOMPANHAMENTO** — lista de influenciadores e status do dia (data na linha 2)
- **BANCO_DE_DADOS** — histórico consolidado (DATA, INFLUENCIADOR, @USERNAME, STATUS)
- **HISTORICO** — arquivo diário de status
- **RANKING** — ranking de assiduidade
- **Ranking Assiduidade** — relatório detalhado de assiduidade

## Instalação

1. Abra a planilha no Google Sheets.
2. Acesse **Extensões → Apps Script**.
3. Crie os arquivos conforme a estrutura acima e cole os códigos.
4. Execute `onOpen` uma vez (ou recarregue a planilha) para criar o menu **⚽ BETEsporte - Painel**.
5. Na primeira execução, autorize o script (fluxo "Avançado → Ir para Projeto sem título (não seguro)").

## Uso

- Use o menu **⚽ BETEsporte - Painel** para todas as ações.
- Na sidebar, selecione um influenciador, escolha o status e o sistema salva e avança automaticamente (1,2s).

## Manutenção

- **onEdit**: mantenha um único handler em `Menu.gs`, despachando por aba.
- **Normalização**: use `normalizarData` e `normalizarTexto` (em `Historico.gs`) para padronizar datas e textos.
- **Ranking**: mantenha a lógica centralizada em `Ranking.gs` (evite duplicar em outros arquivos).
