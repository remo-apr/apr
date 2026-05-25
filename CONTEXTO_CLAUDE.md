# CONTEXTO DO PROJETO — APR REMO ENGENHARIA
## Documento de Handoff para Nova Sessão Claude
> **Leia este arquivo inteiro antes de qualquer ação. Ele representa o estado atual completo do projeto.**

---

## 1. IDENTIDADE DO PROJETO

| Campo | Valor |
|---|---|
| **Nome** | Sistema APR Digital — REMO ENGENHARIA |
| **Obra** | SE JAGUARA 345 kV |
| **Arquivo principal** | `C:\Users\Remo Engenharia\Desktop\apr\index.html` |
| **URL pública (GitHub Pages)** | https://remo-apr.github.io/apr/ |
| **Repositório** | https://github.com/remo-apr/apr |
| **Branch** | `main` |
| **Versão atual** | `v4.1` (variável `APP_VERSAO = 'v4.1'`) |
| **Tamanho do arquivo** | ~430 KB, 4421 linhas |
| **Conformidade** | NR-10 / NR-18 / NR-35 |

### O que é o sistema
Single-file HTML com PWA (Progressive Web App) para elaboração de Análises Preliminares de Risco (APRs) diárias antes de serviços elétricos, civis e de montagem em subestação de alta tensão. Funciona sem servidor, sem banco de dados externo, sem instalação. Integrado ao Google Drive e Google Sheets via Google Apps Script.

---

## 2. INFRAESTRUTURA GOOGLE

### Apps Script
| Campo | Valor |
|---|---|
| **Versão implantada** | v12 (20/05/2026, 14:48) |
| **Project ID** | `1lPJsEOhqHCzdiA5d5WuJbdrmEkfA4ZuQZpVVKDraU7a3d0ArWzs0lj-_` |
| **URL do editor** | https://script.google.com/home/projects/1lPJsEOhqHCzdiA5d5WuJbdrmEkfA4ZuQZpVVKDraU7a3d0ArWzs0lj-_ /edit |
| **Web App URL (endpoint)** | `https://script.google.com/macros/s/AKfycbwPdWSS-Abki7PxArO-cfll3P-mOunnW2IKXpdelflb20W0jiKwIdUjqKqFZGqBA0ah/exec` |
| **executeAs** | `USER_DEPLOYING` |
| **access** | `ANYONE_ANONYMOUS` |
| **OAuth Scopes autorizados** | `drive` + `spreadsheets` (APENAS esses dois — ver seção de erros conhecidos) |

### Google Drive
| Campo | Valor |
|---|---|
| **Pasta da obra** | SE JAGUARA 345 kV |
| **Pasta ID** | `114YydL5vkN9mkJkZGMeO4-HF9LGwPpfp` |

### Google Sheets
| Campo | Valor |
|---|---|
| **Nome da planilha** | `Registro de APRs - REMO JAGUARA 345 kV` |
| **Aba** | `Registro APR` |
| **SS_ID armazenado em** | `PropertiesService.getScriptProperties()` key `SS_ID` |
| **Colunas (13)** | Rastreio \| Data \| Atividade \| Elaborador \| Encarregado \| Equipe \| Arquivo Drive \| Etapa \| Clima \| Local \| PT \| Versão \| Registrado em |

---

## 3. CÓDIGO COMPLETO DO APPS SCRIPT (v12 — ESTADO ATUAL)

```javascript
// APR REMO — Apps Script v12
// PDF gerado pelo browser (html2pdf) — salvo diretamente no Drive
const FOLDER_ID  = '114YydL5vkN9mkJkZGMeO4-HF9LGwPpfp';
const SHEET_NAME = 'Registro APR';

function doPost(e) {
  try {
    const dados    = JSON.parse(e.postData.contents);
    const pasta    = DriveApp.getFolderById(FOLDER_ID);
    const rastreio = dados.rastreio || ('APR-' + Date.now());
    let linkArquivo = '';
    try {
      if(dados.tipo === 'blob' && dados.pdfBase64){
        // PDF gerado pelo browser via html2pdf — salvar diretamente
        const pdfBytes = Utilities.base64Decode(dados.pdfBase64);
        const pdfBlob  = Utilities.newBlob(pdfBytes, 'application/pdf', rastreio + '.pdf');
        const pdfFile  = pasta.createFile(pdfBlob);
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkArquivo = pdfFile.getUrl();
        console.log('PDF (browser) salvo: ' + rastreio);
      } else {
        // Fallback: gerar PDF via Sheets (legado)
        const pdfFile = gerarPDF(dados, rastreio, pasta);
        linkArquivo   = pdfFile.getUrl();
        console.log('PDF (sheets) salvo: ' + rastreio);
      }
    } catch(pdfErr) {
      console.log('PDF falhou: ' + pdfErr.message);
    }
    registrarPlanilha(dados, linkArquivo);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, link: linkArquivo }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    console.log('ERRO doPost: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function registrarPlanilha(d, link) {
  const ss    = obterPlanilha();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  if(sheet.getLastRow() === 0){
    sheet.appendRow(['Rastreio','Data','Atividade','Elaborador','Encarregado','Equipe','Arquivo Drive','Etapa','Clima','Local','PT','Versão','Registrado em']);
    sheet.getRange(1,1,1,13).setFontWeight('bold').setBackground('#1B3E6B').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  const agora = new Date();
  const row = [
    d.rastreio||'',      // col 1
    d.data||'',          // col 2
    d.atividade||'',     // col 3
    d.elaborador||'',    // col 4
    d.encarregado||'',   // col 5
    d.equipe||'',        // col 6
    '',                  // col 7: PDF link (set below via richText)
    d.etapa||'',         // col 8
    d.clima||'',         // col 9
    d.local||'',         // col 10
    d.pt||'',            // col 11
    d.versao||'',        // col 12
    Utilities.formatDate(agora,'America/Sao_Paulo','dd/MM/yyyy HH:mm'), // col 13
  ];
  sheet.appendRow(row);
  if(link){
    const lastRow  = sheet.getLastRow();
    const richText = SpreadsheetApp.newRichTextValue()
      .setText('📄 Abrir PDF')
      .setLinkUrl(0, '📄 Abrir PDF'.length, link)
      .build();
    sheet.getRange(lastRow, 7).setRichTextValue(richText);
  }
}

function obterPlanilha() {
  const props = PropertiesService.getScriptProperties();
  let ssId    = props.getProperty('SS_ID');
  if(ssId){ try { return SpreadsheetApp.openById(ssId); } catch(e){ ssId = null; } }
  const pasta = DriveApp.getFolderById(FOLDER_ID);
  const files = pasta.getFilesByName('Registro de APRs - REMO JAGUARA 345 kV');
  if(files.hasNext()){
    const f = files.next();
    props.setProperty('SS_ID', f.getId());
    return SpreadsheetApp.openById(f.getId());
  }
  const novo = SpreadsheetApp.create('APR REMO - Registros');
  const file = DriveApp.getFileById(novo.getId());
  DriveApp.getFolderById(FOLDER_ID).addFile(file);
  try{ DriveApp.getRootFolder().removeFile(file); }catch(e){}
  props.setProperty('SS_ID', novo.getId());
  return novo;
}

function testeConexao() {
  const pasta = DriveApp.getFolderById(FOLDER_ID);
  const ss    = obterPlanilha();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  console.log('Pasta OK: ' + pasta.getName());
  console.log('Planilha OK: ' + ss.getName() + ' / aba: ' + sheet.getName());
  console.log('Colunas: ' + sheet.getRange(1,1,1,13).getValues()[0].join(' | '));
  console.log('=== TUDO OK ===');
}
```

> **IMPORTANTE:** Toda vez que o código do Apps Script for alterado, é necessário implantar como **Nova versão** em "Gerenciar implantações" → lápis → dropdown "Nova versão" → Implantar. A URL do endpoint NÃO muda entre versões.

---

## 4. FLUXO DE SINCRONIZAÇÃO (index.html → Drive)

### Função crítica: `autoSyncComPDF()` (linha ~4217 do index.html)

```javascript
function autoSyncComPDF(url, dados){
  if(!url) return;
  const badge = document.getElementById('sync-status');
  if(badge){ badge.style.display='inline'; badge.className='sync-badge ing'; badge.textContent='🔄 Gerando PDF…'; }

  carregarHtml2Pdf(function(erro){
    if(erro || typeof html2pdf === 'undefined'){
      autoSync(url, dados); // fallback
      return;
    }
    const el = document.getElementById('apr-out');
    if(!el){ autoSync(url, dados); return; }

    const opt = {
      margin:     [6, 6, 6, 6],
      filename:   (dados.rastreio||'APR') + '.pdf',
      image:      { type:'jpeg', quality:0.90 },
      html2canvas:{ scale:1.5, useCORS:true, logging:false, scrollY:0 },
      jsPDF:      { unit:'mm', format:'a4', orientation:'portrait' },
      pagebreak:  { mode:['css','legacy'] }
    };

    html2pdf().set(opt).from(el).outputPdf('blob')
    .then(function(blob){
      const reader = new FileReader();
      reader.onload = function(){
        const base64  = reader.result.split(',')[1];
        const payload = {
          rastreio: dados.rastreio, data: dados.data, atividade: dados.atividade,
          elaborador: dados.elaborador, encarregado: dados.encarregado, equipe: dados.equipe,
          etapa: dados.etapa, clima: dados.clima, local: dados.local, pt: dados.pt,
          versao: dados.versao, obs: dados.obs,
          tipo: 'blob', pdfBase64: base64
        };
        fetch(url, {
          method:'POST',
          headers:{'Content-Type':'text/plain;charset=utf-8'},
          body: JSON.stringify(payload),
          mode:'no-cors'
        })
        .then(()=>{ if(badge){ badge.className='sync-badge ok'; badge.textContent='✅ Registrado no Drive'; } })
        .catch(()=>{ if(badge){ badge.className='sync-badge fail'; badge.textContent='❌ Falha no envio'; } });
      };
      reader.readAsDataURL(blob);
    }).catch(function(){ autoSync(url, dados); });
  });
}
```

### Pontos críticos do fetch:
- **`mode: 'no-cors'`** é OBRIGATÓRIO — Apps Script bloqueia CORS
- **`Content-Type: 'text/plain;charset=utf-8'`** é OBRIGATÓRIO com `no-cors` (não usar `application/json`)
- Como `no-cors` não retorna body, a resposta do Apps Script não é lida pelo browser — isso é normal
- O URL do endpoint é configurado pelo usuário em `localStorage` key `apr_sync_url`

---

## 5. ESTRUTURA DO `index.html`

### Constantes globais (linhas 1517–1787)
```
APP_VERSAO = 'v4.1'
SYNC_URL_KEY = 'apr_sync_url'           // chave localStorage para URL do Apps Script
ENCARREGADOS[]                           // 5 encarregados com nome, funcao, disc
MAQUINAS[]                               // 33 máquinas (strings)
CHECKLISTS{}                             // objeto chave=nome_maquina → array de itens
LOCAIS_FIXOS[]                           // 14 locais da SE
CREDENCIADOS[]                           // 7 usuários com user/senha/cargo/email
EQUIPE_CAMPO[]                           // 46 colaboradores com num/nome/cargo/equipe
EPIS_CADASTRO[]                          // 20 EPIs cadastrados
RISCOS[]                                 // 35 riscos com n/p/s/r/c
ATIVIDADES[]                             // 32 atividades com id/nome/disc/normas/fases[]
```

### Estado global (Sets/Maps)
```javascript
SEL         // Set — ids de atividade selecionada (máx 1 elemento via radio)
SEL_L       // Set — locais selecionados
SEL_E       // Set — encarregados selecionados (nomes)
SEL_M       // Set — máquinas selecionadas
SEL_ESC     // Set — escopos selecionados
SEL_EQUIPE  // Set — nums dos colaboradores selecionados
SEL_RISCOS  // Map("atvId_faseIdx" → Map(riscoNum → complemento))
SEL_M_QTD   // Map — quantidades de máquinas
customRiscosFase  // Map — riscos customizados por fase
editRiscosFase    // Map — edições de riscos padrão
riscoPSFase       // Map — riscos de PS customizados
```

### Funções principais
```
ir(step)           // navegar entre etapas (1-5 + 'out')
gerarAPR()         // gera o documento + chama autoSyncComPDF()
nova()             // reseta todos os estados para nova APR
autoSyncComPDF()   // gera PDF via html2pdf.js e envia ao Apps Script
autoSync()         // fallback: envia apenas metadados JSON
carregarHtml2Pdf() // carrega html2pdf.js do CDN sob demanda
salvarSyncUrl()    // salva URL do Apps Script no localStorage
```

### Etapas da UI
```
step-1: Identificação (data, PT, elaborador, encarregado, local, escopo, máquinas, clima, etapa APR)
step-2: Atividades (filtro + seleção por radio, 32 atividades)
step-3: Riscos por Fase (grid por fase da atividade selecionada)
step-4: Equipe de Campo (46 colaboradores, filtro + busca)
step-5: Revisão e Geração (resumo + botão Gerar APR)
step-out: Documento APR gerado (element #apr-out) + botão imprimir
```

### Elemento de saída do documento
```javascript
document.getElementById('apr-out').innerHTML = `...HTML completo da APR...`
```
O `#apr-out` é o que `html2pdf.js` captura para gerar o PDF.

---

## 6. DADOS COMPLETOS DO SISTEMA

### Credenciados (login do sistema)
| user | senha | nome | cargo | aprovador |
|---|---|---|---|---|
| albert | APR@2026 | ALBERT KAIK FRANCISCO DE SOUZA | ENCARREGADO DE TURMA | não |
| andre | APR@2026 | ANDRE JOSE DA SILVA | ENCARREGADO FORCA E CONTROLE | não |
| elizandra | APR@2026 | ELIZANDRA V DE F DOS SANTOS | AUX DE SEGURANCA DO TRABALHO | sim |
| gislaine | APR@2026 | GISLAINE PEREIRA DO NASCIMENTO | TECNICO SEGURANCA DO TRABALHO | sim |
| jhony | APR@2026 | JHONY DOS REIS BORGES | OPERADOR DE MAQUINAS | não |
| jose | APR@2026 | JOSE ADAUTO DOS SANTOS DA SILVA JUNIOR | ENCARREGADO GERAL | não |
| leandro | APR@2026 | LEANDRO SILVA FERREIRA | ENC. MONT. ELETROMECANICA | não |

### Encarregados (seleção na Etapa 1)
- Albert Kaik — Encarregado Civil
- José Adauto — Encarregado Civil
- Leandro Silva — Encarregado de Montagem
- André da Silva — Encarregado de Elétrica
- Marcelo de Oliveira — Supervisor de Elétrica

### Responsáveis Técnicos Fixos (assinam toda APR)
- **Leandro Pizani** — Engenheiro Residente — (31) 99223-7798
- **Gislaine Nascimento** — Técnico de Segurança — (34) 98408-3585

### Equipe de Campo (46 colaboradores)
Civil (aprox. 20), Elétrica (aprox. 8), Montagem (aprox. 7), Administrativo (aprox. 11).
Nomes completos no `index.html` linhas 1669–1716.

### 35 Riscos (resumido)
R01 Cansaço Físico/Mental | R02 Uso de Adorno | R03 Impacto Ambiental |
R04 Ferramentas Inadequadas | R05 Procedimentos Técnicos Incorretos | R06 Picadas de Animais |
R07 Trabalho próx. Circuitos Energizados | R08 Retorno de Tensão | R09 Indução Elétrica |
R10 Trip Acidental | R11 Choque Elétrico | R12 Exposição a Ruído | R13 Exposição à Radiação |
R14 Queda em Altura | R15 Exposição a Vibrações | R16 Fumos Metálicos | R17 Incêndio/Explosões |
R18 Exposição a Poeiras | R19 Levantamento Manual de Peso | R20 Prensamento de Membros |
R21 Colisões/Abalroamento | R22 Iluminação Inadequada | R23 Gases/Vapores |
R24 Atropelamento | R25 Máquinas sem Proteção | R26 Contato com Produtos Químicos |
R27 Manobra Indevida | R28 Escoriações/Cortes/Pancadas | R29 Projeção de Partículas |
R30 Falha de Comunicação | R31 Condições Meteorológicas | R32 Desmoronamento/Soterramento |
R33 Queda de Materiais | R34 Queda de Pessoas | R35 Queimaduras

### 32 Atividades (IDs)
1-Organização e Limpeza | 2-Instalação de Tapume/Cercamento | 3-Serviços de Fundação |
4-Estrutura de Concreto | 5-Alvenaria | 6-Cobertura | 7-Instalações Elétricas |
8-Instalações Hidráulicas | 9-Pavimentação | 10-Jardinagem | 11-Desmobilização |
12-Montagem de Estruturas Metálicas | 13-Içamento de Equipamentos |
14-Instalação de Equipamentos | 15-Comissionamento | 16-Inspeção Visual |
17-Medição e Teste | 18-Manutenção Preventiva | 19-Manutenção Corretiva |
20-Escavação Manual | 21-Escavação Mecânica | 22-Drenagem |
23-Aterramento | 24-Instalação de Cabos | 25-Instalação de Transformador |
26-Instalação de Disjuntores | 27-Instalação de Para-raios |
28-Instalação de Barramentos | 29-Serviços de Topografia |
30-Serviços de Solda | 31-Pintura e Acabamento | 32-Limpeza Final e Entrega
(nomes exatos podem variar — verificar no arquivo)

### 33 Máquinas
Máquina de Solda Elétrica | Furadeira de Bancada | Esmeriladeira | Motoserra |
Oxiacetileno | Rosqueadeira | Makitão | Prensa Elétrica | Lavadora de Alta Pressão |
Serra Mármore | Rompedor Elétrico | Vibrador de Concreto | Serra Circular Bancada |
Prensa Manual | Prensa Hidráulica | Furadeira a Bateria | Furadeira Elétrica |
Parafusadeira Elétrica | Parafusadeira a Bateria | Compactador de Solo | Placa vibratória |
Serra Tico Tico | Serra Circular | Betoneira Elétrica | Oxicorte |
Rompedor Hidráulico | Perfuratriz Pneumática | Perfuratriz Hidráulica |
Compressor Diesel | Compressor Elétrico | Furadeira Martelete | Plaina | Lixadeira

---

## 7. IDENTIDADE VISUAL (CSS)

```css
--navy:  #1B3E6B   /* azul principal REMO */
--green: #00A651   /* verde REMO */
--gold:  #F5C400   /* dourado */
--dk:    #2E5F9E   /* azul secundário */
/* Fontes: Barlow + Barlow Condensed (Google Fonts) */
```

---

## 8. ERROS CONHECIDOS E SOLUÇÕES

### Erro: `DocumentApp.create()` — permissão negada
- **Causa:** O scope `documents` não está autorizado no token do Apps Script
- **Solução:** Nunca usar `DocumentApp`. PDF via `SpreadsheetApp.create()` + `getAs('application/pdf')` OU via blob base64 do browser (v12)

### Erro: `UrlFetchApp.fetch()` — permissão negada
- **Causa:** O scope `script.external_request` não está autorizado
- **Solução:** Nunca usar `UrlFetchApp`. Usar `DriveApp.getFileById().getAs()` para export interno

### Fetch retorna opaque response (sem body)
- **Causa:** `mode:'no-cors'` não permite ler a resposta
- **Solução:** Normal e esperado. Não tentar ler `response.json()`. O Apps Script processa mesmo assim

### Colunas da planilha desalinhadas
- **Causa:** Ordem incorreta no array `row` de `registrarPlanilha()`
- **Solução:** rastreio sempre col 1, timestamp sempre col 13 (ver código v12 acima)

### Monaco editor no Apps Script reverte dropdown de função
- **Causa:** Interface do Google Apps Script tem bug no dropdown ao injetar código
- **Solução:** Sempre clicar no dropdown de funções DEPOIS de injetar código, selecionar função desejada, esperar carregar

### Injetar código no Apps Script via browser automation
```javascript
// Executar no console do Chrome com a aba do Apps Script aberta:
const editors = monaco.editor.getEditors();
editors[0].getModel().setValue(novoCodigoAqui);
```
Depois Ctrl+S para salvar, e implantar como Nova Versão.

---

## 9. COMO ATUALIZAR O SISTEMA

### Atualizar `index.html` e publicar:
```bash
cd "C:\Users\Remo Engenharia\Desktop\apr"
git add index.html
git commit -m "descrição da alteração"
git push origin main
```
GitHub Pages publica automaticamente em 1-2 min em https://remo-apr.github.io/apr/

### Atualizar Apps Script:
1. Abrir https://script.google.com/home/projects/1lPJsEOhqHCzdiA5d5WuJbdrmEkfA4ZuQZpVVKDraU7a3d0ArWzs0lj-_/edit
2. Injetar código via Monaco ou editar manualmente
3. Ctrl+S para salvar
4. Implantar → Gerenciar implantações → lápis (editar) → dropdown "Nova versão" → Implantar
5. A URL do endpoint NÃO muda

### Verificar se Apps Script está funcionando:
- No editor, selecionar função `testeConexao` no dropdown
- Clicar Executar
- Ver no "Registro de execução" se aparece `=== TUDO OK ===`

---

## 10. CONFIGURAÇÃO DO SYNC NO SISTEMA

O usuário configura a URL do Apps Script diretamente na interface:
- Botão ⚙️ "Configurar Drive" na tela final (step-out)
- Abre modal onde cola a Web App URL
- Salvo em `localStorage.setItem('apr_sync_url', url)`
- Ao gerar próxima APR, `autoSyncComPDF()` usa essa URL

---

## 11. HISTÓRICO DE VERSÕES DO APPS SCRIPT

| Versão | Data | Problema resolvido |
|---|---|---|
| v1–v8 | 17/05/2026 | Tentativas iniciais com DocumentApp (falhou por permissão) |
| v9 | 17/05/2026 | Tentativa com UrlFetchApp (falhou por permissão) |
| v10 | 17/05/2026 | PDF via SpreadsheetApp.create()+getAs. Colunas corrigidas |
| v11 | 19/05/2026 | PDF completo com fases/riscos/assinaturas via Sheets |
| **v12** | **20/05/2026** | **PDF do browser (html2pdf.js) salvo diretamente. Estado atual** |

---

## 12. POSSÍVEIS MELHORIAS FUTURAS (não implementadas)

- Exportar APR como DOCX além de PDF
- Histórico de APRs com filtro por data/elaborador na UI
- Notificação por e-mail automática ao gerar APR (Apps Script com MailApp)
- Atualização da lista de colaboradores via planilha externa (sem editar HTML)
- Modo multi-obra (trocar o canteiro sem editar o código)
- Campo de número de APR sequencial persistido em Sheets
- Relatório mensal de APRs geradas
- QR code funcional (link para a APR no Drive em vez do rastreio puro)

---

## 13. CONSTRAINTS DE DESENVOLVIMENTO

### NUNCA fazer:
- Usar `DocumentApp` no Apps Script (sem permissão)
- Usar `UrlFetchApp` no Apps Script (sem permissão)
- Usar `application/json` como Content-Type com `mode:'no-cors'`
- Adicionar dependências externas ao `index.html` além do CDN do html2pdf.js (já existe)
- Quebrar o single-file: tudo deve permanecer em `index.html`
- Alterar a URL do endpoint do Apps Script (é fixa por implantação)

### SEMPRE fazer:
- Ao editar Apps Script: implantar como Nova Versão (não apenas salvar)
- Ao editar index.html: `git add index.html && git commit && git push`
- Manter `APP_VERSAO` atualizado ao fazer versão major
- Preservar a ordem das 13 colunas da planilha
- Usar `mode:'no-cors'` + `Content-Type: text/plain` no fetch

---

## 14. AMBIENTE DA MÁQUINA

| Recurso | Estado |
|---|---|
| Git | ✅ Instalado |
| GitHub | ✅ Remote configurado (origin → remo-apr/apr) |
| Python | ❌ Não instalado (usar PowerShell ou browser automation) |
| Node.js | ❌ Não instalado |
| Excel (COM) | ✅ Office 16.0 disponível via PowerShell COM |
| Google Chrome | ✅ Disponível com extensão Claude in Chrome |
| Apps Script editor | Acessível via browser automation (Chrome) |

---

## 15. ARQUIVOS NA PASTA DO PROJETO

```
C:\Users\Remo Engenharia\Desktop\apr\
├── index.html                      # Sistema APR completo (ARQUIVO PRINCIPAL)
├── APR_REMO_Historico_Projeto.xlsx # Excel com histórico do projeto (5 abas)
├── APR_REMO_Documentacao.html      # HTML de documentação compartilhável
└── CONTEXTO_CLAUDE.md              # Este arquivo
```

---

## 16. RESUMO DO ESTADO ATUAL (20/05/2026)

**O sistema está funcionando em produção com:**
- ✅ index.html v4.1 publicado em https://remo-apr.github.io/apr/
- ✅ Apps Script v12 implantado — recebe PDF base64 do browser e salva no Drive
- ✅ PDF salvo no Drive = documento impresso (layout idêntico, gerado pelo html2pdf.js)
- ✅ Planilha de registro com 13 colunas e links para cada PDF
- ✅ Login de elaboradores funcionando
- ✅ PWA offline funcionando

**Última mudança crítica (20/05/2026):**
O PDF enviado ao Drive agora é gerado diretamente do documento HTML renderizado no browser via `html2pdf.js` (`autoSyncComPDF`), não mais reconstruído no Apps Script via Sheets. Isso garante que o PDF salvo é idêntico ao documento impresso.

---

*Documento gerado em 20/05/2026. Para continuar o projeto, leia as seções 2, 3 e 4 primeiro.*
