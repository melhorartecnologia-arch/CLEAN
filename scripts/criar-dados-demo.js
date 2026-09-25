// Cria uma pasta de demonstração com arquivos de exemplo e cadastra um repositório e uma lista de
// referência para experimentar o CLEAN. Uso: npm run demo
import fs from 'node:fs';
import path from 'node:path';
import { loadConfig, PROJECT_ROOT } from '../src/config.js';
import { Store } from '../src/store.js';
import { PRESETS } from '../src/scan/presets.js';

const FIXTURES = path.join(PROJECT_ROOT, 'test', 'fixtures');
const DEMO = path.join(PROJECT_ROOT, 'demo', 'Compartilhamento');

const copies = {
  'RH/Folha_Salarios_2025.xlsx': 'plan.xlsx',
  'RH/Contrato_Joao_da_Silva.docx': 'doc.docx',
  'RH/Antigos/relatorio_1999.doc': 'doc.doc',
  'RH/Antigos/planilha_antiga.xls': 'plan.xls',
  'Financeiro/Apresentacao_Demissoes.pptx': 'apres.pptx',
  'Financeiro/extrato.pdf': 'doc.pdf',
  'Financeiro/documento_protegido.docx': 'senha.docx',
  'Comercial/proposta.rtf': 'doc.rtf',
  'Comercial/apresentacao.odp': 'apres.odp',
};

const texts = {
  'TI/senhas_servidores.txt': 'Servidor: SRV-ARQ01\r\nusuario: administrador\r\nsenha: Pa$$w0rd2025\r\n',
  'Comercial/leia-me.txt': 'Pasta do time comercial. Nada sensível aqui.\r\n',
  'Comercial/contatos.html':
    '<html><head><title>Contatos</title></head><body><table><tr><td>Fornecedor XPTO</td><td>CNPJ 11.222.333/0001-81</td><td>compras@xpto.com.br</td></tr></table></body></html>',
};

// CSV no formato do Excel em português (Windows-1252, separador ";")
const csv = 'Nome;CPF;E-mail;Situação\r\nJoão da Silva;529.982.247-25;joao.silva@empresa.com.br;Ativo\r\nMaria Souza;111.444.777-35;maria@empresa.com.br;Demissão\r\n';

function write(rel, data) {
  const file = path.join(DEMO, rel);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, data);
}

for (const [rel, fixture] of Object.entries(copies)) write(rel, fs.readFileSync(path.join(FIXTURES, fixture)));
for (const [rel, content] of Object.entries(texts)) write(rel, content);
write('TI/clientes.csv', Buffer.from([...csv].map((c) => c.charCodeAt(0))));
console.log(`Arquivos de demonstração criados em ${DEMO}`);

const repository = { name: 'Demonstração', path: DEMO, description: 'Arquivos de exemplo do CLEAN', exclude: [], audit: { enabled: false, days: 30, maxEvents: 200000 } };
const pick = (id) => {
  const p = PRESETS.find((x) => x.id === id);
  return { type: 'regex', value: p.value, validator: p.validator, label: p.label, wholeWord: false };
};
const list = {
  name: 'Exemplo – dados sensíveis',
  description: 'Termos de RH e dados pessoais (LGPD)',
  terms: [
    ...['salário', 'demissão', 'confidencial', 'sigiloso', 'João da Silva'].map((value) => ({ type: 'text', value, wholeWord: false, validator: null, label: '' })),
    pick('cpf'),
    pick('cnpj'),
    pick('email'),
    pick('senha'),
  ],
};

const config = loadConfig();
const base = `http://${config.host === '0.0.0.0' ? '127.0.0.1' : config.host}:${config.port}`;
const headers = { 'Content-Type': 'application/json', 'X-CLEAN': '1' };
if (config.authUser) headers.Authorization = `Basic ${Buffer.from(`${config.authUser}:${config.authPassword}`).toString('base64')}`;

async function viaApi() {
  const repos = await (await fetch(`${base}/api/repositories`, { headers })).json();
  if (!repos.some((r) => r.path === DEMO)) await fetch(`${base}/api/repositories`, { method: 'POST', headers, body: JSON.stringify(repository) });
  const lists = await (await fetch(`${base}/api/lists`, { headers })).json();
  if (!lists.some((l) => l.name === list.name)) await fetch(`${base}/api/lists`, { method: 'POST', headers, body: JSON.stringify(list) });
  console.log(`Repositório e lista cadastrados no servidor em ${base}.`);
}

async function viaStore() {
  const store = await new Store(config.dataDir).init();
  if (!store.listRepositories().some((r) => r.path === DEMO)) store.createRepository(repository);
  if (!store.listLists().some((l) => l.name === list.name)) {
    store.createList({ ...list, terms: list.terms.map((t) => ({ ...t, id: crypto.randomUUID() })) });
  }
  await store.saveNow();
  console.log(`Repositório e lista cadastrados em ${config.dataDir}. Inicie o servidor com: npm start`);
}

try {
  await viaApi();
} catch {
  await viaStore();
}
