// Mundo simulado do Microsoft 365 para os testes de OneDrive e SharePoint: contas, sites, subsites,
// bibliotecas e arquivos, e funções para rodar análises contra o simulador do Graph.
import fs from 'node:fs';
import { Scanner } from '../../src/scan/scanner.js';
import { startMockApis } from './mock-apis.js';

export const TENANT = 'contoso.onmicrosoft.com';
export const CLIENT = '11111111-2222-3333-4444-555555555555';
export const SECRET = 'segredo';
export const TERMS = [
  { id: 'l:conf', type: 'text', value: 'confidencial', listName: 'RH' },
  { id: 'l:sal', type: 'text', value: 'salário', listName: 'RH' },
];
export const LIST_TERMS = TERMS.map(({ id, listName, ...t }) => t);
const DOCX = fs.readFileSync(new URL('../fixtures/doc.docx', import.meta.url));

export const ANA = { name: 'Ana Souza', email: 'ana@contoso.com' };
export const BRUNO = { name: 'Bruno Lima', email: 'bruno@contoso.com' };
export const file = (id, name, text, extra = {}) => ({ id, name, content: Buffer.isBuffer(text) ? text : Buffer.from(text, 'utf8'), lastModifiedBy: ANA, createdBy: BRUNO, ...extra });
export const folder = (id, name, children) => ({ id, name, children });

/** Dados do simulador: contas de OneDrive e sites do SharePoint com bibliotecas e arquivos. */
export function world() {
  const drives = {
    'd-ana': {
      id: 'd-ana',
      name: 'OneDrive',
      webUrl: 'https://contoso-my.sharepoint.com/personal/ana_contoso_com/Documents',
      items: [
        folder('f-docs', 'Documentos', [
          file('i-folha', 'folha.txt', 'salário de todos'),
          file('i-limpo', 'limpo.txt', 'nada aqui'),
          file('i-docx', 'Contrato João.docx', DOCX, { lastModifiedBy: BRUNO }),
          file('i-temp', '~$Contrato.docx', 'confidencial'),
          folder('f-antigo', 'Arquivo morto', [file('i-velho', 'velho.txt', 'confidencial antigo')]),
        ]),
        { ...file('i-atalho', 'Atalho', 'confidencial'), remote: true },
      ],
    },
    'd-carla': {
      id: 'd-carla',
      name: 'OneDrive',
      webUrl: 'https://contoso-my.sharepoint.com/personal/carla_contoso_com/Documents',
      items: [file('i-conf', 'confidencial.txt', 'sem termos no texto', { lastModifiedBy: { name: 'Carla Dias', email: 'carla@contoso.com' } }), file('i-grande', 'grande.txt', `confidencial ${'x'.repeat(1.5 * 1048576)}`)],
    },
    'd-root': { id: 'd-root', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/Shared Documents', items: [file('i-raiz', 'politica.txt', 'material confidencial')] },
    'd-fin': {
      id: 'd-fin',
      name: 'Documentos Compartilhados',
      webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Shared Documents',
      items: [folder('f-2026', '2026', [file('i-fin', 'salarios.txt', 'salário da diretoria'), file('i-fin2', 'orcamento.txt', 'nada')])],
    },
    'd-fin-assets': { id: 'd-fin-assets', name: 'Site Assets', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/SiteAssets', items: [file('i-asset', 'confidencial.css', 'x')] },
    'd-contratos': { id: 'd-contratos', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos/Shared Documents', items: [file('i-ctr', 'contrato.txt', 'cláusula confidencial')] },
    'd-jur': { id: 'd-jur', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/sites/Juridico/Shared Documents', items: [file('i-jur', 'processo.txt', 'confidencial')] },
  };
  const sites = [
    { id: 'contoso.sharepoint.com,r1,r2', name: 'contoso', displayName: 'Comunicação', webUrl: 'https://contoso.sharepoint.com', driveIds: ['d-root'] },
    { id: 'contoso.sharepoint.com,f1,f2', name: 'Financeiro', displayName: 'Financeiro', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro', driveIds: ['d-fin', 'd-fin-assets'] },
    { id: 'contoso.sharepoint.com,c1,c2', name: 'Contratos', displayName: 'Contratos', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos', driveIds: ['d-contratos'], parent: 'contoso.sharepoint.com,f1,f2' },
    { id: 'contoso.sharepoint.com,j1,j2', name: 'Juridico', displayName: 'Jurídico', webUrl: 'https://contoso.sharepoint.com/sites/Juridico', driveIds: ['d-jur'] },
    { id: 'contoso-my.sharepoint.com,p1,p2', name: 'ana', displayName: 'Ana Souza', webUrl: 'https://contoso-my.sharepoint.com/personal/ana_contoso_com', driveIds: ['d-ana'], personal: true },
  ];
  const users = [
    { id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana Souza', driveId: 'd-ana', folders: [], messages: {} },
    { id: 'u-bruno', mail: 'bruno@contoso.com', displayName: 'Bruno Lima', folders: [], messages: {} },
    { id: 'u-carla', mail: 'carla@contoso.com', displayName: 'Carla Dias', driveId: 'd-carla', folders: [], messages: {} },
  ];
  return { tenant: TENANT, clientId: CLIENT, secret: SECRET, users, drives, sites };
}

export async function withGraph(fn) {
  const data = world();
  const mock = await startMockApis({ graph: data });
  try {
    return await fn(data, mock.endpoints);
  } finally {
    await mock.close();
  }
}

export const repo = (type, extra = {}) => ({
  id: `r-${type}`,
  type,
  name: type === 'onedrive' ? 'OneDrive da empresa' : 'SharePoint da empresa',
  path: '',
  exclude: ['Arquivo morto', 'Site Assets'],
  allowDelete: false,
  deleteMode: 'trash',
  graph: { tenantId: TENANT, clientId: CLIENT },
  secrets: { clientSecret: SECRET },
  cloud: { scope: 'all', accounts: [], sites: [], exclude: [] },
  ...extra,
});

export async function scan(repositories, endpoints, options = {}, deps = {}) {
  const messages = [];
  const stats = await new Scanner({ repositories, terms: TERMS, options: { maxFileSizeMB: 1, ...options }, endpoints }, (m) => messages.push(m), deps).run();
  return {
    stats,
    messages,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}
