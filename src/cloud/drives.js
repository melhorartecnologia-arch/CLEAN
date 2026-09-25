// OneDrive e SharePoint pelo Microsoft Graph, com permissões de aplicativo: lista as contas de
// OneDrive (todas as do locatário ou as escolhidas) ou os sites do SharePoint (todos ou os escolhidos,
// com os subsites) e as bibliotecas de documentos de cada um; percorre as pastas; baixa o conteúdo
// dos arquivos (com limite de tamanho) e, se permitido, exclui arquivos.
import { GraphClient, commonGraphError, detailedGraphError, enc } from './graph-client.js';
import { ApiError } from '../mail/http.js';
import { foldText } from '../scan/matcher.js';

export const CLOUD_TYPES = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };

const ITEM_FIELDS = 'id,name,size,file,folder,package,remoteItem,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,eTag,cTag';
const SITE_FIELDS = 'id,name,displayName,webUrl,isPersonalSite';
const MAX_SUBSITE_DEPTH = 5;
// Contas sem OneDrive toleradas no teste da conexão antes de desistir (todas as contas do locatário).
const TEST_MAX_SKIPPED = 50;

/** Identificador de site ("servidor,coleção,site"): as vírgulas ficam como estão no endereço. */
const siteId = (id) => enc(id).replace(/%2C/gi, ',');
const trimSlash = (url) => String(url || '').replace(/\/+$/, '');

/** Pessoa de um identitySet do Graph: { name, email } (usuário; senão aplicativo ou dispositivo). */
export function person(identity) {
  const who = identity?.user || identity?.application || identity?.device;
  if (!who) return null;
  const email = String(who.email || '').trim();
  const name = String(who.displayName || '').trim();
  return email || name ? { name, email } : null;
}

/** Contas ou sites ignorados: padrões com curingas (* e ?), sem diferenciar maiúsculas nem acentos. */
function targetMatcher(patterns = []) {
  const rules = patterns
    .map((p) => trimSlash(String(p || '').trim()))
    .filter(Boolean)
    .map((p) => new RegExp(`^${foldText(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`));
  return (...values) => rules.some((re) => values.some((v) => v && re.test(foldText(trimSlash(v)))));
}

/** Site pessoal (OneDrive): fica de fora das análises de SharePoint (é analisado como OneDrive). */
const isPersonalSite = (site) => site.isPersonalSite === true || /-my\.sharepoint\.[a-z.]+\/personal\//i.test(site.webUrl || '');

export class DrivesConnector extends GraphClient {
  /** repo: repositório do tipo 'onedrive' ou 'sharepoint', com os segredos decifrados. */
  constructor(repo, options = {}) {
    super(repo, options);
    this.kind = repo.type === 'sharepoint' ? 'sharepoint' : 'onedrive';
    this.cloud = repo.cloud || {};
  }

  translate(err) {
    if (!(err instanceof ApiError)) return err;
    const common = commonGraphError(err);
    if (common) return common;
    if (err.status === 403 && /accessDenied/i.test(err.code)) {
      return new ApiError(
        'Acesso negado aos arquivos: conceda ao aplicativo as permissões Files.Read.All e Sites.Read.All (tipo Aplicativo) com consentimento do administrador — ou, com Sites.Selected, libere este site para o aplicativo.',
        err,
      );
    }
    return detailedGraphError(err);
  }

  /** Todos os itens de uma listagem paginada. */
  async collect(url) {
    const out = [];
    while (url) {
      const page = await this.api(url);
      out.push(...(page?.value || []));
      url = this.next(page);
    }
    return out;
  }

  /**
   * Bibliotecas a analisar: { id, kind, account, accountName, library, webUrl, owner, label }, ou
   * { skip, account, reason } (conta sem OneDrive), ou { error, account } (conta ou site inacessível).
   */
  async *drives() {
    if (this.kind === 'onedrive') yield* this.oneDrives();
    else yield* this.siteDrives();
  }

  async *oneDrives() {
    const excluded = targetMatcher(this.cloud.exclude);
    for await (const user of this.cloud.scope === 'all' ? this.allUsers() : this.listedUsers()) {
      if (user.error) {
        yield user;
        continue;
      }
      if (excluded(user.address, user.name)) continue;
      let drive;
      try {
        drive = await this.api(`/users/${enc(user.id || user.address)}/drive?$select=id,name,driveType,webUrl,owner`);
      } catch (err) {
        if (this.signal?.aborted) throw err;
        // Conta sem OneDrive: nunca acessado, sem licença, sala, caixa compartilhada ou convidado.
        if (err.status === 404) yield { skip: true, account: user.address, reason: 'sem OneDrive (não criado ou sem licença)' };
        else yield { error: err, account: user.address };
        continue;
      }
      const owner = person(drive.owner);
      yield {
        id: drive.id,
        kind: 'onedrive',
        account: user.address,
        accountName: user.name || owner?.name || '',
        library: drive.name || 'OneDrive',
        webUrl: drive.webUrl || '',
        owner: owner?.email || user.address,
        label: `OneDrive de ${user.address}`,
      };
    }
  }

  async *listedUsers() {
    for (const address of this.cloud.accounts || []) {
      try {
        const u = await this.resolveUser({ address }, { notFound: `A conta ${address} não foi encontrada no Microsoft 365.` });
        yield { id: u.id, address, name: u.name };
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { error: err, account: address };
      }
    }
  }

  async *siteDrives() {
    const excluded = targetMatcher(this.cloud.exclude);
    const seen = new Set();
    for await (const site of this.cloud.scope === 'all' ? this.allSites() : this.listedSites()) {
      if (site.error) {
        yield site;
        continue;
      }
      if (seen.has(site.id)) continue;
      seen.add(site.id);
      const url = trimSlash(site.webUrl);
      const name = site.displayName || site.name || url;
      if (excluded(url, name)) continue;
      let drives;
      try {
        drives = await this.collect(`/sites/${siteId(site.id)}/drives?$select=id,name,driveType,webUrl`);
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { error: err, account: url };
        continue;
      }
      for (const d of drives) {
        if (d.driveType && d.driveType !== 'documentLibrary') continue;
        yield { id: d.id, kind: 'sharepoint', account: url, accountName: name, library: d.name || 'Documentos', webUrl: d.webUrl || url, owner: null, label: `${name} › ${d.name}` };
      }
    }
  }

  /** Todos os sites do locatário (sem os OneDrive pessoais), com os subsites. */
  async *allSites() {
    let page;
    try {
      page = await this.api(`/sites/getAllSites?$select=${SITE_FIELDS}&$top=999`);
    } catch (err) {
      if (this.signal?.aborted || ![400, 404, 501].includes(err.status)) throw err;
      page = await this.api(`/sites?search=*&$select=${SITE_FIELDS}`);
    }
    for (;;) {
      for (const site of page?.value || []) {
        if (isPersonalSite(site)) continue;
        yield site;
        yield* this.subsites(site, 1);
      }
      const next = this.next(page);
      if (!next) return;
      page = await this.api(next);
    }
  }

  async *subsites(site, depth) {
    if (depth > MAX_SUBSITE_DEPTH) return;
    let list;
    try {
      list = await this.collect(`/sites/${siteId(site.id)}/sites?$select=${SITE_FIELDS}`);
    } catch (err) {
      if (this.signal?.aborted) throw err;
      this.log('warn', `Os subsites de ${trimSlash(site.webUrl)} não puderam ser listados: ${err.message}`);
      return;
    }
    for (const sub of list) {
      yield sub;
      yield* this.subsites(sub, depth + 1);
    }
  }

  async *listedSites() {
    for (const address of this.cloud.sites || []) {
      try {
        const site = await this.resolveSite(address);
        yield site;
        yield* this.subsites(site, 1);
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { error: err, account: address };
      }
    }
  }

  /**
   * Site pelo endereço. Aceita links copiados do navegador (páginas e bibliotecas do site): procura o
   * site mais próximo, sem subir além de /sites/Nome ou /teams/Nome.
   */
  async resolveSite(address) {
    let u;
    let segments;
    try {
      u = new URL(address);
      segments = u.pathname.split('/').filter(Boolean).map((s) => decodeURIComponent(s));
    } catch {
      throw new ApiError(`Endereço de site inválido: ${address}`);
    }
    const floor = /^(sites|teams)$/i.test(segments[0] || '') ? Math.min(2, segments.length) : 0;
    for (let n = segments.length; n >= floor; n--) {
      const rel = segments.slice(0, n);
      const path = rel.length ? `/sites/${u.hostname}:/${rel.map(enc).join('/')}` : `/sites/${u.hostname}`;
      try {
        return await this.api(`${path}?$select=${SITE_FIELDS}`);
      } catch (err) {
        if (err.status !== 404 && err.status !== 400) throw err;
      }
    }
    throw new ApiError(`Site não encontrado: ${address}. Confira o endereço (ex.: https://empresa.sharepoint.com/sites/Financeiro).`, { status: 404 });
  }

  /**
   * Percorre uma biblioteca: { type: 'file', item, relativePath } para cada arquivo, { type: 'dir' }
   * para cada pasta lida e { type: 'error', path, error } para pastas que não puderam ser lidas.
   * Atalhos para itens de outras bibliotecas ("Adicionar atalho a Meus arquivos") são ignorados.
   */
  async *walk(drive, { isExcluded = () => false, shouldStop = () => false } = {}) {
    const stack = [{ id: null, rel: '' }];
    while (stack.length > 0) {
      if (shouldStop()) return;
      const folder = stack.pop();
      const subdirs = [];
      const base = folder.id ? `/drives/${enc(drive.id)}/items/${enc(folder.id)}` : `/drives/${enc(drive.id)}/root`;
      let url = `${base}/children?$select=${ITEM_FIELDS}&$top=999`;
      try {
        while (url) {
          const page = await this.api(url);
          for (const item of page?.value || []) {
            if (item.remoteItem) continue;
            const rel = folder.rel ? `${folder.rel}/${item.name}` : item.name;
            if (isExcluded(item.name, rel)) continue;
            if (item.folder || item.package) subdirs.push({ id: item.id, rel });
            else if (item.file) yield { type: 'file', item, relativePath: rel };
          }
          if (shouldStop()) return;
          url = this.next(page);
        }
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { type: 'error', path: folder.rel ? `${drive.label} › ${folder.rel}` : drive.label, error: err };
        continue;
      }
      yield { type: 'dir', path: folder.rel };
      subdirs.sort((a, b) => b.rel.localeCompare(a.rel));
      stack.push(...subdirs);
    }
  }

  /** Conteúdo de um arquivo, até maxBytes (o restante não é baixado): { data, truncated, size }. */
  download(driveId, itemId, { maxBytes = Infinity } = {}) {
    return this.api(`/drives/${enc(driveId)}/items/${enc(itemId)}/content`, { type: 'buffer', maxBytes, retries: 4, headers: { Accept: '*/*' } });
  }

  /**
   * Exclui um arquivo: 'trash' = move para a lixeira do site ou do OneDrive (pode ser restaurado),
   * 'permanent' = exclusão definitiva. target: { driveId, itemId, cTag, eTag } registrados na análise:
   * se o arquivo mudou depois disso, nada é excluído (status 'changed'), a não ser com force.
   * Retorna { status: 'deleted' | 'missing' | 'changed' | 'failed', error? }.
   */
  async deleteItem(target, mode = 'trash', { force = false, signal } = {}) {
    const base = `/drives/${enc(target.driveId)}/items/${enc(target.itemId)}`;
    const tag = target.cTag || target.eTag || '';
    // Uma tentativa interrompida (falha de rede, 5xx) pode ter sido feita pelo servidor: nesse caso,
    // "não encontrado" na repetição quer dizer que a exclusão funcionou.
    let uncertain = false;
    const onRetry = ({ error }) => {
      if (error.status !== 429) uncertain = true;
    };
    try {
      if (!force && tag) {
        const current = await this.api(`${base}?$select=id,cTag,eTag`, { signal });
        const now = target.cTag ? current?.cTag : current?.eTag;
        if (now !== tag) return { status: 'changed', error: 'O arquivo foi alterado depois da análise.' };
      }
      if (mode === 'permanent') await this.api(`${base}/permanentDelete`, { method: 'POST', retries: 4, signal, onRetry });
      else await this.api(base, { method: 'DELETE', retries: 4, signal, onRetry, headers: !force && tag ? { 'If-Match': tag } : {} });
      return { status: 'deleted' };
    } catch (err) {
      if (err.status === 404) {
        return uncertain ? { status: 'deleted' } : { status: 'missing', error: 'O arquivo não existe mais (já excluído, movido para a lixeira ou para outra biblioteca).' };
      }
      if (err.status === 412) return { status: 'changed', error: 'O arquivo foi alterado depois da análise.' };
      if (err.status === 423) return { status: 'failed', error: 'O arquivo está bloqueado (aberto para edição por alguém ou em check-out).' };
      if (err.status === 403) {
        return {
          status: 'failed',
          error: 'Sem permissão para excluir: conceda ao aplicativo a permissão Files.ReadWrite.All (ou Sites.ReadWrite.All). Arquivos com rótulo de retenção (registro) não podem ser excluídos.',
        };
      }
      return { status: 'failed', error: err.message };
    }
  }

  /** Teste da conexão: autenticação e acesso a até 3 bibliotecas. */
  async test() {
    await this.accessToken();
    const details = [];
    let checked = 0;
    let skipped = 0;
    for await (const d of this.drives()) {
      if (d.error) return { ok: false, message: `${d.account}: ${d.error.message}`, details };
      if (d.skip) {
        if (++skipped <= 3) details.push(`${d.account}: ${d.reason}.`);
        if (skipped >= TEST_MAX_SKIPPED && checked === 0) break;
        continue;
      }
      const page = await this.api(`/drives/${enc(d.id)}/root/children?$select=id&$top=200`);
      const count = page?.value?.length || 0;
      details.push(`${d.label}: acesso OK, ${count}${page?.['@odata.nextLink'] ? ' ou mais' : ''} item(ns) na raiz.`);
      if (++checked >= 3) break;
    }
    if (checked === 0) {
      return { ok: false, message: this.kind === 'onedrive' ? 'Nenhuma das contas testadas tem OneDrive.' : 'Nenhuma biblioteca de documentos encontrada.', details };
    }
    return { ok: true, message: this.kind === 'onedrive' ? 'Conexão com o OneDrive funcionando.' : 'Conexão com o SharePoint funcionando.', details };
  }
}

// -- Proteção por cadastro ------------------------------------------------------------------------

const sameTenant = (a, b) => String(a.graph?.tenantId || '').toLowerCase() === String(b.graph?.tenantId || '').toLowerCase();
const siteKey = (url) => trimSlash(String(url || '').split(/[?#]/)[0]).toLowerCase();
/** Um site contém o outro (ou é o mesmo): vale nos dois sentidos, para errar a favor da proteção. */
const sitesOverlap = (a, b) => a === b || a.startsWith(`${b}/`) || b.startsWith(`${a}/`);

/**
 * Contas e sites protegidos para um repositório na nuvem: outros repositórios do mesmo tipo e
 * locatário, sem "Permitir exclusão", que listam contas ou sites (o cadastro mais específico vale,
 * como nas pastas aninhadas); outro repositório com o mesmo alcance ("todos") também protege.
 */
export function keptCloud(repo, repositories) {
  const keep = { all: null, accounts: [], sites: [] };
  for (const other of repositories) {
    if (other.id === repo.id || other.allowDelete || other.type !== repo.type || !sameTenant(other, repo)) continue;
    const error = `Protegido pelo repositório "${other.name}", que não permite exclusão.`;
    if (other.cloud?.scope === 'all') {
      if (repo.cloud?.scope === 'all') keep.all ||= { error };
      continue;
    }
    for (const account of other.cloud?.accounts || []) keep.accounts.push({ value: String(account).toLowerCase(), error });
    for (const site of other.cloud?.sites || []) keep.sites.push({ value: siteKey(site), error });
  }
  return keep;
}

/** Proteção que vale para um arquivo (record.cloud), ou null. */
export function keptCloudTarget(target, keep) {
  if (!keep) return null;
  if (keep.all) return keep.all;
  if (target.kind === 'onedrive') return keep.accounts.find((k) => k.value === String(target.account || '').toLowerCase()) || null;
  const site = siteKey(target.account);
  return keep.sites.find((k) => sitesOverlap(site, k.value)) || null;
}
