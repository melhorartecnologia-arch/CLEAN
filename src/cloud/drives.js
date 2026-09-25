// OneDrive e SharePoint pelo Microsoft Graph, com permissões de aplicativo: lista as contas de
// OneDrive (todas as do locatário ou as escolhidas) ou os sites do SharePoint (todos ou os escolhidos,
// com os subsites) e as bibliotecas de documentos de cada um; percorre as pastas; baixa o conteúdo
// dos arquivos (com limite de tamanho) e, se permitido, exclui arquivos.
import { GraphClient, commonGraphError, detailedGraphError, enc } from './graph-client.js';
import { ApiError } from '../mail/http.js';
import { foldText } from '../scan/matcher.js';

export const CLOUD_TYPES = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };

const ITEM_FIELDS = 'id,name,size,file,folder,package,remoteItem,parentReference,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,webUrl,eTag,cTag';
const SITE_FIELDS = 'id,name,displayName,webUrl,isPersonalSite';
const MAX_SUBSITE_DEPTH = 10;
// Contas sem OneDrive toleradas no teste da conexão antes de desistir (todas as contas do locatário).
const TEST_MAX_SKIPPED = 50;

/** Identificador de site ("servidor,coleção,site"): as vírgulas ficam como estão no endereço. */
const siteId = (id) => enc(id).replace(/%2C/gi, ',');
const trimSlash = (url) => String(url || '').replace(/\/+$/, '');
const lower = (v) => String(v || '').toLowerCase();

/** Pessoa de um identitySet do Graph: { name, email } (usuário; senão aplicativo ou dispositivo). */
export function person(identity) {
  const who = identity?.user || identity?.application || identity?.device;
  if (!who) return null;
  const email = String(who.email || '').trim();
  const name = String(who.displayName || '').trim();
  return email || name ? { name, email } : null;
}

/**
 * Contas ou sites ignorados: padrões com curingas (* e ?), sem diferenciar maiúsculas nem acentos.
 * Um endereço de site ignora também os subsites (tudo o que começa com ele e uma barra).
 */
export function targetMatcher(patterns = []) {
  const rules = [];
  for (const raw of patterns) {
    const p = trimSlash(String(raw || '').trim());
    if (!p) continue;
    const source = foldText(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    rules.push(new RegExp(/^https?:/i.test(p) ? `^${source}(/.*)?$` : `^${source}$`));
  }
  return (...values) => rules.some((re) => values.some((v) => v && re.test(foldText(trimSlash(v)))));
}

/** Site pessoal (OneDrive): fica de fora das análises de SharePoint (é analisado como OneDrive). */
export const isPersonalUrl = (url) => /-my\.sharepoint\.[a-z.]+\/personal\//i.test(`${trimSlash(url)}/`);
const isPersonalSite = (site) => site.isPersonalSite === true || isPersonalUrl(site.webUrl);

// Partes do endereço que já não são o site: bibliotecas, páginas e listas (links copiados do navegador).
const NOT_SITE = new Set(['shared documents', 'documentos compartilhados', 'documents', 'documentos', 'siteassets', 'site assets', 'sitepages', 'pages', 'lists', 'style library', '_layouts', '_api', '_vti_bin']);

/**
 * Caminho do site num endereço copiado do navegador: corta a partir de uma biblioteca, página ou lista
 * conhecida ("/Forms/..." corta também a biblioteca antes dele). Não sobe além disso: um endereço de
 * subsite com erro de digitação não vira o site acima dele.
 */
export function sitePathSegments(segments) {
  const start = /^(sites|teams)$/i.test(segments[0] || '') ? 2 : 0;
  for (let i = start; i < segments.length; i++) {
    const seg = segments[i].toLowerCase();
    if (seg === 'forms' && i > start) return segments.slice(0, i - 1);
    if (NOT_SITE.has(seg) || seg.endsWith('.aspx')) return segments.slice(0, i);
  }
  return segments;
}

/** Endereço do site (sem páginas e bibliotecas), em minúsculas, para comparações. */
export function siteKey(url) {
  const clean = String(url || '').split(/[?#]/)[0];
  try {
    const u = new URL(clean);
    const segments = u.pathname.split('/').filter(Boolean).map((s) => decodeURIComponent(s));
    const path = sitePathSegments(segments).join('/');
    return lower(`https://${u.hostname}${path ? `/${path}` : ''}`);
  } catch {
    return lower(trimSlash(clean));
  }
}

/**
 * O site `child` é o próprio `parent` ou um subsite dele. Os sites em /sites/ e /teams/ não são
 * subsites do site raiz (são coleções separadas).
 */
export function siteWithin(parent, child) {
  if (child === parent) return true;
  if (!child.startsWith(`${parent}/`)) return false;
  const rest = child.slice(parent.length);
  let parentIsRoot = false;
  try {
    parentIsRoot = !new URL(parent).pathname.replace(/\/+$/, '');
  } catch {
    // endereço inválido: compara só o texto
  }
  return !(parentIsRoot && /^\/(sites|teams)\//i.test(rest));
}

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
   * Bibliotecas a analisar: { id, kind, account, accountName, aliases, library, webUrl, owner, label },
   * ou { skip, account, reason } (conta sem OneDrive), ou { error, account } (conta ou site inacessível).
   */
  async *drives() {
    if (this.kind === 'onedrive') yield* this.oneDrives();
    else yield* this.siteDrives();
  }

  async *oneDrives() {
    const excluded = targetMatcher(this.cloud.exclude);
    const seen = new Set();
    for await (const user of this.cloud.scope === 'all' ? this.allUsers() : this.listedUsers()) {
      if (user.error) {
        yield user;
        continue;
      }
      // Contas ignoradas pelo e-mail, pelo nome de logon (UPN) ou pelo nome.
      if (excluded(user.address, user.mail, user.upn, user.name)) continue;
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
      // A mesma conta informada duas vezes (e-mail e nome de logon) é analisada uma vez só.
      if (seen.has(drive.id)) continue;
      seen.add(drive.id);
      const owner = person(drive.owner);
      yield {
        id: drive.id,
        kind: 'onedrive',
        account: user.address,
        accountName: user.name || owner?.name || '',
        aliases: [...new Set([user.address, user.mail, user.upn, owner?.email].filter(Boolean).map(lower))],
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
        yield { ...u, address };
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { error: err, account: address };
      }
    }
  }

  async *siteDrives() {
    const excluded = targetMatcher(this.cloud.exclude);
    const seen = new Set();
    for await (const site of this.cloud.scope === 'all' ? this.allSites(excluded) : this.listedSites(excluded)) {
      if (site.error) {
        yield site;
        continue;
      }
      if (seen.has(site.id)) continue;
      seen.add(site.id);
      const url = trimSlash(site.webUrl);
      const name = site.displayName || site.name || url;
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
        yield { id: d.id, kind: 'sharepoint', account: url, accountName: name, aliases: [lower(url)], library: d.name || 'Documentos', webUrl: d.webUrl || url, owner: null, label: `${name} › ${d.name}` };
      }
    }
  }

  /** Site ignorado (pelo endereço ou pelo nome): nem ele nem os subsites são analisados. */
  static skipSite(site, excluded) {
    return excluded(site.webUrl, site.displayName || site.name);
  }

  /** Todos os sites do locatário (sem os OneDrive pessoais), com os subsites. */
  async *allSites(excluded) {
    let page;
    try {
      page = await this.api(`/sites/getAllSites?$select=${SITE_FIELDS}&$top=999`);
    } catch (err) {
      if (this.signal?.aborted || ![400, 404, 501].includes(err.status)) throw err;
      page = await this.api('/sites?search=*');
    }
    for (;;) {
      for (const site of page?.value || []) {
        if (isPersonalSite(site) || DrivesConnector.skipSite(site, excluded)) continue;
        yield site;
        yield* this.subsites(site, 1, excluded);
      }
      const next = this.next(page);
      if (!next) return;
      page = await this.api(next);
    }
  }

  async *subsites(site, depth, excluded) {
    if (depth > MAX_SUBSITE_DEPTH) {
      this.log('warn', `Subsites de ${trimSlash(site.webUrl)} com mais de ${MAX_SUBSITE_DEPTH} níveis não foram analisados.`);
      return;
    }
    let list;
    try {
      list = await this.collect(`/sites/${siteId(site.id)}/sites?$select=${SITE_FIELDS}`);
    } catch (err) {
      if (this.signal?.aborted) throw err;
      yield { error: new ApiError(`Os subsites não puderam ser listados: ${err.message}`, err), account: trimSlash(site.webUrl) };
      return;
    }
    for (const sub of list) {
      if (DrivesConnector.skipSite(sub, excluded)) continue;
      yield sub;
      yield* this.subsites(sub, depth + 1, excluded);
    }
  }

  async *listedSites(excluded) {
    for (const address of this.cloud.sites || []) {
      try {
        const site = await this.resolveSite(address);
        if (DrivesConnector.skipSite(site, excluded)) continue;
        yield site;
        yield* this.subsites(site, 1, excluded);
      } catch (err) {
        if (this.signal?.aborted) throw err;
        yield { error: err, account: address };
      }
    }
  }

  /**
   * Site pelo endereço. Aceita links copiados do navegador (páginas, bibliotecas e listas do site),
   * mas não procura outro site: um endereço inexistente é um erro.
   */
  async resolveSite(address) {
    let u;
    let segments;
    try {
      u = new URL(address);
      segments = sitePathSegments(u.pathname.split('/').filter(Boolean).map((s) => decodeURIComponent(s)));
    } catch {
      throw new ApiError(`Endereço de site inválido: ${address}`);
    }
    const path = segments.length ? `/sites/${u.hostname}:/${segments.map(enc).join('/')}` : `/sites/${u.hostname}`;
    try {
      return await this.api(`${path}?$select=${SITE_FIELDS}`);
    } catch (err) {
      if (err.status !== 404 && err.status !== 400) throw err;
      throw new ApiError(`Site não encontrado: ${address}. Confira o endereço (ex.: https://empresa.sharepoint.com/sites/Financeiro).`, { status: 404 });
    }
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
      let failure = null;
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
        failure = err;
      }
      if (failure) yield { type: 'error', path: folder.rel ? `${drive.label} › ${folder.rel}` : drive.label, error: failure };
      else yield { type: 'dir', path: folder.rel };
      // As subpastas já listadas continuam na fila mesmo se uma página seguinte falhar.
      subdirs.sort((a, b) => b.rel.localeCompare(a.rel));
      stack.push(...subdirs);
    }
  }

  /** Conteúdo de um arquivo, até maxBytes (o restante não é baixado): { data, truncated, size }. */
  download(driveId, itemId, { maxBytes = Infinity, signal } = {}) {
    return this.api(`/drives/${enc(driveId)}/items/${enc(itemId)}/content`, { type: 'buffer', maxBytes, retries: 4, signal, headers: { Accept: '*/*' } });
  }

  /**
   * Exclui um arquivo: 'trash' = move para a lixeira do site ou do OneDrive (pode ser restaurado),
   * 'permanent' = exclusão definitiva. target (record.cloud): driveId, itemId e o que a análise viu
   * (cTag/eTag, nome, pasta, tamanho e data). Se o arquivo mudou de conteúdo, de nome ou de pasta
   * depois disso, nada é excluído (status 'changed'), a não ser com force.
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
      if (!force) {
        const current = await this.api(`${base}?$select=id,name,size,cTag,eTag,lastModifiedDateTime,parentReference`, { signal });
        const changed = changedSince(target, current);
        if (changed) return { status: 'changed', error: changed };
      }
      // Com If-Match, o servidor também recusa se o arquivo mudar entre a conferência e a exclusão.
      const headers = !force && tag ? { 'If-Match': tag } : {};
      if (mode === 'permanent') await this.api(`${base}/permanentDelete`, { method: 'POST', retries: 4, signal, onRetry, headers });
      else await this.api(base, { method: 'DELETE', retries: 4, signal, onRetry, headers });
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

  /**
   * Protegidos por outros repositórios (keptCloud): as contas de OneDrive são localizadas no Microsoft
   * 365 (o mesmo usuário pode estar cadastrado pelo e-mail ou pelo nome de logon) e comparadas pelo
   * identificador do OneDrive. Se a consulta falhar, os arquivos do OneDrive ficam protegidos.
   */
  async resolveKept(keep) {
    if (!keep || this.kind !== 'onedrive' || keep.driveIds || !keep.accounts.length) return keep;
    keep.driveIds = new Map();
    for (const account of keep.accounts) {
      try {
        const u = await this.resolveUser({ address: account.value });
        const drive = await this.api(`/users/${enc(u.id)}/drive?$select=id`);
        keep.driveIds.set(drive.id, account);
      } catch (err) {
        if (err.status === 404) continue; // conta inexistente ou sem OneDrive: nada a proteger
        keep.unresolved = { error: `Não foi possível conferir as contas protegidas por outros repositórios (${err.message}); o arquivo não foi excluído.` };
        break;
      }
    }
    return keep;
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

/** O que mudou no arquivo desde a análise (conteúdo, nome ou pasta), ou null. */
function changedSince(target, current) {
  const tag = target.cTag ? current?.cTag : current?.eTag;
  if (target.cTag || target.eTag) {
    if (tag !== (target.cTag || target.eTag)) return 'O arquivo foi alterado depois da análise.';
  } else if (Number(current?.size) !== Number(target.size) || Date.parse(current?.lastModifiedDateTime) !== Date.parse(target.modified)) {
    return 'O arquivo foi alterado depois da análise (tamanho ou data de modificação diferentes).';
  }
  if (target.name && current?.name !== target.name) return 'O arquivo foi renomeado depois da análise.';
  if (target.parentId && current?.parentReference?.id && current.parentReference.id !== target.parentId) return 'O arquivo foi movido para outra pasta depois da análise.';
  return null;
}

// -- Proteção por cadastro ------------------------------------------------------------------------

/**
 * Contas e sites protegidos para um repositório na nuvem: outros repositórios do mesmo tipo, sem
 * "Permitir exclusão", que listam contas ou sites (o cadastro mais específico vale, como nas pastas
 * aninhadas); outro repositório com o mesmo alcance ("todos") também protege. Endereços de e-mail e
 * de sites são únicos entre locatários, então o locatário não precisa ser comparado.
 */
export function keptCloud(repo, repositories) {
  const keep = { all: null, accounts: [], sites: [] };
  for (const other of repositories) {
    if (other.id === repo.id || other.allowDelete || other.type !== repo.type) continue;
    const error = `Protegido pelo repositório "${other.name}", que não permite exclusão.`;
    if (other.cloud?.scope === 'all') {
      if (repo.cloud?.scope === 'all') keep.all ||= { error };
      continue;
    }
    for (const account of other.cloud?.accounts || []) keep.accounts.push({ value: lower(account), error });
    for (const site of other.cloud?.sites || []) keep.sites.push({ value: siteKey(site), error });
  }
  return keep;
}

/** Proteção que vale para um arquivo (record.cloud), ou null. Use resolveKept antes (OneDrive). */
export function keptCloudTarget(target, keep) {
  if (!keep) return null;
  if (keep.all) return keep.all;
  if (target.kind === 'onedrive') {
    const aliases = new Set([target.account, ...(target.aliases || [])].map(lower));
    return keep.driveIds?.get(target.driveId) || keep.accounts.find((k) => aliases.has(k.value)) || (keep.accounts.length ? keep.unresolved : null) || null;
  }
  const site = siteKey(target.account);
  // Um site protegido protege os subsites; um subsite protegido também protege o site acima (a favor da proteção).
  return keep.sites.find((k) => siteWithin(k.value, site) || siteWithin(site, k.value)) || null;
}

/**
 * O arquivo (record.cloud) continua no alcance do repositório: a conta ou o site está na lista (ou
 * o repositório abrange todos) e não é ignorado. Arquivos de contas/sites retirados do cadastro
 * depois da análise não são excluídos pelo relatório.
 */
export function coveredByRepo(repo, target) {
  const c = repo.cloud || {};
  const aliases = [target.account, ...(target.aliases || [])];
  if (targetMatcher(c.exclude)(...aliases, target.accountName)) return false;
  if (c.scope === 'all') return true;
  if (target.kind === 'onedrive') {
    const known = new Set(aliases.map(lower));
    return (c.accounts || []).some((a) => known.has(lower(a)));
  }
  const site = siteKey(target.account);
  return (c.sites || []).some((s) => siteWithin(siteKey(s), site));
}

/** O que a exclusão confere de um arquivo da análise (record): identificação e como ele estava. */
export const cloudTarget = (record) => ({ ...record.cloud, name: record.name, size: record.size, modified: record.modified });
