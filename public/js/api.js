// Acesso à API do servidor. O cabeçalho X-CLEAN é exigido pelo servidor em requisições que alteram dados.

export async function api(method, url, body) {
  const headers = { 'X-CLEAN': '1' };
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  let res;
  try {
    res = await fetch(url, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) });
  } catch {
    throw new Error('Não foi possível falar com o servidor do CLEAN. Verifique se ele está em execução.');
  }
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const error = new Error(data?.error || `Erro ${res.status} ao acessar ${url}`);
    error.status = res.status;
    error.code = data?.code || null; // ex.: 'changed' (arquivo alterado depois da análise)
    throw error;
  }
  return data;
}

export const get = (url) => api('GET', url);
export const post = (url, body = {}) => api('POST', url, body);
export const put = (url, body) => api('PUT', url, body);
export const del = (url) => api('DELETE', url);
