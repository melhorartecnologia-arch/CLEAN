// Microsoft Intune pela API Microsoft Graph (device management): lista os dispositivos gerenciados
// (iPhone/iPad) e envia a ação de reinício (rebootNow). Reaproveita as credenciais de uma conexão
// Microsoft 365 (Graph) já cadastrada — o aplicativo precisa, além do acesso que já tem, das
// permissões DeviceManagementManagedDevices.Read.All (listar) e
// DeviceManagementManagedDevices.PrivilegedOperations.All (reiniciar), tipo Aplicativo, com
// consentimento do administrador.
import { ApiError } from '../mail/http.js';
import { GraphClient, enc, commonGraphError, detailedGraphError } from '../cloud/graph-client.js';

// Sistemas da Apple no Intune (o iPad pode aparecer como "iOS" ou "iPadOS").
export const APPLE_OS = ['iOS', 'iPadOS'];

const SELECT = [
  'id',
  'deviceName',
  'managedDeviceName',
  'operatingSystem',
  'osVersion',
  'model',
  'manufacturer',
  'serialNumber',
  'userPrincipalName',
  'userDisplayName',
  'emailAddress',
  'lastSyncDateTime',
  'enrolledDateTime',
  'complianceState',
  'managementAgent',
  'deviceEnrollmentType',
].join(',');

/** Um dispositivo gerenciado do Intune no formato usado pelo CLEAN. */
export function intuneDevice(d) {
  return {
    id: d.id,
    name: d.deviceName || d.managedDeviceName || '(sem nome)',
    os: d.operatingSystem || '',
    osVersion: d.osVersion || '',
    model: d.model || '',
    manufacturer: d.manufacturer || '',
    serialNumber: d.serialNumber || '',
    user: d.userPrincipalName || d.emailAddress || '',
    userName: d.userDisplayName || '',
    lastSync: d.lastSyncDateTime || null,
    enrolled: d.enrolledDateTime || null,
    compliance: d.complianceState || '',
  };
}

export class IntuneConnector extends GraphClient {
  translate(err) {
    if (!(err instanceof ApiError)) return err;
    // Um 403 no Intune é, de regra, falta das permissões de device management (e não da leitura de
    // usuários que commonGraphError supõe): a mensagem específica vem primeiro.
    if (err.status === 403) {
      return new ApiError(
        'Permissão insuficiente no Intune: conceda ao aplicativo as permissões DeviceManagementManagedDevices.Read.All (listar os aparelhos) e DeviceManagementManagedDevices.PrivilegedOperations.All (reiniciar), do tipo Aplicativo, com consentimento do administrador no Microsoft Entra ID.',
        err,
      );
    }
    const common = commonGraphError(err);
    if (common) return common;
    return detailedGraphError(err);
  }

  /**
   * Dispositivos gerenciados dos sistemas escolhidos (padrão: iPhone/iPad). Percorre as páginas do
   * Graph; o filtro por sistema é feito aqui (e não só no $filter) para cobrir os nomes "iOS"/"iPadOS".
   */
  async *devices({ os = APPLE_OS } = {}) {
    const wanted = new Set(os.map((o) => String(o).toLowerCase()));
    const filter = os.length ? `&$filter=${enc(os.map((o) => `operatingSystem eq '${o}'`).join(' or '))}` : '';
    let url = `/deviceManagement/managedDevices?$select=${SELECT}${filter}&$top=100`;
    while (url) {
      const page = await this.api(url);
      for (const d of page?.value || []) {
        if (!wanted.size || wanted.has(String(d.operatingSystem || '').toLowerCase())) yield intuneDevice(d);
      }
      url = this.next(page);
    }
  }

  /** Envia a ação de reinício imediato a um dispositivo (rebootNow). Lança em caso de falha. */
  async reboot(deviceId, { signal } = {}) {
    await this.api(`/deviceManagement/managedDevices/${enc(deviceId)}/rebootNow`, { method: 'POST', retries: 2, signal });
  }

  /** Teste da conexão para o Intune: autentica e lê o primeiro dispositivo iPhone/iPad. */
  async test() {
    await this.accessToken();
    const page = await this.api(`/deviceManagement/managedDevices?$select=id,deviceName,operatingSystem&$filter=${enc("operatingSystem eq 'iOS' or operatingSystem eq 'iPadOS'")}&$top=5`);
    const n = (page?.value || []).length;
    return { ok: true, message: `Intune acessível: ${n}${page?.['@odata.nextLink'] ? '+' : ''} dispositivo(s) iPhone/iPad.` };
  }

  async close() {}
}
