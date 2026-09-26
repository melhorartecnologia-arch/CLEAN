// Novo agendamento (#/agendamentos/novo?tipo=email) ou edição (#/agendamentos/<id>): o mesmo
// formulário da nova análise de arquivos ou de e-mail, no modo de agendamento.
import { get } from '../api.js';
import * as scanNew from './scan-new.js';
import * as mailScanNew from './mail-scan-new.js';

export async function render(root, args) {
  const id = args.params[0];
  const schedule = id ? await get(`/api/schedules/${encodeURIComponent(id)}`) : null;
  if (schedule?.purpose === 'retention') {
    // Uma política de retenção aberta pelo endereço dos agendamentos (sem nova entrada no
    // histórico: "Voltar" não cai de novo aqui).
    location.replace(`#/retencao/${schedule.id}`);
    return null;
  }
  const kind = schedule ? schedule.kind : args.query.get('tipo') === 'email' ? 'mail' : 'files';
  const view = kind === 'mail' ? mailScanNew : scanNew;
  return view.render(root, { ...args, props: { ...args.props, scheduling: true, schedule } });
}
