// Servidor IMAP mínimo para os testes: LOGIN, LIST, EXAMINE/SELECT, UID SEARCH, UID FETCH (com
// busca parcial BODY.PEEK[]<0.N>), STATUS e LOGOUT. Suficiente para o conector do CLEAN.
import net from 'node:net';

function imapDate(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const p = (n) => String(n).padStart(2, '0');
  return `${p(date.getUTCDate())}-${months[date.getUTCMonth()]}-${date.getUTCFullYear()} ${p(date.getUTCHours())}:${p(date.getUTCMinutes())}:${p(date.getUTCSeconds())} +0000`;
}

function parseSet(set, max) {
  const out = new Set();
  for (const part of set.split(',')) {
    const [a, b] = part.split(':');
    const from = a === '*' ? max : Number(a);
    const to = b === undefined ? from : b === '*' ? max : Number(b);
    for (let i = Math.min(from, to); i <= Math.max(from, to); i++) out.add(i);
  }
  return out;
}

/** Lê os argumentos de um comando: átomos, "strings" e literais {n}. */
function tokenize(text) {
  const out = [];
  let i = 0;
  while (i < text.length) {
    if (text[i] === ' ') {
      i++;
    } else if (text[i] === '"') {
      let value = '';
      i++;
      while (i < text.length && text[i] !== '"') {
        if (text[i] === '\\') i++;
        value += text[i++];
      }
      i++;
      out.push(value);
    } else if (text[i] === '(') {
      const end = text.indexOf(')', i);
      out.push(text.slice(i, end + 1));
      i = end + 1;
    } else {
      let j = i;
      while (j < text.length && text[j] !== ' ') j++;
      out.push(text.slice(i, j));
      i = j;
    }
  }
  return out;
}

/**
 * accounts: { login: { password, folders: { 'INBOX': [{ raw: Buffer, date: Date }], ... },
 *             flags?: { folder: '\\Trash' } } }
 */
export function startFakeImap(accounts, { log = [], maxLine = Infinity, sizeOffset = 0 } = {}) {
  const server = net.createServer((socket) => {
    let buffer = Buffer.alloc(0);
    let user = null;
    let selected = null;
    let pendingLiteral = null; // { size, line }
    const send = (text) => socket.write(text);
    send('* OK [CAPABILITY IMAP4rev1 UIDPLUS] Fake IMAP pronto\r\n');

    const handle = (line, literals) => {
      const space = line.indexOf(' ');
      const tag = line.slice(0, space);
      // Como o Exchange (MaxCommandSize = 10240): comandos longos demais são recusados.
      if (line.length > maxLine) return socket.write(`${tag} BAD Command Error. 10\r\n`);
      const rest = line.slice(space + 1);
      const args = tokenize(rest);
      let command = (args.shift() || '').toUpperCase();
      let uidMode = false;
      if (command === 'UID') {
        uidMode = true;
        command = (args.shift() || '').toUpperCase();
      }
      log.push(`${command}${uidMode ? ' (UID)' : ''}`);
      const ok = (text = 'OK') => send(`${tag} OK ${text}\r\n`);
      const folders = user ? accounts[user].folders : {};
      const messages = selected ? folders[selected] : [];
      switch (command) {
        case 'CAPABILITY':
          send('* CAPABILITY IMAP4rev1 UIDPLUS\r\n');
          return ok();
        case 'LOGIN': {
          const [login, password] = literals.length ? literals : args;
          if (!accounts[login] || accounts[login].password !== password) return send(`${tag} NO [AUTHENTICATIONFAILED] Credenciais inválidas\r\n`);
          user = login;
          return ok('[CAPABILITY IMAP4rev1 UIDPLUS] Logado');
        }
        case 'LIST':
        case 'LSUB': {
          // LIST "" "" pergunta apenas o separador de pastas.
          if (args[1] === '') {
            send(`* ${command} (\\Noselect) "/" ""\r\n`);
            return ok();
          }
          for (const name of Object.keys(folders)) {
            const special = accounts[user].special?.[name] || '';
            send(`* ${command} (\\HasNoChildren${special ? ` ${special}` : ''}) "/" "${name}"\r\n`);
          }
          return ok();
        }
        case 'SELECT':
        case 'EXAMINE': {
          const name = args[0];
          if (!folders[name]) return send(`${tag} NO Pasta inexistente\r\n`);
          selected = name;
          const list = folders[name];
          send(`* FLAGS (\\Seen \\Answered)\r\n* ${list.length} EXISTS\r\n* 0 RECENT\r\n* OK [UIDVALIDITY 7] ok\r\n* OK [UIDNEXT ${list.length + 1}] ok\r\n`);
          return ok(command === 'EXAMINE' ? '[READ-ONLY] ok' : '[READ-WRITE] ok');
        }
        case 'STATUS': {
          const list = folders[args[0]] || [];
          send(`* STATUS "${args[0]}" (MESSAGES ${list.length})\r\n`);
          return ok();
        }
        case 'SEARCH': {
          const sinceIndex = args.findIndex((a) => a.toUpperCase() === 'SINCE');
          const since = sinceIndex >= 0 ? new Date(`${args[sinceIndex + 1]} 00:00:00 UTC`) : null;
          const uids = messages.map((m, i) => ({ uid: i + 1, date: m.date })).filter((m) => !since || m.date >= since).map((m) => m.uid);
          send(`* SEARCH${uids.map((u) => ` ${u}`).join('')}\r\n`);
          return ok();
        }
        case 'FETCH': {
          const set = parseSet(args[0], messages.length);
          const items = args.slice(1).join(' ').toUpperCase();
          for (let i = 0; i < messages.length; i++) {
            const uid = i + 1;
            if (!set.has(uid)) continue;
            const m = messages[i];
            const parts = [`UID ${uid}`];
            // sizeOffset simula o tamanho estimado do Exchange (EnableExactRFC822Size = false).
            if (items.includes('RFC822.SIZE')) parts.push(`RFC822.SIZE ${m.raw.length + sizeOffset}`);
            if (items.includes('INTERNALDATE')) parts.push(`INTERNALDATE "${imapDate(m.date)}"`);
            const partial = /BODY\.PEEK\[\]<(\d+)\.(\d+)>/.exec(items);
            const head = `* ${uid} FETCH (${parts.join(' ')}`;
            if (partial) {
              const data = m.raw.subarray(Number(partial[1]), Number(partial[1]) + Number(partial[2]));
              socket.write(`${head} BODY[]<${partial[1]}> {${data.length}}\r\n`);
              socket.write(data);
              send(')\r\n');
            } else if (items.includes('BODY.PEEK[]')) {
              socket.write(`${head} BODY[] {${m.raw.length}}\r\n`);
              socket.write(m.raw);
              send(')\r\n');
            } else {
              send(`${head})\r\n`);
            }
          }
          return ok();
        }
        case 'LOGOUT':
          send('* BYE Até logo\r\n');
          ok();
          return socket.end();
        default:
          // NOOP, CLOSE, UNSELECT, ENABLE, ID...
          if (command === 'CLOSE' || command === 'UNSELECT') selected = null;
          return ok();
      }
    };

    let literals = [];
    let lineParts = '';
    socket.on('data', (chunk) => {
      buffer = Buffer.concat([buffer, chunk]);
      for (;;) {
        if (pendingLiteral) {
          if (buffer.length < pendingLiteral.size) return;
          literals.push(buffer.subarray(0, pendingLiteral.size).toString('utf8'));
          buffer = buffer.subarray(pendingLiteral.size);
          pendingLiteral = null;
          continue;
        }
        const idx = buffer.indexOf('\r\n');
        if (idx === -1) return;
        const line = buffer.subarray(0, idx).toString('utf8');
        buffer = buffer.subarray(idx + 2);
        const literal = /\{(\d+)\+?\}$/.exec(line);
        if (literal) {
          lineParts += line.slice(0, literal.index);
          pendingLiteral = { size: Number(literal[1]) };
          if (!line.endsWith('+}')) send('+ pode enviar\r\n');
          continue;
        }
        const full = lineParts + line;
        lineParts = '';
        const lits = literals;
        literals = [];
        handle(full, lits);
      }
    });
    socket.on('error', () => {});
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port, log, close: () => new Promise((r) => server.close(r)) }));
  });
}
