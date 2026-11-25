// Função "elaborar": é o servidor do RDO Digital. Faz três coisas:
//  1) (sem "acao")            pede ao Claude a descrição técnica da atividade;
//  2) acao = "criar-empresa"  cria uma empresa nova + o administrador dela (exige código de convite);
//  3) acao = "usuarios"       o administrador da obra lista/cria/edita/remove os usuários dela
//                             (perfis: admin = administrador, apr = aprovador, enc = elaborador, vis = visualizador);
//  3b) acao = "criar-obra"    o administrador cria uma obra nova (a mesma pessoa passa a ter acesso às duas);
//  3c0) acao = "respostas"  lê a caixa de e-mail (IMAP de qualquer provedor, ou Microsoft 365 pelo Graph) e guarda no Drive, na pasta Respondidos, as respostas do cliente que citam o nº do RDO
//  3c1) acao = "fotos-drive" guarda no Drive (07. Fotos da obra / MM-AAAA / DD-MM-AAAA) as fotos do dia que estão em Fotografias;
//  3c) acao = "enviar-email"  manda o PDF (ao cliente) ou o Excel (ao coordenador) por e-mail (SMTP: Outlook/Gmail, ou Resend);
//  3d) acao = "destravar"     libera um RDO travado (já enviado ao cliente) com a senha de destravamento;
//  4) acao = "excluir-conta"   o usuário apaga a própria conta (se for o único da empresa, apaga a empresa toda);
//  5) acao = "excluir-empresa" o administrador apaga a empresa, todos os usuários, dados e fotos.
// As chaves (Anthropic e de serviço do Supabase) ficam só aqui no servidor, nunca no celular.
import { createClient } from "npm:@supabase/supabase-js@2";
import { Buffer } from "node:buffer";
import nodemailer from "npm:nodemailer@6.9.16";
import CFB from "npm:cfb@1.2.2";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};
const json = (o: unknown, status = 200) =>
  new Response(JSON.stringify(o), { status, headers: { ...cors, "Content-Type": "application/json" } });
const str = (v: unknown, max = 200) => (typeof v === "string" ? v.trim().slice(0, max) : "");
// Contas de e-mail de cada pessoa (segredo CONTAS_EMAIL): {"login@do.app": "senha"} ou {"login@do.app": {"user": "caixa@x", "pass": "senha"}}
function contasEmail(): { login: string; user: string; pass: string }[] {
  let o: any = {};
  try { o = JSON.parse(Deno.env.get("CONTAS_EMAIL") || "{}"); } catch { /* JSON inválido: ignora */ }
  const r: { login: string; user: string; pass: string }[] = [];
  for (const [login, v] of Object.entries(o ?? {})) {
    const user = typeof v === "string" ? login : String((v as any)?.user || login);
    const pass = typeof v === "string" ? v : String((v as any)?.pass || "");
    if (user && pass) r.push({ login: login.toLowerCase(), user, pass });
  }
  return r;
}
const emailOk = (e: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
const PERFIS = ["admin", "apr", "enc", "vis"];
const perfilDe = (b: any, padrao = "enc") =>
  typeof b.perfil === "string" && PERFIS.includes(b.perfil) ? b.perfil : (b.admin === true ? "admin" : padrao);


// ---------- Google Drive (opcional): salva o PDF/Excel na pasta do mês e o e-mail enviado ao cliente ----------
// Dois modos (use um): OAuth (GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN, DRIVE_PASTA_DIARIO, DRIVE_PASTA_EMAIL)
// ou Apps Script (DRIVE_SCRIPT_URL, DRIVE_SCRIPT_TOKEN; veja apps-script/Code.gs).
const modoScript = () => !!(Deno.env.get("DRIVE_SCRIPT_URL") && Deno.env.get("DRIVE_SCRIPT_TOKEN"));
const modoOAuth = () => !!(Deno.env.get("GOOGLE_CLIENT_ID") && Deno.env.get("GOOGLE_CLIENT_SECRET") && Deno.env.get("GOOGLE_REFRESH_TOKEN"));
const driveAtivo = () => modoScript() || modoOAuth();
async function scriptEnvia(corpo: Record<string, unknown>): Promise<string> {
  const r = await fetch(Deno.env.get("DRIVE_SCRIPT_URL")!, {
    method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, redirect: "follow",
    body: JSON.stringify({ ...corpo, token: Deno.env.get("DRIVE_SCRIPT_TOKEN") }),
  });
  const t = await r.text();
  let j: any = {};
  try { j = JSON.parse(t); } catch { /* resposta que não é JSON */ }
  if (!r.ok || !j.ok) throw new Error("Google Drive: " + (j.erro || (r.ok ? "resposta inesperada do script (confira a implantação)" : "erro " + r.status)));
  return String(j.nome || "");
}
async function driveToken(): Promise<string> {
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: Deno.env.get("GOOGLE_CLIENT_ID")!, client_secret: Deno.env.get("GOOGLE_CLIENT_SECRET")!,
      refresh_token: Deno.env.get("GOOGLE_REFRESH_TOKEN")!, grant_type: "refresh_token" }),
  });
  const j: any = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new Error("autorização do Google recusada (" + (j.error_description || j.error || r.status) + ")");
  return j.access_token;
}
const gq = (v: string) => v.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
async function gj(tok: string, url: string, init: any = {}) {
  const r = await fetch(url, { ...init, headers: { Authorization: `Bearer ${tok}`, ...(init.headers || {}) } });
  const j: any = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error("Google Drive: " + (j?.error?.message || r.status));
  return j;
}
async function driveLista(tok: string, pai: string, extra = ""): Promise<any[]> {
  const q = encodeURIComponent(`'${pai}' in parents and trashed=false${extra}`);
  const j = await gj(tok, `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType)&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true`);
  return j.files ?? [];
}
async function driveSubpasta(tok: string, pai: string, nome: string): Promise<string> {
  const ja = await driveLista(tok, pai, ` and name='${gq(nome)}' and mimeType='application/vnd.google-apps.folder'`);
  if (ja.length) return ja[0].id;
  const c = await gj(tok, "https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: nome, mimeType: "application/vnd.google-apps.folder", parents: [pai] }) });
  return c.id;
}
// grava (ou substitui, se já houver um arquivo com o mesmo nome) um arquivo na pasta
async function driveGrava(tok: string, pai: string, nome: string, tipo: string, dados: Uint8Array, modificado?: Date) {
  const ja = await driveLista(tok, pai, ` and name='${gq(nome)}' and mimeType!='application/vnd.google-apps.folder'`);
  const lim = "rdo" + crypto.randomUUID().replace(/-/g, "");
  const mt = modificado && !isNaN(modificado.getTime()) ? { modifiedTime: modificado.toISOString() } : {};
  const meta = ja.length ? { name: nome, ...mt } : { name: nome, parents: [pai], ...mt };
  const cab = new TextEncoder().encode(`--${lim}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${lim}\r\nContent-Type: ${tipo}\r\n\r\n`);
  const fim = new TextEncoder().encode(`\r\n--${lim}--`);
  const corpo = new Uint8Array(cab.length + dados.length + fim.length);
  corpo.set(cab, 0); corpo.set(dados, cab.length); corpo.set(fim, cab.length + dados.length);
  const url = ja.length
    ? `https://www.googleapis.com/upload/drive/v3/files/${ja[0].id}?uploadType=multipart&supportsAllDrives=true&fields=id`
    : "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id";
  return await gj(tok, url, { method: ja.length ? "PATCH" : "POST", headers: { "Content-Type": `multipart/related; boundary=${lim}` }, body: corpo });
}
async function proximoNumero(tok: string, pai: string): Promise<number> {
  const arqs = await driveLista(tok, pai, " and mimeType!='application/vnd.google-apps.folder'");
  let mx = 0;
  for (const a of arqs) { const m = /^(\d{1,4})\s*\./.exec(a.name || ""); if (m) mx = Math.max(mx, Number(m[1])); }
  return mx + 1;
}
// as duas operações usadas pelo envio de e-mail, no modo escolhido
async function driveSalvaArquivo(mes: string, nome: string, mime: string, dados: Uint8Array): Promise<string> {
  if (modoScript()) return await scriptEnvia({ acao: "arquivo", mes, nome, mime, dados: Buffer.from(dados).toString("base64") });
  const raiz = Deno.env.get("DRIVE_PASTA_DIARIO");
  if (!raiz) throw new Error("falta o segredo DRIVE_PASTA_DIARIO");
  const tok = await driveToken();
  await driveGrava(tok, await driveSubpasta(tok, raiz, mes), nome, mime, dados);
  return nome;
}
async function driveSalvaEmail(base: string, dados: Uint8Array, ext = "msg"): Promise<string> {
  const mime = ext === "msg" ? "application/vnd.ms-outlook" : "message/rfc822";
  if (modoScript()) return await scriptEnvia({ acao: "email", nome: base, ext, mime, dados: Buffer.from(dados).toString("base64") });
  const pasta = Deno.env.get("DRIVE_PASTA_EMAIL");
  if (!pasta) throw new Error("falta o segredo DRIVE_PASTA_EMAIL");
  const tok = await driveToken();
  const nome = `${String(await proximoNumero(tok, pasta)).padStart(2, "0")}. ${base}.${ext}`;
  await driveGrava(tok, pasta, nome, mime, dados);
  return nome;
}

// ---------- e-mail em .msg (formato do Outlook, MS-OXMSG) ----------
type MsgAnexo = { name: string; data: Uint8Array; cid?: string; mime?: string };
type MsgDados = { fromName: string; fromEmail: string; to: string[]; cc: string[]; subject: string; text: string; html: string; date: Date; anexos: MsgAnexo[] };
function montaMsg(CFB: any, o: MsgDados): Uint8Array {
  const u16 = (s: string) => Buffer.from(s, "utf16le");
  const u32 = (n: number) => { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; };
  const ft = (d: Date) => { const b = Buffer.alloc(8); b.writeBigUInt64LE(BigInt(d.getTime() + 11644473600000) * 10000n); return b; };
  const hex = (n: number, w: number) => n.toString(16).toUpperCase().padStart(w, "0");
  type Prop = { id: number; tipo: number; v: number | boolean | string | Date | Uint8Array };
  const cfb = CFB.utils.cfb_new();
  const dir = (p: string) => { cfb.FileIndex.push({ name: p.split("/").pop(), type: 1, content: null, size: 0 }); cfb.FullPaths.push(cfb.FullPaths[0] + p + "/"); };
  const arq = (p: string, c: Uint8Array) => CFB.utils.cfb_add(cfb, "/" + p, c);
  // grava as propriedades de um objeto (mensagem, destinatário ou anexo): fluxo de propriedades + um fluxo por valor variável
  const grava = (base: string, props: Prop[], cab: Buffer) => {
    const ent: Buffer[] = [];
    for (const p of props) {
      const tag = u32(((p.id << 16) | p.tipo) >>> 0), flags = u32(6), val = Buffer.alloc(8);
      if (p.tipo === 0x0003) val.writeInt32LE(Number(p.v));
      else if (p.tipo === 0x000b) val.writeUInt16LE(p.v ? 1 : 0);
      else if (p.tipo === 0x0040) ft(p.v as Date).copy(val);
      else {
        const dados: Uint8Array = p.tipo === 0x001f ? u16(String(p.v)) : (p.v as Uint8Array);
        val.writeUInt32LE(dados.length + (p.tipo === 0x001f ? 2 : 0));
        arq(`${base}__substg1.0_${hex(p.id, 4)}${hex(p.tipo, 4)}`, dados);
      }
      ent.push(Buffer.concat([tag, flags, val]));
    }
    arq(`${base}__properties_version1.0`, Buffer.concat([cab, ...ent]));
  };
  const S = (id: number, v: string): Prop => ({ id, tipo: 0x001f, v });
  const L = (id: number, v: number): Prop => ({ id, tipo: 0x0003, v });
  const B = (id: number, v: Uint8Array): Prop => ({ id, tipo: 0x0102, v });
  const T = (id: number, v: Date): Prop => ({ id, tipo: 0x0040, v });
  const nomes = (l: string[]) => l.join("; ");
  const dest = [...o.to.map((e) => [1, e] as const), ...o.cc.map((e) => [2, e] as const)];
  // mensagem
  const cab = Buffer.concat([Buffer.alloc(8), u32(dest.length), u32(o.anexos.length), u32(dest.length), u32(o.anexos.length), Buffer.alloc(8)]);
  const html = Buffer.from(o.html, "utf8");
  grava("", [
    S(0x001a, "IPM.Note"), S(0x0037, o.subject), S(0x0e1d, o.subject),
    S(0x0c1a, o.fromName || o.fromEmail), S(0x0c1e, "SMTP"), S(0x0c1f, o.fromEmail),
    S(0x0042, o.fromName || o.fromEmail), S(0x0064, "SMTP"), S(0x0065, o.fromEmail),
    S(0x0e04, nomes(o.to)), S(0x0e03, nomes(o.cc)),
    S(0x1000, o.text), B(0x1013, html), L(0x3fde, 65001), L(0x3ffd, 65001),
    L(0x0e07, 1), L(0x0e08, html.length + o.text.length), { id: 0x0e1b, tipo: 0x000b, v: o.anexos.length > 0 },
    T(0x0039, o.date), T(0x0e06, o.date), T(0x3007, o.date), T(0x3008, o.date),
    L(0x340d, 0x00040000), L(0x3ff1, 0x0416),
  ], cab);
  dir("__nameid_version1.0");
  for (const n of ["00020102", "00030102", "00040102"]) arq(`__nameid_version1.0/__substg1.0_${n}`, Buffer.alloc(0));
  dest.forEach(([tipo, email], i) => {
    const b = `__recip_version1.0_#${hex(i, 8)}`;
    dir(b);
    grava(`${b}/`, [L(0x0c15, tipo), L(0x3000, i), L(0x0ffe, 6), S(0x3001, email), S(0x3002, "SMTP"), S(0x3003, email), S(0x39fe, email), L(0x3900, 0)], Buffer.alloc(8));
  });
  o.anexos.forEach((a, i) => {
    const b = `__attach_version1.0_#${hex(i, 8)}`;
    dir(b);
    const ext = /\.[A-Za-z0-9]{1,5}$/.exec(a.name)?.[0] ?? "";
    const p: Prop[] = [B(0x3701, a.data), S(0x3704, a.name), S(0x3707, a.name), S(0x3703, ext), L(0x3705, 1), L(0x0ffe, 7), L(0x0e21, i)];
    if (a.mime) p.push(S(0x370e, a.mime));
    if (a.cid) { p.push(S(0x3712, a.cid)); p.push(L(0x3714, 4)); p.push({ id: 0x7ffe, tipo: 0x000b, v: true }); }
    grava(`${b}/`, p, Buffer.alloc(8));
  });
  return new Uint8Array(CFB.write(cfb, { type: "buffer" }));
}

// largura e altura de uma imagem PNG, JPEG ou GIF (para a assinatura não esticar o logo)
function dimImg(b: Uint8Array): { w: number; h: number } | null {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50) return { w: dv.getUint32(16), h: dv.getUint32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49) return { w: dv.getUint16(6, true), h: dv.getUint16(8, true) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const t = b[i + 1];
      if (t >= 0xc0 && t <= 0xcf && t !== 0xc4 && t !== 0xc8 && t !== 0xcc) return { h: dv.getUint16(i + 5), w: dv.getUint16(i + 7) };
      i += 2 + dv.getUint16(i + 2);
    }
  }
  return null;
}

// Apaga todos os arquivos do bucket "fotos" de uma empresa (pasta <empresa>/<data>/<foto>.jpg).
async function listar(admin: any, prefixo: string) {
  const out: any[] = [];
  for (let off = 0; ; off += 1000) {
    const { data } = await admin.storage.from("fotos").list(prefixo, { limit: 1000, offset: off });
    if (!data || !data.length) break;
    out.push(...data);
    if (data.length < 1000) break;
  }
  return out;
}
async function apagarPasta(admin: any, prefixo: string) {
  for (const it of await listar(admin, prefixo)) {
    const caminho = prefixo ? `${prefixo}/${it.name}` : it.name;
    if (it.id) await admin.storage.from("fotos").remove([caminho]);
    else await apagarPasta(admin, caminho);
  }
}
async function apagarEmpresa(admin: any, empresaId: string) {
  const { data: emp } = await admin.from("empresas").select("legado").eq("id", empresaId).maybeSingle();
  await apagarPasta(admin, empresaId);
  if (emp?.legado) {   // empresa migrada do app antigo: fotos antigas ficam em <data>/ na raiz
    for (const it of await listar(admin, "")) if (!it.id && /^\d{4}-\d{2}-\d{2}$/.test(it.name)) await apagarPasta(admin, it.name);
  }
  await admin.from("kv").delete().eq("empresa_id", empresaId);
  // pessoas desta obra: quem tem outra obra continua com a conta; quem só tinha esta é removido
  const ids = new Set<string>();
  for (const m of (await admin.from("acessos").select("user_id").eq("empresa_id", empresaId)).data ?? []) ids.add(m.user_id);
  for (const m of (await admin.from("profiles").select("id").eq("empresa_id", empresaId)).data ?? []) ids.add(m.id);
  await admin.from("acessos").delete().eq("empresa_id", empresaId);
  for (const id of ids) {
    const { data: outros } = await admin.from("acessos").select("empresa_id, role, pode_gerar, so_dia").eq("user_id", id).limit(1);
    if (outros && outros.length) {
      await admin.from("profiles").update({ empresa_id: outros[0].empresa_id, role: outros[0].role,
        pode_gerar: outros[0].pode_gerar, so_dia: outros[0].so_dia }).eq("id", id);
    } else await admin.auth.admin.deleteUser(id);
  }
  await admin.from("empresas").delete().eq("id", empresaId);
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  try {
    const url = Deno.env.get("SUPABASE_URL")!;
    const admin = createClient(url, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, {
      auth: { persistSession: false },
    });
    const body = await req.json().catch(() => ({}));

    // ---------- criar empresa (ainda sem login) ----------
    if (body.acao === "criar-empresa") {
      const codigo = str(body.codigo, 80), empresa = str(body.empresa, 120), nome = str(body.nome, 120);
      const email = str(body.email, 200).toLowerCase(), senha = typeof body.senha === "string" ? body.senha : "";
      if (!codigo || !empresa || !nome || !emailOk(email)) return json({ error: "Preencha todos os campos." }, 400);
      if (senha.length < 6) return json({ error: "A senha precisa ter pelo menos 6 caracteres." }, 400);
      const { data: cv } = await admin.from("convites").select("codigo, usado_em").eq("codigo", codigo).maybeSingle();
      if (!cv) return json({ error: "Código de convite inválido." }, 403);
      if (cv.usado_em) return json({ error: "Este código de convite já foi usado." }, 403);
      // reserva o código antes de criar (evita uso duplo)
      const { data: res } = await admin.from("convites").update({ usado_em: new Date().toISOString() })
        .eq("codigo", codigo).is("usado_em", null).select("codigo");
      if (!res || !res.length) return json({ error: "Este código de convite já foi usado." }, 403);
      const desfazer = () => admin.from("convites").update({ usado_em: null, empresa_id: null }).eq("codigo", codigo);
      const { data: emp, error: e1 } = await admin.from("empresas").insert({ nome: empresa }).select("id").single();
      if (e1 || !emp) { await desfazer(); return json({ error: "Não foi possível criar a empresa." }, 500); }
      const { data: u, error: e2 } = await admin.auth.admin.createUser({
        email, password: senha, email_confirm: true, user_metadata: { nome },
      });
      if (e2 || !u?.user) {
        await admin.from("empresas").delete().eq("id", emp.id); await desfazer();
        const ja = /already|registered|exists/i.test(e2?.message ?? "");
        return json({ error: ja ? "Este e-mail já tem cadastro." : "Não foi possível criar o usuário." }, ja ? 409 : 500);
      }
      await admin.from("profiles").upsert({ id: u.user.id, nome, role: "admin", pode_gerar: true, so_dia: false, empresa_id: emp.id });
      await admin.from("empresas").update({ grupo_id: emp.id }).eq("id", emp.id);
      await admin.from("acessos").upsert({ user_id: u.user.id, empresa_id: emp.id, role: "admin", pode_gerar: true, so_dia: false });
      await admin.from("convites").update({ empresa_id: emp.id }).eq("codigo", codigo);
      return json({ ok: true });
    }

    // ---------- daqui para baixo, só usuário logado ----------
    const sb = createClient(url, Deno.env.get("SUPABASE_ANON_KEY")!, {
      global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } },
    });
    const { data, error } = await sb.auth.getUser();
    if (error || !data?.user) return json({ error: "não autorizado" }, 401);
    const { data: eu } = await admin.from("profiles").select("id, role, empresa_id").eq("id", data.user.id).single();
    if (!eu?.empresa_id) return json({ error: "sem empresa" }, 403);

    // ---------- excluir a própria conta ----------
    if (body.acao === "excluir-conta") {
      const { data: ps } = await admin.from("acessos").select("user_id, role").eq("empresa_id", eu.empresa_id);
      const todos = ps ?? [], admins = todos.filter((x: any) => x.role === "admin");
      if (todos.length > 1 && eu.role === "admin" && admins.length <= 1) {
        return json({ error: "Você é o único administrador. Torne outra pessoa administradora antes, ou use \"Excluir a empresa\"." }, 400);
      }
      const { data: minhas } = await admin.from("acessos").select("empresa_id").eq("user_id", eu.id);
      await admin.auth.admin.deleteUser(eu.id);
      for (const o of minhas ?? []) {   // obra que ficou sem ninguém é apagada
        const { count } = await admin.from("acessos").select("user_id", { count: "exact", head: true }).eq("empresa_id", o.empresa_id);
        if (!count) await apagarEmpresa(admin, o.empresa_id);
      }
      return json({ ok: true, empresa: todos.length <= 1 });
    }

    // ---------- excluir a empresa inteira (somente administrador) ----------
    if (body.acao === "excluir-empresa") {
      if (eu.role !== "admin") return json({ error: "Somente o administrador pode excluir a empresa." }, 403);
      if (body.confirmar !== true) return json({ error: "confirmação ausente" }, 400);
      await apagarEmpresa(admin, eu.empresa_id);
      return json({ ok: true, empresa: true });
    }

    // ---------- nova obra (somente administrador) ----------
    if (body.acao === "criar-obra") {
      if (eu.role !== "admin") return json({ error: "Somente o administrador pode criar obras." }, 403);
      const nome = str(body.nome, 120);
      if (!nome) return json({ error: "Informe o nome da obra." }, 400);
      const { count: tot } = await admin.from("acessos").select("empresa_id", { count: "exact", head: true }).eq("user_id", eu.id);
      if ((tot ?? 0) >= 60) return json({ error: "Limite de obras atingido." }, 400);
      const { data: cur } = await admin.from("empresas").select("grupo_id").eq("id", eu.empresa_id).maybeSingle();
      const { data: nova, error: e1 } = await admin.from("empresas")
        .insert({ nome, grupo_id: cur?.grupo_id ?? eu.empresa_id }).select("id").single();
      if (e1 || !nova) return json({ error: "Não foi possível criar a obra." }, 500);
      await admin.from("acessos").insert({ user_id: eu.id, empresa_id: nova.id, role: "admin", pode_gerar: true, so_dia: false });
      // aproveita da obra atual só o básico (empresa, logo da contratada e quadros de funções/equipamentos)
      const { data: c } = await admin.from("kv").select("value").eq("empresa_id", eu.empresa_id).eq("key", "cfg").eq("deleted", false).maybeSingle();
      if (c?.value) {
        const v = c.value as any, base: Record<string, unknown> = {};
        for (const k of ["contratada", "logoPot", "efetivoCat", "equipCat", "disc", "foto", "prefixo", "funcTipo"]) if (v[k] !== undefined) base[k] = v[k];
        await admin.from("kv").insert({ empresa_id: nova.id, key: "cfg", value: { ...base, obra: nome, obraCurta: nome, configurado: false }, autor: "" });
      }
      return json({ ok: true, id: nova.id });
    }

    // ---------- usuários da obra (somente administrador) ----------
    if (body.acao === "usuarios") {
      if (eu.role !== "admin") return json({ error: "Somente o administrador pode gerenciar usuários." }, 403);
      const op = body.op;
      const alvo = async (id: unknown) => {
        const { data: a } = await admin.from("acessos").select("user_id, role").eq("user_id", str(id, 60)).eq("empresa_id", eu.empresa_id).maybeSingle();
        return a ? { id: a.user_id as string, role: a.role as string } : null;
      };
      const outrasObras = async (id: string) => {
        const { count } = await admin.from("acessos").select("empresa_id", { count: "exact", head: true }).eq("user_id", id).neq("empresa_id", eu.empresa_id);
        return count ?? 0;
      };
      const ajustar = async (id: string, up: Record<string, unknown>) => {
        await admin.from("acessos").update(up).eq("user_id", id).eq("empresa_id", eu.empresa_id);
        await admin.from("profiles").update(up).eq("id", id).eq("empresa_id", eu.empresa_id);   // se a obra aberta dela é esta
      };
      if (op === "listar") {
        const { data: ac } = await admin.from("acessos").select("user_id, role, pode_gerar, so_dia").eq("empresa_id", eu.empresa_id);
        const lista = await Promise.all((ac ?? []).map(async (a: any) => {
          const { data: p } = await admin.from("profiles").select("nome").eq("id", a.user_id).maybeSingle();
          const { data: au } = await admin.auth.admin.getUserById(a.user_id);
          return { id: a.user_id, nome: p?.nome ?? "", role: a.role, pode_gerar: a.pode_gerar, so_dia: a.so_dia,
            email: au?.user?.email ?? "", outras: await outrasObras(a.user_id) };
        }));
        lista.sort((x: any, y: any) => String(x.nome).localeCompare(String(y.nome), "pt-BR"));
        return json({ ok: true, usuarios: lista });
      }
      if (op === "criar") {
        const nome = str(body.nome, 120), email = str(body.email, 200).toLowerCase();
        const senha = typeof body.senha === "string" ? body.senha : "";
        const perfil = perfilDe(body);
        if (!emailOk(email)) return json({ error: "Informe um e-mail válido." }, 400);
        const flags = { pode_gerar: perfil === "enc" ? !!body.pode_gerar : true, so_dia: perfil === "enc" ? !!body.so_dia : false };
        // e-mail que já existe: só dá acesso a esta obra (a senha continua sendo da pessoa)
        const { data: ja } = await admin.rpc("uid_por_email", { e: email });
        if (ja) {
          const { data: ex } = await admin.from("acessos").select("user_id").eq("user_id", ja).eq("empresa_id", eu.empresa_id).maybeSingle();
          if (ex) return json({ error: "Esta pessoa já tem acesso a esta obra." }, 409);
          await admin.from("acessos").insert({ user_id: ja, empresa_id: eu.empresa_id, role: perfil, ...flags });
          return json({ ok: true, existente: true });
        }
        if (!nome) return json({ error: "Informe o nome." }, 400);
        if (senha.length < 6) return json({ error: "A senha precisa ter pelo menos 6 caracteres." }, 400);
        const { data: u, error: e } = await admin.auth.admin.createUser({
          email, password: senha, email_confirm: true, user_metadata: { nome },
        });
        if (e || !u?.user) {
          const dup = /already|registered|exists/i.test(e?.message ?? "");
          return json({ error: dup ? "Este e-mail já tem cadastro." : "Não foi possível criar o usuário." }, dup ? 409 : 500);
        }
        await admin.from("profiles").upsert({ id: u.user.id, nome, role: perfil, ...flags, empresa_id: eu.empresa_id });
        await admin.from("acessos").upsert({ user_id: u.user.id, empresa_id: eu.empresa_id, role: perfil, ...flags });
        return json({ ok: true });
      }
      const p = await alvo(body.id);
      if (!p) return json({ error: "Usuário não encontrado." }, 404);
      if (op === "atualizar") {
        const up: Record<string, unknown> = {};
        if (typeof body.pode_gerar === "boolean") up.pode_gerar = body.pode_gerar;
        if (typeof body.so_dia === "boolean") up.so_dia = body.so_dia;
        const novo = typeof body.perfil === "string" && PERFIS.includes(body.perfil) ? body.perfil
          : typeof body.admin === "boolean" ? (body.admin ? "admin" : "enc") : null;
        if (novo) {
          if (novo !== "admin" && p.role === "admin") {
            const { count } = await admin.from("acessos").select("user_id", { count: "exact", head: true })
              .eq("empresa_id", eu.empresa_id).eq("role", "admin");
            if ((count ?? 0) <= 1) return json({ error: "A obra precisa ter pelo menos um administrador." }, 400);
          }
          up.role = novo;
          if (novo !== "enc") up.so_dia = false;
          if (novo === "admin" || novo === "apr") up.pode_gerar = true;
        }
        if (Object.keys(up).length) await ajustar(p.id, up);
        if (typeof body.nome === "string" && body.nome.trim() && !(await outrasObras(p.id))) {
          await admin.from("profiles").update({ nome: str(body.nome, 120) }).eq("id", p.id);
        }
        return json({ ok: true });
      }
      if (op === "senha") {
        if (await outrasObras(p.id)) return json({ error: "Esta pessoa também usa outras obras. Ela mesma troca a senha em \"Minha conta\"." }, 400);
        const senha = typeof body.senha === "string" ? body.senha : "";
        if (senha.length < 6) return json({ error: "A senha precisa ter pelo menos 6 caracteres." }, 400);
        const { error: e } = await admin.auth.admin.updateUserById(p.id, { password: senha });
        return e ? json({ error: "Não foi possível trocar a senha." }, 500) : json({ ok: true });
      }
      if (op === "remover") {
        if (p.id === eu.id) return json({ error: "Você não pode remover o seu próprio usuário." }, 400);
        if (p.role === "admin") {
          const { count } = await admin.from("acessos").select("user_id", { count: "exact", head: true })
            .eq("empresa_id", eu.empresa_id).eq("role", "admin");
          if ((count ?? 0) <= 1) return json({ error: "A obra precisa ter pelo menos um administrador." }, 400);
        }
        const { data: outros } = await admin.from("acessos").select("empresa_id, role, pode_gerar, so_dia")
          .eq("user_id", p.id).neq("empresa_id", eu.empresa_id).limit(1);
        if (outros && outros.length) {   // tem outra obra: só tira o acesso a esta
          await admin.from("acessos").delete().eq("user_id", p.id).eq("empresa_id", eu.empresa_id);
          await admin.from("profiles").update({ empresa_id: outros[0].empresa_id, role: outros[0].role,
            pode_gerar: outros[0].pode_gerar, so_dia: outros[0].so_dia }).eq("id", p.id).eq("empresa_id", eu.empresa_id);
        } else await admin.auth.admin.deleteUser(p.id);
        return json({ ok: true });
      }
      return json({ error: "operação inválida" }, 400);
    }

    // ---------- preferências de cada usuário: assinatura de e-mail (get/set) e imagem do carimbo de assinatura (carimbo-get/carimbo-set) ----------
    // Ficam na tabela usr_prefs (uma linha por usuário, só o servidor lê e grava).
    if (body.acao === "assinatura") {
      const op = body.op;
      const { data: pr, error: ep } = await admin.from("usr_prefs").select("sig, carimbo").eq("user_id", eu.id).maybeSingle();
      if (ep) return json({ error: "Não consegui acessar as preferências: " + ep.message }, 500);
      const grava = async (campos: Record<string, unknown>) => {
        const { error } = await admin.from("usr_prefs").upsert({ user_id: eu.id, empresa_id: eu.empresa_id, ...campos, updated_at: new Date().toISOString() }, { onConflict: "user_id" });
        if (error) throw new Error(error.message);
      };
      try {
        if (op === "get") return json({ ok: true, sig: pr?.sig ?? null });
        if (op === "set") {
          const g = body.sig ?? {};
          const sig = { nome: str(g.nome, 80), cargo: str(g.cargo, 80), fone: str(g.fone, 40), email: str(g.email, 120), site: str(g.site, 120) };
          await grava({ sig }); return json({ ok: true });
        }
        if (op === "carimbo-get") {
          const c: any = pr?.carimbo ?? null;
          if (c && c.removed) return json({ ok: true, removed: true });
          return json({ ok: true, ass: c && c.img ? c : null });
        }
        if (op === "carimbo-set") {
          const a = body.ass;
          if (a === null || a === undefined) { await grava({ carimbo: { removed: true } }); return json({ ok: true }); }
          if (typeof a !== "object" || typeof a.img !== "string" || !/^data:image\//.test(a.img) || a.img.length > 600000) return json({ error: "imagem de assinatura inválida ou grande demais" }, 400);
          await grava({ carimbo: { img: a.img, on: a.on !== false, sz: ["p", "m", "g"].includes(a.sz) ? a.sz : "m" } }); return json({ ok: true });
        }
      } catch (e) { return json({ error: "Não consegui salvar: " + String((e as any)?.message ?? e).slice(0, 200) }, 500); }
      return json({ error: "pedido inválido" }, 400);
    }

    // ---------- destravar um RDO já enviado ao cliente (exige a senha de destravamento) ----------
    if (body.acao === "destravar") {
      const dia = str(body.data, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dia)) return json({ error: "pedido inválido" }, 400);
      const { data: ac } = await admin.from("acessos").select("role").eq("user_id", eu.id).eq("empresa_id", eu.empresa_id).maybeSingle();
      if (!ac || (ac.role !== "admin" && ac.role !== "apr")) return json({ error: "Só o administrador e o aprovador podem destravar um RDO." }, 403);
      const senha = Deno.env.get("DESTRAVA_SENHA") || "potctr";
      if (typeof body.senha !== "string" || body.senha !== senha) return json({ error: "Senha incorreta." }, 403);
      const chave = `${dia}/apr`;
      const { data: row } = await admin.from("kv").select("value").eq("empresa_id", eu.empresa_id).eq("key", chave).eq("deleted", false).maybeSingle();
      const ap: any = row?.value;
      if (!ap || ap.st !== "cliente") return json({ error: "Este RDO não está travado." }, 400);
      const nome = str(data.user.user_metadata?.nome, 120) || eu.id;
      const v = (Number(ap.v) || 1) + 1;
      const novo = { ...ap, st: "", v, ap: null, cl: null, logos: null, mail: null, mot: "",
        hist: [...(Array.isArray(ap.hist) ? ap.hist : []), { t: Date.now(), by: nome, ev: "destravado", v }] };
      const { error: eg } = await admin.from("kv").upsert({ empresa_id: eu.empresa_id, key: chave, value: novo, deleted: false, autor: eu.id, updated_at: new Date().toISOString() },
        { onConflict: "empresa_id,key" });
      if (eg) return json({ error: "Não consegui destravar: " + eg.message }, 500);
      return json({ ok: true });
    }

    // ---------- buscar no e-mail (Microsoft 365) as respostas do cliente e guardar no Drive ----------
    // Segredos: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET (app registrado no Azure com a permissão Mail.Read, tipo "Aplicativo"),
    // MS_MAILBOX (a caixa que recebe as respostas; se faltar, usa MAIL_FROM/SMTP_USER) e DRIVE_PASTA_RESPONDIDOS (pasta "Respondidos Diase").
    if (body.acao === "respostas") {
      const { data: ac } = await admin.from("acessos").select("role").eq("user_id", eu.id).eq("empresa_id", eu.empresa_id).maybeSingle();
      if (!ac || (ac.role !== "admin" && ac.role !== "apr")) return json({ error: "Só o administrador e o aprovador podem fazer isso." }, 403);
      const tid = Deno.env.get("MS_TENANT_ID"), cid = Deno.env.get("MS_CLIENT_ID"), sec = Deno.env.get("MS_CLIENT_SECRET");
      const caixa = Deno.env.get("MS_MAILBOX") || Deno.env.get("MAIL_FROM") || Deno.env.get("SMTP_USER") || "";
      const pastaResp = Deno.env.get("DRIVE_PASTA_RESPONDIDOS"), pastaEnv = Deno.env.get("DRIVE_PASTA_EMAIL");
      const empDrive = Deno.env.get("DRIVE_EMPRESA_ID");
      // Dois jeitos de ler a caixa: Microsoft 365 (MS_*) ou IMAP de qualquer provedor (usa IMAP_HOST/IMAP_USER/IMAP_PASS; se faltarem, usa os dados do SMTP de envio)
      const smtpH = Deno.env.get("SMTP_HOST") || "";
      const imapHost = Deno.env.get("IMAP_HOST") || smtpH.replace(/^smtp\./i, "imap.");
      const imapUser = Deno.env.get("IMAP_USER") || Deno.env.get("SMTP_USER") || "";
      const imapPass = Deno.env.get("IMAP_PASS") || Deno.env.get("SMTP_PASS") || "";
      const imapPorta = Number(Deno.env.get("IMAP_PORT") || "993");
      const usaMs = !!(tid && cid && sec && caixa);
      const caixasImap: { user: string; pass: string }[] = [];
      for (const c of contasEmail()) if (!caixasImap.some((x) => x.user.toLowerCase() === c.user.toLowerCase())) caixasImap.push({ user: c.user, pass: c.pass });
      if (imapUser && imapPass && !caixasImap.some((x) => x.user.toLowerCase() === imapUser.toLowerCase())) caixasImap.push({ user: imapUser, pass: imapPass });
      const usaImap = !usaMs && !!imapHost && caixasImap.length > 0;
      if (!usaMs && !usaImap) return json({ ok: true, config: false, aviso: "A leitura do e-mail ainda não foi configurada no servidor (CONTAS_EMAIL, ou IMAP_HOST/IMAP_USER/IMAP_PASS, ou MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET)." });
      if (!modoOAuth() || !pastaResp) return json({ ok: true, config: false, aviso: "Falta o Google Drive por OAuth e o segredo DRIVE_PASTA_RESPONDIDOS para guardar as respostas." });
      if (empDrive && empDrive !== eu.empresa_id) return json({ ok: true, config: false });
      const rd = (n: number) => String(n).padStart(2, "0");
      try {
        const { data: kc } = await admin.from("kv").select("value").eq("empresa_id", eu.empresa_id).eq("key", "cfg").maybeSingle();
        const cfg: any = kc?.value ?? {};
        const dominios = new Set(String(cfg.emailCli ?? "").split(/[;,\s]+/).map((e: string) => e.trim().toLowerCase()).filter(emailOk).map((e: string) => e.split("@")[1]));
        const enderecos = new Set(String(cfg.emailCli ?? "").split(/[;,\s]+/).map((e: string) => e.trim().toLowerCase()).filter(emailOk));
        if (!dominios.size) return json({ ok: true, config: true, salvos: [], aviso: "Cadastre o e-mail do cliente em Configurações para eu reconhecer as respostas." });
        const ini = /^\d{4}-\d{2}-\d{2}$/.test(String(cfg.inicio ?? "")) ? new Date(cfg.inicio + "T12:00:00Z") : null;
        // já guardadas (para não repetir)
        const { data: kr } = await admin.from("kv").select("value").eq("empresa_id", eu.empresa_id).eq("key", "resp-reg").eq("deleted", false).maybeSingle();
        const reg: any = kr?.value && typeof kr.value === "object" ? kr.value : { ids: {} };
        reg.ids = reg.ids || {};
        type Carga = { fromName: string; fromEmail: string; to: string[]; cc: string[]; subject: string; html: string; texto: string; date: Date; anexos: MsgAnexo[]; raw: () => Promise<Uint8Array> };
        type Achado = { id: string; nro: number; carregar: () => Promise<Carga> };
        const achados: Achado[] = [];
        let verificados = 0, avisoCaixas = "";
        const rxRdo = /(?:RDO|di[aá]rio\s+de\s+obra)[^\d\n]{0,25}(\d{2,4})/i;
        const doCliente = (de: string) => !!de && (enderecos.has(de) || dominios.has(de.split("@")[1]));
        if (usaMs) {
          // token da Microsoft
          const rt = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tid!)}/oauth2/v2.0/token`, {
            method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({ client_id: cid!, client_secret: sec!, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }) });
          const tj: any = await rt.json().catch(() => ({}));
          if (!rt.ok || !tj.access_token) throw new Error("a Microsoft recusou o acesso (" + (tj.error_description || tj.error || rt.status) + ")");
          const g = async (u: string, raw = false) => {
            const r = await fetch(u, { headers: { Authorization: "Bearer " + tj.access_token } });
            if (!r.ok) { const t: any = await r.json().catch(() => ({})); throw new Error("Microsoft Graph: " + (t?.error?.message || r.status)); }
            return raw ? new Uint8Array(await r.arrayBuffer()) : await r.json();
          };
          const u0 = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(caixa)}`;
          const desde = new Date(Date.now() - 60 * 864e5).toISOString();
          const lista: any = await g(`${u0}/mailFolders/inbox/messages?$top=100&$orderby=receivedDateTime desc&$filter=${encodeURIComponent("receivedDateTime ge " + desde)}&$select=id,subject,from,receivedDateTime,bodyPreview`);
          verificados = (lista.value ?? []).length;
          for (const m of (lista.value ?? [])) {
            if (reg.ids[m.id]) continue;
            const de = String(m.from?.emailAddress?.address ?? "").toLowerCase();
            if (!doCliente(de)) continue;
            const ass = String(m.subject ?? ""), txt = ass + " " + String(m.bodyPreview ?? "");
            const mm = rxRdo.exec(ass) || rxRdo.exec(txt);
            if (!mm) continue;
            achados.push({ id: m.id, nro: Number(mm[1]), carregar: async () => {
              const full: any = await g(`${u0}/messages/${encodeURIComponent(m.id)}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body`);
              let anexos: any[] = [];
              try { anexos = ((await g(`${u0}/messages/${encodeURIComponent(m.id)}/attachments?$top=50`)) as any).value ?? []; } catch { /* sem anexos legíveis */ }
              const nomes = (l: any[]) => (l ?? []).map((x: any) => String(x?.emailAddress?.address ?? "")).filter(emailOk);
              const html = full.body?.contentType === "html" ? String(full.body.content ?? "") : "";
              const texto = full.body?.contentType === "html" ? html.replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/\s+/g, " ").trim() : String(full.body?.content ?? "");
              return { fromName: String(full.from?.emailAddress?.name ?? ""), fromEmail: String(full.from?.emailAddress?.address ?? ""),
                to: nomes(full.toRecipients), cc: nomes(full.ccRecipients), subject: String(full.subject ?? ""), html, texto, date: new Date(full.receivedDateTime),
                anexos: anexos.filter((x: any) => x["@odata.type"] === "#microsoft.graph.fileAttachment" && x.contentBytes).map((x: any) =>
                  ({ name: String(x.name ?? "anexo"), data: new Uint8Array(Buffer.from(x.contentBytes, "base64")), cid: x.isInline ? String(x.contentId ?? "") || undefined : undefined, mime: String(x.contentType ?? "application/octet-stream") })),
                raw: async () => await g(`${u0}/messages/${encodeURIComponent(m.id)}/$value`, true) as Uint8Array };
            } });
          }
        } else {
          // IMAP (qualquer provedor). As bibliotecas só são carregadas aqui, para não pesar nas outras ações.
          const { ImapFlow } = await import("npm:imapflow@1.0.164");
          const { simpleParser } = await import("npm:mailparser@3.7.1");
          const vistos = new Set<string>(); const errosCaixa: string[] = [];
          for (const cx of caixasImap) {
            try {
            const cliente: any = new ImapFlow({ host: imapHost, port: imapPorta, secure: imapPorta === 993, auth: { user: cx.user, pass: cx.pass }, logger: false });
            cliente.on("error", () => { /* erros de conexão aparecem nas chamadas abaixo */ });
            await cliente.connect();
            const trava = await cliente.getMailboxLock("INBOX");
            try {
              const uids: number[] = (await cliente.search({ since: new Date(Date.now() - 60 * 864e5) }, { uid: true })) || [];
              const ultimos = uids.slice(-300);
              const cand: { uid: number; id: string; subject: string }[] = [];
              if (ultimos.length) {
                for await (const msg of cliente.fetch(ultimos, { uid: true, envelope: true }, { uid: true })) {
                  verificados++;
                  const env: any = msg.envelope || {};
                  const id = String(env.messageId || "uid" + msg.uid);
                  if (reg.ids[id] || vistos.has(id)) continue;
                  const de = String(env.from?.[0]?.address ?? "").toLowerCase();
                  if (!doCliente(de)) continue;
                  cand.push({ uid: msg.uid, id, subject: String(env.subject ?? "") });
                }
              }
              const emails = (v: any) => (Array.isArray(v) ? v.flatMap((x: any) => x.value ?? []) : (v?.value ?? [])).map((x: any) => String(x.address ?? "")).filter(emailOk);
              for (const c of cand) {
                const m: any = await cliente.fetchOne(String(c.uid), { source: true }, { uid: true });
                if (!m || !m.source) continue;
                const p: any = await simpleParser(m.source);
                const texto0 = String(p.text ?? "");
                const mm = rxRdo.exec(c.subject) || rxRdo.exec(c.subject + " " + texto0.slice(0, 600));
                if (!mm) continue;
                const raw = new Uint8Array(m.source);
                vistos.add(c.id); achados.push({ id: c.id, nro: Number(mm[1]), carregar: async () => ({
                  fromName: String(p.from?.value?.[0]?.name ?? ""), fromEmail: String(p.from?.value?.[0]?.address ?? ""),
                  to: emails(p.to), cc: emails(p.cc), subject: String(p.subject ?? c.subject), html: typeof p.html === "string" ? p.html : "", texto: texto0, date: p.date instanceof Date ? p.date : new Date(),
                  anexos: (p.attachments ?? []).filter((x: any) => x.content).map((x: any) =>
                    ({ name: String(x.filename ?? "anexo"), data: new Uint8Array(x.content), cid: (x.related || x.contentDisposition === "inline") && x.cid ? String(x.cid) : undefined, mime: String(x.contentType ?? "application/octet-stream") })),
                  raw: async () => raw }) });
              }
            } finally {
              trava.release();
              try { await cliente.logout(); } catch { /* ignora */ }
            }
            } catch (e) { errosCaixa.push(cx.user + ": " + String((e as any)?.message ?? e).slice(0, 120)); }
          }
          if (errosCaixa.length && errosCaixa.length === caixasImap.length) throw new Error("não consegui ler o e-mail (" + errosCaixa.join("; ") + ")");
          avisoCaixas = errosCaixa.length ? "Algumas caixas não abriram: " + errosCaixa.join("; ") : "";
        }
        const salvos: string[] = [];
        if (achados.length) {
          const tokD = await driveToken();
          const enviados = pastaEnv ? await driveLista(tokD, pastaEnv, " and mimeType!='application/vnd.google-apps.folder'") : [];
          for (const a of achados.reverse()) {
            // nome no padrão da pasta: "NN. RE Diário de obra DD-MM" (NN = próxima do sequencial da pasta; DD-MM = dia do RDO respondido)
            const full0 = await a.carregar();
            let dm = "";
            if (ini) { const dt = new Date(ini.getTime() + (a.nro - 1) * 864e5); dm = `${rd(dt.getUTCDate())}-${rd(dt.getUTCMonth() + 1)}`; }
            if (!dm) { const m2 = /(\d{1,2})[\/.-](\d{1,2})(?:[\/.-]\d{2,4})?/.exec(full0.subject); if (m2) dm = `${rd(Number(m2[1]))}-${rd(Number(m2[2]))}`; }
            const nome = `${String(await proximoNumero(tokD, pastaResp)).padStart(2, "0")}. RE Diário de obra ${dm || "RDO " + a.nro}`;
            // monta o .msg a partir da mensagem completa e dos anexos
            const full = full0;
            let ext = "msg", dados: Uint8Array;
            try {
              dados = montaMsg(CFB, { fromName: full.fromName, fromEmail: full.fromEmail, to: full.to, cc: full.cc, subject: full.subject, text: full.texto, html: full.html, date: full.date, anexos: full.anexos });
            } catch { ext = "eml"; dados = await full.raw(); }
            // se já existir um arquivo com esse nome (mais de uma resposta), acrescenta (2), (3)...
            const ja = await driveLista(tokD, pastaResp, " and mimeType!='application/vnd.google-apps.folder'");
            const usados = new Set(ja.map((f: any) => String(f.name)));
            let final = `${nome}.${ext}`, n = 2;
            while (usados.has(final)) final = `${nome} (${n++}).${ext}`;
            await driveGrava(tokD, pastaResp, final, ext === "msg" ? "application/vnd.ms-outlook" : "message/rfc822", dados, full.date);
            reg.ids[a.id] = { nro: a.nro, nome: final, t: Date.now() };
            salvos.push(final);
          }
          const ids = Object.entries(reg.ids).sort((x: any, y: any) => (y[1].t || 0) - (x[1].t || 0)).slice(0, 500);
          reg.ids = Object.fromEntries(ids);
          await admin.from("kv").upsert({ empresa_id: eu.empresa_id, key: "resp-reg", value: reg, deleted: false, autor: eu.id, updated_at: new Date().toISOString() }, { onConflict: "empresa_id,key" });
        }
        return json({ ok: true, config: true, via: usaMs ? "microsoft" : "imap", verificados, salvos, aviso: avisoCaixas });
      } catch (e) { return json({ ok: false, config: true, erro: String((e as any)?.message ?? e).slice(0, 300) }); }
    }

    // ---------- guarda no Drive as fotos do dia (as da aba Fotografias), em "07. Fotos da obra / MM-AAAA / DD-MM-AAAA" ----------
    // Uma pasta por dia, dentro da pasta do mês (a mesma sequência já usada). Pode ser chamada várias vezes: só sobem as fotos que ainda não estão lá.
    if (body.acao === "fotos-drive") {
      const dia = str(body.data, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dia)) return json({ error: "pedido inválido" }, 400);
      const { data: ac } = await admin.from("acessos").select("role").eq("user_id", eu.id).eq("empresa_id", eu.empresa_id).maybeSingle();
      if (!ac || ac.role === "vis") return json({ error: "Seu perfil não pode fazer isso." }, 403);
      const empDrive = Deno.env.get("DRIVE_EMPRESA_ID");
      if (!modoOAuth() || (empDrive && empDrive !== eu.empresa_id)) return json({ ok: true, config: false });
      const raiz = Deno.env.get("DRIVE_PASTA_FOTOS") || "14FDs4bnzXM0SHrt3O_qpyJi-d_lSIsRh";   // pasta "07. Fotos da obra"
      try {
        const { data: rows } = await admin.from("kv").select("key,value").eq("empresa_id", eu.empresa_id).like("key", `${dia}/ph/%`).eq("deleted", false).limit(1000);
        const fotos = (rows ?? []).map((r: any) => r.value).filter((v: any) => v && v.id && v.path && !v.del).sort((a: any, b: any) => (a.t || 0) - (b.t || 0));
        const [y, m, d] = dia.split("-");
        const pastaTxt = `${m}-${y}/${d}-${m}-${y}`;
        if (!fotos.length) return json({ ok: true, config: true, salvas: 0, jaExistiam: 0, restantes: 0, total: 0, pasta: pastaTxt });
        const tok = await driveToken();
        const mesId = await driveSubpasta(tok, raiz, `${m}-${y}`);
        const diaId = await driveSubpasta(tok, mesId, `${d}-${m}-${y}`);
        const q = encodeURIComponent(`'${diaId}' in parents and trashed=false and mimeType!='application/vnd.google-apps.folder'`);
        const ex: any = await gj(tok, `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,appProperties)&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true`);
        const arqs: any[] = ex.files ?? [];
        const ja = new Set<string>(arqs.map((f: any) => f.appProperties?.rdoFoto).filter(Boolean));
        const usados = new Set<string>(arqs.map((f: any) => String(f.name)));
        const falta = fotos.filter((f: any) => !ja.has(f.id));
        const lote = falta.slice(0, 15);
        let n = arqs.length, salvas = 0, falhas = 0;
        for (const f of lote) {
          try {
            const { data: blob, error: ed } = await admin.storage.from("fotos").download(f.path);
            if (ed || !blob) { falhas++; continue; }
            const dados = new Uint8Array(await blob.arrayBuffer());
            const atv = String(f.atv ?? "").replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim().slice(0, 60);
            let nome = "", k = n + 1;
            do { nome = `Foto ${String(k++).padStart(2, "0")}${atv ? " - " + atv : ""}.jpg`; } while (usados.has(nome));
            usados.add(nome); n = k - 1;
            const lim = "rdo" + crypto.randomUUID().replace(/-/g, "");
            const meta = { name: nome, parents: [diaId], appProperties: { rdoFoto: String(f.id) } };
            const cab = new TextEncoder().encode(`--${lim}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${lim}\r\nContent-Type: image/jpeg\r\n\r\n`);
            const fim = new TextEncoder().encode(`\r\n--${lim}--`);
            const corpo = new Uint8Array(cab.length + dados.length + fim.length);
            corpo.set(cab, 0); corpo.set(dados, cab.length); corpo.set(fim, cab.length + dados.length);
            await gj(tok, "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": `multipart/related; boundary=${lim}` }, body: corpo });
            salvas++;
          } catch { falhas++; }
        }
        return json({ ok: true, config: true, salvas, falhas, jaExistiam: ja.size, restantes: Math.max(0, falta.length - lote.length), total: fotos.length, pasta: pastaTxt });
      } catch (e) { return json({ ok: false, config: true, erro: String((e as any)?.message ?? e).slice(0, 300) }); }
    }

    // ---------- enviar o RDO por e-mail (PDF ao cliente, Excel ao coordenador) ----------
    // Os destinatários, o assunto e o texto vêm da configuração da obra (só o administrador altera), nunca do pedido.
    if (body.acao === "enviar-email") {
      // Envio por SMTP (Outlook/Microsoft 365, Gmail...) com SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS; ou pelo Resend com RESEND_API_KEY.
      // MAIL_FROM (opcional) é o remetente; no Outlook/Gmail precisa ser o mesmo endereço do SMTP_USER.
      const smtpHost = Deno.env.get("SMTP_HOST"), smtpUser = Deno.env.get("SMTP_USER"), smtpPass = Deno.env.get("SMTP_PASS");
      const chave = Deno.env.get("RESEND_API_KEY");
      // conta de e-mail de quem está enviando (CONTAS_EMAIL); sem ela, vale a conta geral (SMTP_USER)
      const loginEmail = String(data.user.email ?? "").toLowerCase();
      const conta = contasEmail().find((c) => c.login === loginEmail);
      const propria = !!(conta && smtpHost);
      const smtpU = propria ? conta!.user : smtpUser, smtpP = propria ? conta!.pass : smtpPass;
      const de = propria ? conta!.user : (Deno.env.get("MAIL_FROM") || (smtpUser || ""));
      const usaSmtp = !!(smtpHost && smtpU && smtpP);
      if (!de || (!usaSmtp && !chave)) return json({ error: "O envio por e-mail ainda não foi configurado no servidor (SMTP_HOST, SMTP_USER e SMTP_PASS)." }, 501);
      const tipo = body.tipo === "pdf" ? "pdf" : body.tipo === "xlsx" ? "xlsx" : "";
      const dia = str(body.data, 10), caminho = str(body.caminho, 300);
      if (!tipo || !/^\d{4}-\d{2}-\d{2}$/.test(dia) || caminho.includes("..") || !caminho.startsWith(`${eu.empresa_id}/${dia}/envio/`))
        return json({ error: "pedido inválido" }, 400);
      const { data: ac } = await admin.from("acessos").select("role, pode_gerar").eq("user_id", eu.id).eq("empresa_id", eu.empresa_id).maybeSingle();
      if (!ac || (ac.role !== "admin" && ac.role !== "apr")) return json({ error: "Só o administrador e o aprovador podem enviar e-mail." }, 403);
      const { data: kc } = await admin.from("kv").select("value").eq("empresa_id", eu.empresa_id).eq("key", "cfg").maybeSingle();
      const cfg: any = kc?.value ?? {};
      const cli = tipo === "pdf";
      const para = String((cli ? cfg.emailCli : cfg.emailCoord) ?? "").split(/[;,\s]+/).map((e) => e.trim().toLowerCase()).filter(emailOk).slice(0, 10);
      const cc = String((cli ? cfg.emailCcPdf : cfg.emailCcXlsx) ?? "").split(/[;,\s]+/).map((e) => e.trim().toLowerCase()).filter((e) => emailOk(e) && !para.includes(e)).slice(0, 10);   // com cópia
      const cco = String(cfg.emailCco ?? "").split(/[;,\s]+/).map((e) => e.trim().toLowerCase()).filter((e) => emailOk(e) && !para.includes(e) && !cc.includes(e)).slice(0, 10);   // cópia oculta
      if (!para.length) return json({ error: cli ? "Cadastre o e-mail do cliente em Configurações > Obra e logos." : "Cadastre o e-mail do coordenador em Configurações > Obra e logos." }, 400);
      const [y, m, d] = dia.split("-");
      const vars: Record<string, string> = {
        obra: String(cfg.obra ?? ""), obraCurta: String(cfg.obraCurta ?? ""), contratada: String(cfg.contratada ?? ""),
        contratante: String(cfg.contratante ?? ""), data: `${d}/${m}/${y}`, rdo: str(body.rdo, 10), usuario: str(data.user.user_metadata?.nome, 120),
      };
      // saudação pelo horário de quem está enviando (o app manda a hora local; sem ela, vale o horário de Cuiabá)
      const hora = Number.isInteger(body.hora) && body.hora >= 0 && body.hora < 24 ? body.hora
        : Number(new Intl.DateTimeFormat("en-GB", { hour: "2-digit", hourCycle: "h23", timeZone: "America/Cuiaba" }).format(new Date()));
      vars.saudacao = hora >= 5 && hora < 12 ? "Bom dia" : hora >= 12 && hora < 18 ? "Boa tarde" : "Boa noite";
      const preencher = (s: string) => s.replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? "");
      const assuntoCfg = String(cfg.emailAssunto || "").trim();
      const assunto = preencher(!assuntoCfg || assuntoCfg === "RDO {obraCurta} - {data}" ? "RES: Diário de obra" : assuntoCfg).replace(/[\r\n]+/g, " ").slice(0, 200);
      const texto0 = preencher(String((cli ? cfg.emailTxtCli : cfg.emailTxtCoord) ||
        "Prezados,\n\nSegue em anexo o diário de obra n° {rdo} da {contratada}, referente a obra {obraCurta}.\n\nAtenciosamente,\n{usuario}")).slice(0, 5000);
      // se o texto já começa com "Bom dia", "Boa tarde" ou "Boa noite", troca pela saudação certa do horário
      const texto = texto0.replace(/^(\s*)(bom dia|boa tarde|boa noite)\b/i, (_m, sp) => sp + vars.saudacao);
      const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      const { data: arq, error: ea } = await admin.storage.from("fotos").download(caminho);
      if (ea || !arq) return json({ error: "Não encontrei o arquivo enviado. Tente de novo." }, 400);
      const nome = str(body.nome, 120).replace(/[^\p{L}\p{N}._ ()-]+/gu, "_") || `RDO.${tipo}`;
      const respostas = Deno.env.get("MAIL_REPLY_TO") || (propria ? undefined : data.user.email) || undefined;   // para onde vão as respostas (na conta própria, as respostas voltam para o próprio remetente)
      const msgId = `<${crypto.randomUUID()}@rdo.potenciaeng>`, dataEnvio = new Date();
      const anexo = Buffer.from(await arq.arrayBuffer());
      const nomeArq = nome.endsWith("." + tipo) ? nome : `${nome}.${tipo}`;
      const ICONES: Record<string, string> = { fone: "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAADcElEQVR42n2UQWhcVRSGv3Pve/Pim5kkNSGJJTSaLCJpE1oSGqLtKAaqWAsipCCCKLgsFLoxVprJWJqdIKULQ1YigjG7SkFqKdYurNUWhSlIam1FMKHRGG1j82bevdfFe0MyafDAgQP38HPO+f/7ax4KJ7BTw5zjf8MJoOBSXZ9X31RUIBYwUOwEbxTcUyBPgBHgPnAR1NcgPwIOxjTMmRqC1IOVLIy3QHgK5DXwc+AAu6FdAVVAfwnRBJSuwGcaDhsAXQ92YgSy5yEzClEmeVMpUC0tUHXg9YD3OjxdhSOXE4xLTlIwB+/tAnUFCMHEkPEg/gHiz4Ff0hFzwMugRxNwU4W8D/eOweQH8Iz20hEMmF7wbTKBisC+CSc+Se5UF2eguBe8DyG7B9ZWwLYlfc/aDTc82QPSB3IMKtNQ+jQ5eJ88zHAphrd3QFMx2SC8BgvL8P5q2lycgkfegWgV1BK4F8HchJ2uduz1EGAi3czNQeYASANUvgUOqLRrXzKy86DhcXC7kilubDGdleSNHKjnwTYkJMkeoLmmQwWegAQJ8TYla1nD2KYbTgoULQQCFQ0+iUarPnguBZSPodIN8jfcXwAuaF2yQGQMaA0iEMegVJJxzJ8weQbMCyB5UF/B8h+bViq+BcGrUIngQUNCcA74F0i3pJJmzoAfQ/wrNB+Fo1H69cY09DlQB8GfUUphreLQob2Eoc/s7PcUCgP09rYwM/MN/f09FApdTE9/hzEO53LAXwbcETisPNimoFSFiX1KZXFuJWpry3lnz74CwPXrPzM7+xIdHXmuXr3J6dOjFApd3L27zNzcNeN527w45jkQB85uMAdZtdY6pXy9tPSPHh+/QGNjwK1bv3P8+EWGhzspl39jauoyd+6scO7cT4g0YIxVqWlsdhsnIuLAEIYB+/fvIJvN0NycZ3i4k6Gh7bS2NjE4uJ3duzvo7n6UcnkRpcQ5t6V9SQBaOeeiTMZjZKSTfD4gDDMMDXUwOPgY+XxAf38bAwPttLfnKJdNLCJBylZN9kUFkw5OPglyXiTb6dwDWlub8H3NwsISjY1NtLSE3L69SBBk6epqYn5+EQhS9tfegMmPYExLvX292wP5gxBZWBMApXJYm8gmqatAhFKhszZQcG8eTn2ROLi4TW69HjUBQyLqrer1GNO16j9rky13k2D54QAAAABJRU5ErkJggg==", mail: "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAB9klEQVR42s3UPWtUQRQG4Gfu3myMIRvWIJjGTgVRTCXYWFgrgiIoFsFCbEwhaKduVjQKiiBYpVKwERt/QEAEG60TGy1FCxM/SbLZvdlrMTdmExMhSMBp5hzOzHve88n/fsLyXSv9I1abenuzGNZ2kNwhz8kDoSd6zPO/fC2hSb7A1pS5J9Qn0mhcrNI7XMhYmEYvIWAt0BL5T5IBeuJTjUlMJNFezpjPmPtBG8kYU30MVhjsj/JUH3mFfBuj3WQnSD7TmONnhgYsMVykN2X2Fq0DbL/PvoNcONORmiWmGdfP0z/O7Aey85Sf006hYJjmUQyzjJ7lyzWqp7n1giu7lsFOlbhxl4Fx5ieYGyJ5TVdR5d+Aij95d5Sv3WRmhPQw1Rdc3UNtC/ufUr3Mt0e0jnFnBpXONCd/9pLApR7qD/l+iFClPEl4T+Uk01eonSsqHGIKrAtYUK20qKWMveHLEOElaZOvx6nfi7a3gfBHB6Sr9CQ+lkbPI93cfcfwUQZ7Y4i1coxkb4laoFRaAzDNCW38oJ51hFHcjxtLbUG9uZLD7enoILItAFtdhAQXqR0pUtExl6GYqM7JCSHqja30J+SlVX3Y+EjYid2xauEvu2SF3iT7GKNbsW0elJlq83WDG+MUPqW8yni2uGn70DoxbuTkm7KxfwFDJ6AluThKwAAAAABJRU5ErkJggg==", site: "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAACYklEQVR42rVUPWtUQRQ958772HWzQsCAiZqsEMRGFNJYCKawiFonNv4F7RQkwiroD9BGgr9AoqCg9sE6jdiocQ0IQQsx62bN7nvz7rXY7CabzUuw8DRTzJ0zZ849d4guqlWZej3mgGX8K5aHRg1L9z3+Bzg7O+sWFxeziavVaUhhWnTTOeccVA88rIRm6jJae2317cMFwBjUasMCIIPqlbBUurXxs66pzzKS+7OZgRQrD49EWSupAVgAgKC7H1BaG79+tG5en3lyqjJ2utFomsjerATMRREddO3O4+dTqhb1eHpFgiBNElcZO3Lm5PixS+vr63DO5YgzRHERmja/ClnvE9Z3M4l22yf1el0bGxsqpOQ9OGwnjAP+AazvFcHuShFQRERIiIjkKDQRIWAD+wOE1qk3bK95nFvlBxACiKO4wKDVdvt4yCguQNPNIgx+b0KDktRSHK2EgrHQ0ZxYXnYsoMqhcvETiFEQ8QChN5Xy8Eh8+9Gz845MzEAQlp9DkMTxTOWcEF96hENDo7YVrlX16bJ6X/B9ymkdq9ixjESnszQYAhdG7wH5sCOj+2P88t2X4sIJZKlCAhH4udqbB5/z6nc1ZevmfpylCytmBroQmZrbY3DQPbeDsCoA9cTM/AUJi5P0LTWFAIg1TRTm1UzFoNcqM/OrFoSiqW9+O/zxBRaZDSq8CMESlMCNICzOqSkYCNS31Mx3vh5TFRfdoyNEAljqf082jr5aAbKuyYM5JAlKx3wSEpUEwPZEmPaaQ8Lle7gE7aQhe5olrXfqE4Pl/WEZREIAaK6Uv/vtIfsP+AuDzw24EfQWZAAAAABJRU5ErkJggg==" };
      // assinatura de quem está enviando (Configurações > Minha assinatura de e-mail); se ele ainda não tiver, vale a antiga da obra
      const { data: prs } = await admin.from("usr_prefs").select("sig").eq("user_id", eu.id).maybeSingle();
      const sp: any = (prs as any)?.sig;
      const sig = sp && (sp.nome || sp.email)
        ? { nome: str(sp.nome, 80), cargo: str(sp.cargo, 80), fone: str(sp.fone, 40), mail: str(sp.email, 120), site: str(sp.site, 120) }
        : { nome: str(cfg.sigNome, 80), cargo: str(cfg.sigCargo, 80), fone: str(cfg.sigFone, 40), mail: str(cfg.sigEmail, 120), site: str(cfg.sigSite, 120) };
      const logoM = /^data:image\/(png|jpeg|jpg|gif);base64,([A-Za-z0-9+\/=]+)$/.exec(String(cfg.logoSig || cfg.logoPot || ""));
      const cid = (n: string) => `${n}@rdo`;
      const fonte = "font-family:Calibri,Arial,sans-serif;font-size:11pt";
      const serif = "font-family:'Times New Roman',Times,serif;font-weight:bold;color:#999999;font-size:11pt;line-height:1.15";
      const fonteL = "font-family:Arial,Helvetica,sans-serif;font-size:9pt";
      const linha = (ic: string, txt: string, cor: string, link = "", sub = false) => txt ? `<tr><td style="padding:1px 8px 1px 0;width:16px;vertical-align:middle"><img src="cid:${cid(ic)}" width="14" height="14" alt="" style="display:block;width:14px;height:14px"></td><td style="${fonteL};color:${cor};padding:1px 0;vertical-align:middle;line-height:1.15">${link ? `<a href="${link}" style="color:${cor};text-decoration:${sub ? "underline" : "none"}">${esc(txt)}</a>` : esc(txt)}</td></tr>` : "";
      const logoBytes = logoM ? Buffer.from(logoM[2], "base64") : null;
      const dim = logoBytes ? dimImg(new Uint8Array(logoBytes)) : null;
      const LW = 200, LH = dim && dim.w > 0 ? Math.max(1, Math.round(LW * dim.h / dim.w)) : 0;
      const assinatura = (sig.nome || sig.mail) ? `<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:0"><tr>
        ${logoM ? `<td style="padding:0 8px 0 0;vertical-align:middle"><img src="cid:${cid("logo")}" width="${LW}"${LH ? ` height="${LH}"` : ""} alt="" style="display:block;width:${LW}px${LH ? `;height:${LH}px` : ""}"></td>` : ""}
        <td style="padding:0 0 0 8px;${logoM ? "border-left:1px solid #808080;" : ""}vertical-align:middle">
        <div style="${serif};margin:0">${sig.nome ? esc(sig.nome) : ""}${sig.nome && sig.cargo ? "<br>" : ""}${sig.cargo ? esc(sig.cargo) : ""}</div>
        <table cellpadding="0" cellspacing="0" style="margin-top:2px;border-collapse:collapse">${linha("fone", sig.fone, "#1F3864")}${linha("mail", sig.mail, "#1F3864", sig.mail ? `mailto:${sig.mail}` : "")}${linha("site", sig.site, "#0563C1", sig.site ? (/^https?:/i.test(sig.site) ? sig.site : `http://${sig.site}`) : "", true)}</table>
        </td></tr></table><div style="border-top:1px solid #808080;width:440px;margin-top:2px;font-size:1px;line-height:1px">&nbsp;</div>` : "";
      const corpo = esc(texto).replace(/\r?\n/g, "<br>");
      const html = `<div style="${fonte}">${corpo}${assinatura}</div>`;
      const textoPuro = texto + (sig.nome ? `\n\n${sig.nome}${sig.cargo ? "\n" + sig.cargo : ""}${sig.fone ? "\n" + sig.fone : ""}${sig.mail ? "\n" + sig.mail : ""}${sig.site ? "\n" + sig.site : ""}` : "");
      const inl: any[] = (sig.nome || sig.mail) ? [
        ...(logoM ? [{ filename: "logo." + (logoM[1] === "jpeg" ? "jpg" : logoM[1]), content: Buffer.from(logoM[2], "base64"), cid: cid("logo"), contentType: "image/" + logoM[1] }] : []),
        ...["fone", "mail", "site"].map((n) => ({ filename: n + ".png", content: Buffer.from(ICONES[n], "base64"), cid: cid(n), contentType: "image/png" })),
      ] : [];
      let falha = "";
      try {
        if (usaSmtp) {
          const porta = Number(Deno.env.get("SMTP_PORT") || "587");
          const smtp = nodemailer.createTransport({
            host: smtpHost, port: porta, secure: porta === 465, requireTLS: porta !== 465,
            auth: { user: smtpU, pass: smtpP }, connectionTimeout: 20000, socketTimeout: 60000,
          });
          await smtp.sendMail({ from: de, to: para, cc: cc.length ? cc : undefined, bcc: cco.length ? cco : undefined, subject: assunto, text: textoPuro, html, replyTo: respostas, messageId: msgId, date: dataEnvio,
            attachments: [{ filename: nomeArq, content: anexo }, ...inl] });
        } else {
          const r = await fetch("https://api.resend.com/emails", {
            method: "POST",
            headers: { Authorization: `Bearer ${chave}`, "Content-Type": "application/json" },
            body: JSON.stringify({ from: de, to: para, cc: cc.length ? cc : undefined, bcc: cco.length ? cco : undefined, subject: assunto, text: textoPuro, html, reply_to: respostas,
              attachments: [{ filename: nomeArq, content: anexo.toString("base64") }, ...inl.map((a) => ({ filename: a.filename, content: a.content.toString("base64"), content_id: a.cid }))] }),
          });
          if (!r.ok) falha = (await r.text()).slice(0, 300);
        }
      } catch (e) { falha = String((e as any)?.message ?? e).slice(0, 300); }
      await admin.storage.from("fotos").remove([caminho]);   // o arquivo temporário não fica guardado
      if (falha) return json({ error: "O servidor de e-mail recusou o envio: " + falha }, 502);
      // cópia na pasta "Enviados" da caixa de quem enviou (o envio por SMTP não guarda cópia sozinho). Um problema aqui não desfaz o e-mail enviado.
      let copiaEnviados = "";
      if (usaSmtp) {
        try {
          const raw: any = await nodemailer.createTransport({ streamTransport: true, buffer: true, newline: "windows", keepBcc: true } as any).sendMail({
            from: de, to: para, cc: cc.length ? cc : undefined, bcc: cco.length ? cco : undefined, subject: assunto, text: textoPuro, html, replyTo: respostas, messageId: msgId, date: dataEnvio,
            attachments: [{ filename: nomeArq, content: anexo }, ...inl] });
          const { ImapFlow } = await import("npm:imapflow@1.0.164");
          const imapH = Deno.env.get("IMAP_HOST") || String(smtpHost).replace(/^smtp\./i, "imap.");
          const ip = Number(Deno.env.get("IMAP_PORT") || "993");
          const cl: any = new ImapFlow({ host: imapH, port: ip, secure: ip === 993, auth: { user: smtpU!, pass: smtpP! }, logger: false });
          cl.on("error", () => { /* aparece abaixo */ });
          await cl.connect();
          try {
            const lista: any[] = await cl.list();
            const env = lista.find((b: any) => b.specialUse === "\\Sent") || lista.find((b: any) => /^(inbox\.)?(sent|enviados|itens enviados|sent items|sent messages)$/i.test(String(b.name))) || lista.find((b: any) => /sent|enviad/i.test(String(b.path)));
            if (!env) copiaEnviados = "não achei a pasta de enviados (pastas da caixa: " + lista.map((b: any) => b.path).join(", ").slice(0, 200) + ")";
            else { await cl.append(env.path, raw.message, ["\\Seen"], dataEnvio); copiaEnviados = "ok"; console.log("cópia guardada em", env.path); }
          } finally { try { await cl.logout(); } catch { /* ignora */ } }
        } catch (e) { copiaEnviados = "não consegui guardar a cópia em Itens Enviados (" + String((e as any)?.message ?? e).slice(0, 160) + ")"; console.log(copiaEnviados); }
      }
      // Google Drive (opcional): um problema aqui não desfaz o e-mail que já foi enviado
      const drive: any[] = [];
      const empDrive = Deno.env.get("DRIVE_EMPRESA_ID");
      if (driveAtivo() && (!empDrive || empDrive === eu.empresa_id)) {
        try {
          const tipoMime = tipo === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
          drive.push({ ok: await driveSalvaArquivo(`${m}-${y}`, nomeArq, tipoMime, new Uint8Array(anexo)) });
          if (cli && usaSmtp) {
            const base = `${assunto.replace(/:\s*/g, "_").replace(/[\\/*?"<>|]+/g, "_")} ${d}-${m}-${y.slice(2)}`;
            let msg: Uint8Array | null = null;
            try {
              const mm = /^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/.exec(de);
              msg = montaMsg(CFB, { fromName: (mm ? mm[1] : "").trim(), fromEmail: (mm ? mm[2] : de).trim(), to: para, cc, subject: assunto, text: textoPuro, html, date: new Date(),
                anexos: [{ name: nomeArq, data: new Uint8Array(anexo), mime: "application/pdf" }, ...inl.map((a) => ({ name: a.filename, data: new Uint8Array(a.content), cid: a.cid, mime: a.contentType }))] });
            } catch (e) { drive.push({ erro: "não consegui montar o .msg (" + String((e as any)?.message ?? e).slice(0, 120) + "); salvei como .eml" }); }
            if (msg) drive.push({ ok: await driveSalvaEmail(base, msg, "msg") });
            else {
              const eml: any = await nodemailer.createTransport({ streamTransport: true, buffer: true, newline: "windows" }).sendMail({
                from: de, to: para, cc: cc.length ? cc : undefined, subject: assunto, text: textoPuro, html, replyTo: respostas,
                attachments: [{ filename: nomeArq, content: anexo }, ...inl] });
              drive.push({ ok: await driveSalvaEmail(base, new Uint8Array(eml.message), "eml") });
            }
          }
        } catch (e) { drive.push({ erro: String((e as any)?.message ?? e).slice(0, 200) }); }
      }
      return json({ ok: true, para, de, assinatura: !!(sig.nome || sig.mail), drive, copiaEnviados });
    }

    // ---------- descrição técnica pelo Claude ----------
    const { prompt } = body;
    if (typeof prompt !== "string" || prompt.length < 20 || prompt.length > 30000) {
      return json({ error: "pedido inválido" }, 400);
    }
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": Deno.env.get("ANTHROPIC_API_KEY")!,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: Deno.env.get("ANTHROPIC_MODEL") ?? "claude-sonnet-5-5",
        max_tokens: 3000,
        messages: [{ role: "user", content: prompt }],
      }),
    });
    if (r.status === 429) return json({ error: "limite" }, 429);
    if (!r.ok) return json({ error: "falha na API", detalhe: (await r.text()).slice(0, 300) }, 502);
    const out = await r.json();
    const text = (out.content ?? []).filter((b: any) => b.type === "text").map((b: any) => b.text).join("\n");
    return json({ text });
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
});
