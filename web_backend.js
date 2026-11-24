/* Versão web da Ferramenta PDPAS: as mesmas rotas /api/* do programa do computador, respondidas por Python (Pyodide) no próprio navegador.
   Nada vai a servidor: pedidos e rascunhos ficam no localStorage; o backup leva tudo para outro aparelho. */
(function () {
  window.PDPAS_WEB = true;   // as páginas mostram mensagens de erro próprias do navegador, não as do programa de mesa
  if (window.top !== window && /[?&]bpmn=1/.test(location.search)) return;   // fluxograma embutido: só desenha, não carrega o Python
  const PYO = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";
  const BASE = new URL("./", document.currentScript ? document.currentScript.src : location.href).href;
  const ARQS = ["motor.py", "pecas.py", "textos.py", "historico.py", "controle.py", "web_api.py", "regras.json", "dados/catalogo.json", "dados/textos_itens.json"];
  const realFetch = window.fetch.bind(window);
  let py = null, msg = null;

  function aviso(t, erro) {
    if (!msg) {
      msg = document.createElement("div");
      msg.setAttribute("role", "status");
      msg.style.cssText = "position:fixed;left:8px;right:8px;bottom:8px;z-index:99999;background:#1f5fa8;color:#fff;padding:10px 14px;border-radius:10px;font:14px system-ui,sans-serif;box-shadow:0 2px 10px #0004";
      document.addEventListener("DOMContentLoaded", () => document.body.appendChild(msg));
      if (document.body) document.body.appendChild(msg);
    }
    msg.style.background = erro ? "#b3261e" : "#1f5fa8";
    msg.textContent = t;
    msg.style.display = t ? "block" : "none";
  }

  const pronto = (async () => {
    aviso("Preparando a ferramenta (na primeira vez leva cerca de um minuto; depois fica guardada no aparelho)…");
    await new Promise((ok, no) => { const s = document.createElement("script"); s.src = PYO + "pyodide.js"; s.crossOrigin = "anonymous"; s.onload = ok; s.onerror = () => no(new Error("Não consegui baixar o Pyodide. Verifique a internet.")); document.head.appendChild(s); });
    py = await loadPyodide({ indexURL: PYO });
    await py.loadPackage(["lxml", "typing-extensions", "micropip"]);
    aviso("Carregando o catálogo e as regras…");
    const wheel = await (await realFetch(BASE + "python_docx-1.2.0-py3-none-any.whl")).arrayBuffer();
    py.FS.writeFile("/tmp/python_docx-1.2.0-py3-none-any.whl", new Uint8Array(wheel));
    await py.runPythonAsync("import micropip\nawait micropip.install('emfs:/tmp/python_docx-1.2.0-py3-none-any.whl', deps=False)");
    py.FS.mkdirTree("/app/dados");
    await Promise.all(ARQS.map(async a => {
      const r = await realFetch(BASE + "py/" + a);
      if (!r.ok) throw new Error("Faltou o arquivo " + a);
      py.FS.writeFile("/app/" + a, new Uint8Array(await r.arrayBuffer()));
    }));
    await py.runPythonAsync("import sys,os\nsys.path.insert(0,'/app')\nos.chdir('/app')\nimport web_api\nweb_api.carregar_historico_guardado()");
    aviso("");
    return py;
  })();
  pronto.catch(e => aviso(e.message, true));

  function baixar(d) {
    const bin = atob(d.b64), u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([u], { type: d.mime }));
    const a = document.createElement("a");
    a.href = url; a.download = d.nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  window.fetch = async function (url, opt) {
    const u = new URL(url instanceof Request ? url.url : String(url), location.href);
    if (!u.pathname.startsWith("/api/")) return realFetch(url, opt);
    const p = await pronto;
    const metodo = (opt && opt.method) || "GET";
    const corpo = (opt && opt.body) || "";
    const r = JSON.parse(p.globals.get("web_api").tratar(metodo, u.pathname, u.search.replace(/^\?/, ""), corpo));
    if (r.code === 200 && r.obj && r.obj.__download) { baixar(r.obj.__download); r.obj = { ok: true }; }
    if (r.code === 200 && metodo === "POST") nuvemApos(u.pathname, corpo, r.obj);
    return new Response(JSON.stringify(r.obj), { status: r.code, headers: { "Content-Type": "application/json" } });
  };


  /* Nuvem (Google Drive + Planilha mestra) por um Apps Script da própria pessoa. Endereço e código ficam só neste aparelho. */
  const CFG = "pdpas-nuvem-v1";
  function cfg() { try { return JSON.parse(localStorage.getItem(CFG) || "{}"); } catch (e) { return {}; } }
  function setCfg(c) { localStorage.setItem(CFG, JSON.stringify(c)); }
  let estadoNuvem = "";
  function mostrarNuvem(t) { estadoNuvem = t; const el = document.getElementById("wbNuvem"); if (el) el.textContent = t; }
  async function chamar(acao, dados) {
    const c = cfg();
    if (!c.url || !c.codigo) throw new Error("Nuvem não configurada.");
    const r = await realFetch(c.url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(Object.assign({ acao, codigo: c.codigo }, dados || {})) });
    const j = await r.json();
    if (!j.ok) throw new Error(j.erro || "falha");
    return j;
  }
  const api = () => py.globals.get("web_api");
  let tBackup = null; const fila = new Set(); let apagados = [];
  function nuvemApos(caminho, corpo, obj) {
    if (!cfg().url || !cfg().auto) return;
    let b = {}; try { b = corpo ? JSON.parse(corpo) : {}; } catch (e) { /* ignora */ }
    if (caminho === "/api/gerar" && obj.pasta) fila.add(obj.pasta.split("/").pop());
    else if (["/api/fluxo", "/api/gerar_peca"].includes(caminho) && b.id) fila.add(b.id);
    else if (caminho === "/api/apagar_pedido" && b.id) { fila.delete(b.id); apagados.push(b.id); }
    else if (caminho !== "/api/rascunho") return;
    clearTimeout(tBackup); tBackup = setTimeout(sincronizar, 3000);
  }
  async function sincronizar() {
    const ids = [...fila], ap = apagados; fila.clear(); apagados = [];
    try {
      mostrarNuvem("Salvando no Drive…");
      for (const id of ids) { try { await chamar("linha", { linha: JSON.parse(api().linha_mestra(id)) }); } catch (e) { if (!/Não encontrei/.test(e.message)) throw e; } }
      for (const id of ap) await chamar("apagado", { id });
      await chamar("salvar_backup", { conteudo: api().exportar() });
      mostrarNuvem("Salvo no Drive em " + new Date().toLocaleTimeString("pt-BR"));
    } catch (e) { ids.forEach(i => fila.add(i)); apagados = apagados.concat(ap); mostrarNuvem("Não consegui salvar no Drive: " + e.message); }
  }

  /* backup: levar os pedidos de um computador para outro */
  function barra() {
    const b = document.createElement("details");
    /* a caixa fica numa faixa própria no alto, para não cobrir os botões da página */
    const st = document.createElement("style");
    st.textContent = "body{box-sizing:border-box;padding-top:30px!important}.app{height:calc(100vh - 30px)!important}.app.wide{height:auto!important;min-height:calc(100vh - 30px)!important}";
    document.head.appendChild(st);
    b.style.cssText = "position:fixed;right:8px;top:3px;z-index:9999;font:13px system-ui,sans-serif;background:var(--cartao,#fff);color:var(--tx,#222);border:1px solid #9aa;border-radius:8px;padding:4px 8px;max-width:260px;max-height:calc(100vh - 12px);overflow:auto";
    b.innerHTML = '<summary style="cursor:pointer">Levar para outro computador</summary><p style="margin:6px 0">Os pedidos ficam só neste aparelho. Salve o backup (guarde no Drive ou OneDrive) e carregue no outro aparelho. As minutas são refeitas na hora.</p><button type="button" id="wbSalvar" style="margin:2px">Salvar backup</button><label style="display:inline-block;border:1px solid #889;border-radius:6px;padding:2px 8px;margin:2px;cursor:pointer">Carregar backup<input id="wbCarregar" type="file" accept=".json,application/json" hidden></label><div id="wbMsg" style="margin-top:4px"></div><hr style="margin:8px 0"><b>Drive e planilha mestra</b><p style="margin:4px 0">Cole o endereço do seu script do Google e o código que você definiu nele.</p><input id="wbUrl" placeholder="endereço do script (https://script.google.com/…/exec)" style="width:100%;box-sizing:border-box;margin:2px 0"><input id="wbCod" type="password" placeholder="código" autocomplete="off" style="width:100%;box-sizing:border-box;margin:2px 0"><label style="display:block;margin:2px 0"><input id="wbAuto" type="checkbox"> Salvar sozinho a cada alteração</label><button type="button" id="wbTestar" style="margin:2px">Conectar e testar</button><button type="button" id="wbEnviar" style="margin:2px">Enviar tudo agora</button><button type="button" id="wbTrazer" style="margin:2px">Trazer do Drive</button><button type="button" id="wbSair" style="margin:2px">Desconectar</button><div id="wbNuvem" style="margin-top:4px"></div><p style="margin:6px 0"><a href="https://github.com/gapapsnorte/pdpas-fluxo/blob/main/apps_script/TRANSPLANTAR.md" target="_blank" rel="noopener">Como trocar de conta</a></p>';
    document.body.appendChild(b);
    nuvemUI(b);
    const m = t => { b.querySelector("#wbMsg").textContent = t; };
    b.querySelector("#wbSalvar").onclick = async () => {
      try {
        const p = await pronto; const txt = p.globals.get("web_api").exportar();
        const bytes = new TextEncoder().encode(txt); let s = ""; bytes.forEach(c => s += String.fromCharCode(c));
        baixar({ nome: "backup_pdpas_" + new Date().toISOString().slice(0, 10) + ".json", b64: btoa(s), mime: "application/json" });
        m("Backup baixado.");
      } catch (e) { m(e.message); }
    };
    b.querySelector("#wbCarregar").onchange = async e => {
      const f = e.target.files[0]; if (!f) return;
      try {
        const p = await pronto; const r = JSON.parse(p.globals.get("web_api").importar_backup(await f.text()));
        m(`Carregado: ${r.pedidos} pedidos e ${r.rascunhos} rascunhos. Atualizando…`); setTimeout(() => location.reload(), 900);
      } catch (e2) { m(String(e2.message || e2).split("\n").pop()); }
      e.target.value = "";
    };
  }
  function nuvemUI(b) {
    const c = cfg(), q = id => b.querySelector(id);
    q("#wbUrl").value = c.url || ""; q("#wbCod").value = c.codigo || ""; q("#wbAuto").checked = !!c.auto;
    mostrarNuvem(c.url ? "Conectado a este script. Código guardado só neste aparelho." : "Não conectado.");
    const ler = () => ({ url: q("#wbUrl").value.trim(), codigo: q("#wbCod").value, auto: q("#wbAuto").checked });
    q("#wbAuto").onchange = () => setCfg(Object.assign(cfg(), { auto: q("#wbAuto").checked }));
    q("#wbTestar").onclick = async () => {
      const n = ler(); if (!/^https:\/\/script\.google\.com\//.test(n.url)) { mostrarNuvem("O endereço deve começar com https://script.google.com/"); return; }
      setCfg(n); mostrarNuvem("Testando…");
      try { await chamar("ping"); mostrarNuvem("Conexão certa. Use “Enviar tudo agora” para guardar o que já existe."); } catch (e) { mostrarNuvem("Não conectou: " + e.message); }
    };
    q("#wbEnviar").onclick = async () => {
      try {
        await pronto; mostrarNuvem("Enviando…");
        await chamar("linhas", { linhas: JSON.parse(api().todas_linhas()) });
        await chamar("salvar_backup", { conteudo: api().exportar() });
        mostrarNuvem("Tudo enviado em " + new Date().toLocaleTimeString("pt-BR"));
      } catch (e) { mostrarNuvem("Não enviou: " + e.message); }
    };
    q("#wbTrazer").onclick = async () => {
      try {
        await pronto; mostrarNuvem("Buscando…");
        const j = await chamar("ler_backup");
        if (!j.conteudo) { mostrarNuvem("Ainda não há backup no Drive."); return; }
        const r = JSON.parse(api().importar_backup(j.conteudo));
        mostrarNuvem(`Trazido: ${r.pedidos} pedidos e ${r.rascunhos} rascunhos. Atualizando…`); setTimeout(() => location.reload(), 900);
      } catch (e) { mostrarNuvem("Não trouxe: " + String(e.message || e).split("\n").pop()); }
    };
    q("#wbSair").onclick = () => { localStorage.removeItem(CFG); q("#wbUrl").value = ""; q("#wbCod").value = ""; q("#wbAuto").checked = false; mostrarNuvem("Desconectado. O Drive não foi alterado."); };
  }
  if (window.top === window) document.addEventListener("DOMContentLoaded", barra);   // no fluxograma embutido não há caixa de backup
})();
