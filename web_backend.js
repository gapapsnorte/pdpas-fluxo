/* Versão web da Ferramenta PDPAS: as mesmas rotas /api/* do programa do computador, respondidas por Python (Pyodide) no próprio navegador.
   Nada vai a servidor: pedidos e rascunhos ficam no localStorage; o backup leva tudo para outro aparelho. */
(function () {
  const PYO = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";
  const BASE = new URL("./", document.currentScript ? document.currentScript.src : location.href).href;
  const ARQS = ["motor.py", "pecas.py", "textos.py", "historico.py", "web_api.py", "regras.json", "dados/catalogo.json", "dados/textos_itens.json"];
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
    await new Promise((ok, no) => { const s = document.createElement("script"); s.src = PYO + "pyodide.js"; s.onload = ok; s.onerror = () => no(new Error("Não consegui baixar o Pyodide. Verifique a internet.")); document.head.appendChild(s); });
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
    const u = new URL(typeof url === "string" ? url : url.url, location.href);
    if (!u.pathname.startsWith("/api/")) return realFetch(url, opt);
    const p = await pronto.catch(e => { throw e; });
    const metodo = (opt && opt.method) || "GET";
    const corpo = (opt && opt.body) || "";
    const r = JSON.parse(p.globals.get("web_api").tratar(metodo, u.pathname, u.search.replace(/^\?/, ""), corpo));
    if (r.code === 200 && r.obj && r.obj.__download) { baixar(r.obj.__download); r.obj = { ok: true }; }
    return new Response(JSON.stringify(r.obj), { status: r.code, headers: { "Content-Type": "application/json" } });
  };

  /* backup: levar os pedidos de um computador para outro */
  function barra() {
    const b = document.createElement("details");
    b.style.cssText = "position:fixed;right:8px;top:8px;z-index:9999;font:13px system-ui,sans-serif;background:var(--cartao,#fff);color:var(--tx,#222);border:1px solid #9aa;border-radius:8px;padding:4px 8px;max-width:260px";
    b.innerHTML = '<summary style="cursor:pointer">Levar para outro computador</summary><p style="margin:6px 0">Os pedidos ficam só neste aparelho. Salve o backup (guarde no Drive ou OneDrive) e carregue no outro aparelho. As minutas são refeitas na hora.</p><button type="button" id="wbSalvar" style="margin:2px">Salvar backup</button><label style="display:inline-block;border:1px solid #889;border-radius:6px;padding:2px 8px;margin:2px;cursor:pointer">Carregar backup<input id="wbCarregar" type="file" accept=".json,application/json" hidden></label><div id="wbMsg" style="margin-top:4px"></div>';
    document.body.appendChild(b);
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
  document.addEventListener("DOMContentLoaded", barra);
})();
