// popup.js

// Performance pass by ChatGPT: keeps all HTML-facing handler IDs/names intact.
// Main changes: lighter API usage, bounded background loading, cache reuse, safer pagination, and one-query role privilege lookup.

const __D365_PERF_CACHE_TTL_MS = 10 * 60 * 1000;
const __d365PerfCache = new Map();

function __d365CacheGet(key) {
  const hit = __d365PerfCache.get(key);
  if (!hit || (Date.now() - hit.t) > __D365_PERF_CACHE_TTL_MS) {
    __d365PerfCache.delete(key);
    return null;
  }
  return hit.v;
}

function __d365CacheSet(key, value) {
  __d365PerfCache.set(key, { t: Date.now(), v: value });
  return value;
}

async function __d365GetActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs?.[0] || null;
}

document.getElementById("ribbondebug").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id || !tab?.url) return;

  if (tab.url.includes("ribbondebug=true")) return;

  const joiner = tab.url.includes("?") ? "&" : "?";
  chrome.tabs.update(tab.id, { url: tab.url + joiner + "ribbondebug=true" });
});

document.getElementById("tabsname").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  // 1) Collect tabs via Xrm from ALL frames
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: true },
    world: "MAIN",
    func: () => {
      try {
        const Xrm = window.Xrm;
        const page = Xrm?.Page;

        if (!page?.ui?.tabs?.forEach) return null;

        const tabs = [];
        page.ui.tabs.forEach((t) => {
          tabs.push({ name: t.getName(), label: t.getLabel() });
        });

        if (!tabs.length) return null;
        return {
          href: location.href,
          count: tabs.length,
          tabs
        };
      } catch (e) {
        return { error: String(e) };
      }
    }
  });

  // 2) Pick the best frame (most tabs)
  const candidates = results
    .map(r => r.result)
    .filter(r => r && !r.error && Array.isArray(r.tabs));

  const best = candidates.reduce((a, b) => (b.count > a.count ? b : a), { count: 0, tabs: [] });

  if (!best.count) {
    alert("Couldn't read tabs via Xrm.Page.ui.tabs.\nMake sure you are on a record form.");
    return;
  }

  // 3) Build text output
 const rows = best.tabs.map(t => ({
  label: t.label || "",
  name: t.name || ""
}));

const col1Width = Math.max(
  "name".length,
  ...rows.map(r => r.label.length)
);

const col2Width = Math.max(
  "logical name".length,
  ...rows.map(r => r.name.length)
);

const padRight = (s, len) => (s + " ".repeat(len)).slice(0, len);

const text =
  `Found ${best.count} tabs\n` +
  `${padRight("name", col1Width)} | ${padRight("logical name", col2Width)}\n` +
  `${"-".repeat(col1Width)}-+-${"-".repeat(col2Width)}\n` +
  rows.map(r => `${padRight(r.label, col1Width)} | ${padRight(r.name, col2Width)}`).join("\n");

  // 4) Show copyable multiline modal in the PAGE
  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    args: [text],
    func: (text) => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,.35);
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(820px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `
        padding: 12px 14px;
        font-weight: 700;
        border-bottom: 1px solid #e5e7eb;
      `;
      header.textContent = "D365 Tabs (copyable)";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px;`;

      const ta = document.createElement("textarea");
      ta.value = text;
      ta.readOnly = true;
      ta.style.cssText = `
        width: 100%;
        height: 320px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
      `;
        ta.style.fontFamily = "Consolas, Monaco, 'Courier New', monospace";
        ta.style.whiteSpace = "pre";
        ta.style.direction = "ltr";       // ✅ הכי חשוב
        ta.style.textAlign = "left";    
      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 14px;
        border-top: 1px solid #e5e7eb;
      `;

      const btnClose = document.createElement("button");
      btnClose.textContent = "Close";
      btnClose.style.cssText = `
        border: 1px solid #cbd5e1;
        padding: 10px 14px;
        border-radius: 10px;
        cursor: pointer;
        background: #fff;
        font-weight: 700;
      `;

      const btnCopy = document.createElement("button");
      btnCopy.textContent = "Copy";
      btnCopy.style.cssText = `
        border: none;
        padding: 10px 14px;
        border-radius: 10px;
        cursor: pointer;
        background: #2563eb;
        color: #fff;
        font-weight: 700;
      `;

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(ta.value);
        } catch (e) {
          ta.focus();
          ta.select();
          document.execCommand("copy");
        }
        btnCopy.textContent = "Copied ✅";
        setTimeout(() => (btnCopy.textContent = "Copy"), 900);
      };

      const close = () => overlay.remove();
      btnClose.onclick = close;
    
      body.appendChild(ta);
      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      ta.focus();
      ta.select();
    }
  });
});

document.getElementById("getFieldValue").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  const input = prompt("Enter field logical names separated by comma:\nexample: firstname,lastname,emailaddress1");
  if (!input) return;

  const fields = input
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);

  if (!fields.length) {
    alert("No fields entered.");
    return;
  }

  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: true },
    world: "MAIN",
    args: [fields],
    func: async (fields) => {
      try {
        const Xrm = window.Xrm;
        const page = Xrm?.Page;

        if (!Xrm || !page?.data?.entity?.getId || !page?.data?.entity?.getEntityName) {
          return { ok: false, error: "Open a record form first (Xrm record context not found)." };
        }

        const entityName = page.data.entity.getEntityName();
        const id = (page.data.entity.getId() || "").replace(/[{}]/g, "");
        if (!entityName || !id) return { ok: false, error: "Record id/entity not found." };

        const webApi = Xrm.WebApi || Xrm?.WebApi?.online;
        if (!webApi?.retrieveRecord) {
          return { ok: false, error: "Xrm.WebApi.retrieveRecord not available." };
        }

        // ✅ build $select
        const select = fields.join(",");

        try {
          const rec = await webApi.retrieveRecord(entityName, id, `?$select=${select}`);

          // ✅ build response values
          const out = fields.map((f) => {
            return {
              field: f,
              value: rec[f] ?? null,
              formatted: rec[`${f}@OData.Community.Display.V1.FormattedValue`] ?? null
            };
          });

          return { ok: true, entityName, id, fields, values: out };
        } catch (err) {
          return {
            ok: false,
            error:
              `Retrieve failed.\nCheck your field names:\n${fields.join(", ")}\n\n` +
              (err?.message || err?.toString?.() || "Unknown error")
          };
        }
      } catch (e) {
        return { ok: false, error: String(e) };
      }
    }
  });

  const best =
    results.map(r => r.result).find(r => r?.ok === true) ||
    results.map(r => r.result).find(r => r?.ok === false);

  if (!best) {
    alert("No result returned. Try again.");
    return;
  }

  if (!best.ok) {
    alert(best.error);
    return;
  }

  // ✅ Pretty output
  const text =
    `Retrieved Fields\n\n` +
    `Entity: ${best.entityName}\n` +
    `Id: ${best.id}\n\n` +
    best.values
      .map(v => `${v.field} => ${v.formatted ?? JSON.stringify(v.value)}`)
      .join("\n");

  alert(text);
});
// If the user leaves "Fields" empty -> retrieveRecord WITHOUT $select (returns the full object)

document.getElementById("retrieveByIdUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(860px, 96vw); background: #fff; border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35); overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 800; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "D365 Retrieve By Entity + Id";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const row = (label, inputEl) => {
        const wrap = document.createElement("div");
        wrap.style.cssText = `display:grid; gap:6px;`;
        const l = document.createElement("div");
        l.textContent = label;
        l.style.cssText = `font-size: 12px; font-weight: 700; color: #111827;`;
        wrap.appendChild(l);
        wrap.appendChild(inputEl);
        return wrap;
      };

      const inputStyle = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const entityInput = document.createElement("input");
      entityInput.placeholder = "Entity logical name (e.g. contact)";
      entityInput.style.cssText = inputStyle;

      const idInput = document.createElement("input");
      idInput.placeholder = "GUID (with or without {})";
      idInput.style.cssText = inputStyle;

      const fieldsInput = document.createElement("input");
      fieldsInput.placeholder = "Fields (comma separated) e.g. firstname,lastname,emailaddress1  | leave empty = ALL fields";
      fieldsInput.style.cssText = inputStyle;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Result will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 260px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: ltr;
        text-align: left;
      `;

      body.appendChild(row("Entity", entityInput));
      body.appendChild(row("GUID", idInput));
      body.appendChild(row("Fields (optional)", fieldsInput));
      body.appendChild(status);
      body.appendChild(resultTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex; gap: 10px; justify-content: flex-end;
        padding: 12px 14px; border-top: 1px solid #e5e7eb;
      `;

      const btn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 800;
        `;
        return b;
      };

      const btnClose = btn("Close");

      const btnCopy = btn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnRetrieve = btn("Retrieve");
      btnRetrieve.style.border = "none";
      btnRetrieve.style.background = "#111827";
      btnRetrieve.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch (e) {
          resultTa.focus();
          resultTa.select();
          document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      btnRetrieve.onclick = async () => {
        const entityName = (entityInput.value || "").trim();
        const id = (idInput.value || "").trim().replace(/[{}]/g, "");
        const rawFields = (fieldsInput.value || "").trim();

        status.textContent = "";
        resultTa.value = "";

        if (!entityName || !id) {
          status.textContent = "❌ Please fill Entity and GUID.";
          return;
        }

        if (!/^[0-9a-fA-F-]{36}$/.test(id)) {
          status.textContent = "❌ GUID looks invalid. Example: f557616e-26ec-e611-a8a7-0050568c00dc";
          return;
        }

        const Xrm = window.Xrm;
        const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;

        if (!webApi?.retrieveRecord) {
          status.textContent = "❌ Xrm.WebApi.retrieveRecord not available on this page.";
          return;
        }

        status.textContent = "⏳ Retrieving…";

        // If empty => no $select => ALL fields (full object)
        const fields = rawFields
          ? rawFields.split(",").map(s => s.trim()).filter(Boolean)
          : [];

        const formatOne = (f, rec) => {
          const raw = rec[f];
          const formatted = rec[`${f}@OData.Community.Display.V1.FormattedValue`];
          const lookupLn = rec[`${f}@Microsoft.Dynamics.CRM.lookuplogicalname`];
          const shown = (formatted != null)
            ? `${formatted} (raw: ${JSON.stringify(raw)})`
            : JSON.stringify(raw);
          const extra = lookupLn ? ` (lookup: ${lookupLn})` : "";
          return `${f}${extra} => ${shown}`;
        };

        try {
          // ✅ ALL fields
          if (fields.length === 0) {
            const rec = await webApi.retrieveRecord(entityName, id);
            resultTa.value =
              `Entity: ${entityName}\nId: ${id}\n\n` +
              JSON.stringify(rec, null, 2);
            status.textContent = "✅ Done (ALL fields).";
            resultTa.focus();
            resultTa.select();
            return;
          }

          // ✅ Selected fields (same behavior as before)
          const select = fields.join(",");
          const rec = await webApi.retrieveRecord(entityName, id, `?$select=${select}`);
          const lines = fields.map(f => formatOne(f, rec));

          resultTa.value =
            `Entity: ${entityName}\nId: ${id}\n\n` +
            lines.join("\n");

          status.textContent = `✅ Done (${fields.length} fields).`;
        } catch (err1) {
          // fallback per-field if user provided fields (keep your logic)
          if (fields.length === 0) {
            status.textContent = "❌ Failed.";
            resultTa.value =
              "ERROR:\n" +
              (err1?.message || err1?.toString?.() || "Unknown error");
            return;
          }

          const lines = [];
          for (const f of fields) {
            try {
              const rec = await webApi.retrieveRecord(entityName, id, `?$select=${f}`);
              lines.push(formatOne(f, rec));
            } catch (err2) {
              const msg = err2?.message || err2?.toString?.() || "Unknown error";
              lines.push(`${f} => ❌ Failed (${msg})`);
            }
          }

          resultTa.value =
            `Entity: ${entityName}\nId: ${id}\n\n` +
            lines.join("\n");

          status.textContent = "✅ Done (some fields may have failed).";
        }
      };

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnRetrieve);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      entityInput.focus();
    }
  });
});

// popup.js  (RetrieveMultiple UI button - FULL CODE)
// Requires a button in popup.html: <button id="retrieveMultipleUi">RetrieveMultiple</button>

document
  .getElementById("retrieveMultiple")
  .addEventListener("click", async () => {
    const tab = await __d365GetActiveTab();
    if (!tab?.id) return;

    await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: false },
      world: "MAIN",
      func: async () => {
        const MODAL_ID = "__rv_advanced_retrieve_multiple";
        document.getElementById(MODAL_ID)?.remove();

        const clientUrl =
          window.Xrm?.Utility?.getGlobalContext?.()?.getClientUrl?.();

        if (!clientUrl) {
          alert("D365 context not found.");
          return;
        }

        const API = `${clientUrl}/api/data/v9.2`;
        const PAGE_SIZE = 100;

        const html = (value) =>
          String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");

        const labelOf = (label) =>
          label?.UserLocalizedLabel?.Label ||
          label?.LocalizedLabels?.[0]?.Label ||
          "";

        const requestJson = async (url) => {
          const response = await fetch(url, {
            method: "GET",
            credentials: "same-origin",
            headers: {
              Accept: "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0",
              Prefer:
                `odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname",odata.maxpagesize=${PAGE_SIZE}`
            }
          });

          const text = await response.text();
          let body = null;
          try {
            body = text ? JSON.parse(text) : null;
          } catch {
            body = text;
          }

          if (!response.ok) {
            throw new Error(
              body?.error?.message ||
                `HTTP ${response.status} ${response.statusText}`
            );
          }

          return body;
        };

        const apiGet = (relativeUrl) =>
          requestJson(relativeUrl.startsWith("http") ? relativeUrl : `${API}/${relativeUrl}`);

        const style = document.createElement("style");
        style.textContent = `
          #${MODAL_ID} * { box-sizing:border-box; }
          #${MODAL_ID} {
            position:fixed; inset:0; z-index:2147483647;
            padding:12px; direction:rtl;
            background:rgba(241,245,249,.96);
            backdrop-filter:blur(3px);
            font-family:system-ui,-apple-system,"Segoe UI",Arial,sans-serif;
          }
          #${MODAL_ID} .arm-dialog {
            width:calc(100vw - 24px); height:calc(100vh - 24px);
            display:flex; flex-direction:column; overflow:hidden;
            color:#111827; background:#f8fafc;
            border:1px solid #cbd5e1; border-radius:18px;
            box-shadow:0 24px 70px rgba(15,23,42,.16);
          }
          #${MODAL_ID} .arm-header {
            display:flex; align-items:center; justify-content:space-between;
            gap:16px; padding:18px 22px;
            background:#ffffff;
            border-bottom:1px solid #e2e8f0;
          }
          #${MODAL_ID} .arm-title { color:#0f172a; font-size:30px; font-weight:950; line-height:1.1; }
          #${MODAL_ID} .arm-subtitle { margin-top:5px; color:#64748b; font-size:13px; }
          #${MODAL_ID} .arm-close {
            width:42px; height:42px; border-radius:12px; cursor:pointer;
            color:#0f172a; background:#ffffff; border:1px solid #cbd5e1;
            font-size:18px; font-weight:900;
          }
          #${MODAL_ID} .arm-close:hover { background:#f8fafc; }
          #${MODAL_ID} .arm-body {
            flex:1; min-height:0; overflow:auto;
            display:flex; flex-direction:column;
            gap:14px; padding:14px;
            background:#f8fafc;
          }
          #${MODAL_ID} .arm-search-panel,
          #${MODAL_ID} .arm-results-panel {
            width:100%; background:#ffffff;
            border:1px solid #d7dee8; border-radius:16px;
            box-shadow:0 8px 24px rgba(15,23,42,.04);
          }
          #${MODAL_ID} .arm-search-panel {
            padding:16px;
            display:flex; flex-direction:column; gap:14px;
          }
          #${MODAL_ID} .arm-results-panel {
            flex:1; min-height:360px;
            padding:14px;
            display:flex; flex-direction:column; gap:10px;
          }
          #${MODAL_ID} .arm-builder-head {
            display:flex; align-items:flex-start; justify-content:space-between;
            gap:12px; flex-wrap:wrap;
            padding-bottom:2px;
          }
          #${MODAL_ID} .arm-builder-title { color:#111827; font-size:18px; font-weight:950; }
          #${MODAL_ID} .arm-builder-desc { margin-top:3px; color:#64748b; font-size:12px; }
          #${MODAL_ID} .arm-builder-meta {
            display:flex; gap:8px; flex-wrap:wrap;
          }
          #${MODAL_ID} .arm-pill {
            display:inline-flex; align-items:center; justify-content:center;
            min-height:30px; padding:6px 10px;
            color:#334155; background:#f8fafc;
            border:1px solid #d7dee8; border-radius:999px;
            font-size:11px; font-weight:800;
          }
          #${MODAL_ID} .arm-top-grid {
            display:grid; grid-template-columns:minmax(320px,1.1fr) minmax(340px,1.25fr) minmax(280px,.9fr);
            gap:14px; align-items:stretch;
          }
          #${MODAL_ID} .arm-block {
            min-width:0; padding:14px;
            background:#f8fafc; border:1px solid #e2e8f0; border-radius:14px;
          }
          #${MODAL_ID} .arm-block-title {
            display:flex; align-items:center; justify-content:space-between;
            gap:10px; margin-bottom:10px;
            color:#0f172a; font-size:14px; font-weight:950;
          }
          #${MODAL_ID} .arm-block-subtitle {
            margin-top:4px; color:#64748b; font-size:11px; font-weight:600;
          }
          #${MODAL_ID} .arm-fields-help {
            margin-top:6px; color:#64748b; font-size:11px; line-height:1.5;
          }
          #${MODAL_ID} .arm-compact-grid {
            display:grid; grid-template-columns:1fr 1fr; gap:10px;
          }
          #${MODAL_ID} .arm-condition-section {
            padding:14px; background:#f8fafc;
            border:1px solid #e2e8f0; border-radius:14px;
          }
          #${MODAL_ID} .arm-actions-bar {
            display:flex; align-items:center; justify-content:space-between;
            gap:12px; flex-wrap:wrap;
            padding:12px 14px;
            background:#0f172a; border-radius:14px;
          }
          #${MODAL_ID} .arm-status-wrap { display:flex; flex-direction:column; gap:4px; }
          #${MODAL_ID} .arm-status-label { color:#cbd5e1; font-size:11px; font-weight:800; }
          #${MODAL_ID} .arm-status { color:#ffffff; font-size:13px; font-weight:700; }
          #${MODAL_ID} #armCount { color:#0f172a; font-size:13px; font-weight:800; }
          #${MODAL_ID} .arm-query-wrap {
            display:grid; gap:8px;
            padding:14px; background:#f8fafc;
            border:1px solid #e2e8f0; border-radius:14px;
          }
          #${MODAL_ID} .arm-query-title {
            color:#0f172a; font-size:13px; font-weight:900;
          }
          #${MODAL_ID} label {
            display:grid; gap:6px; color:#334155; font-size:11px; font-weight:850;
          }
          #${MODAL_ID} input, #${MODAL_ID} select, #${MODAL_ID} textarea {
            width:100%; min-width:0; padding:10px 11px;
            color:#111827; background:#ffffff; border:1px solid #b8c3d1;
            border-radius:10px; outline:none; font-size:12px;
          }
          #${MODAL_ID} input::placeholder, #${MODAL_ID} textarea::placeholder { color:#94a3b8; }
          #${MODAL_ID} input:focus, #${MODAL_ID} select:focus, #${MODAL_ID} textarea:focus {
            border-color:#0f172a; box-shadow:0 0 0 2px rgba(15,23,42,.08);
          }
          #${MODAL_ID} select[multiple] {
            min-height:220px; max-height:320px; padding:5px;
          }
          #${MODAL_ID} select[multiple] option { padding:7px 9px; border-radius:6px; }
          #${MODAL_ID} select[multiple] option:checked {
            color:#ffffff; background:#111827 linear-gradient(0deg,#111827,#111827);
          }
          #${MODAL_ID} .arm-actions { display:flex; flex-wrap:wrap; gap:8px; }
          #${MODAL_ID} button.arm-btn {
            padding:9px 13px; border-radius:10px; cursor:pointer;
            color:#111827; background:#ffffff; border:1px solid #b8c3d1;
            font-size:12px; font-weight:900;
          }
          #${MODAL_ID} button.arm-btn:hover { background:#f8fafc; border-color:#64748b; }
          #${MODAL_ID} button.arm-primary { color:#ffffff; background:#111827; border-color:#111827; }
          #${MODAL_ID} button.arm-primary:hover { color:#ffffff; background:#000000; }
          #${MODAL_ID} button.arm-primary-soft { color:#111827; background:#ffffff; border-color:#ffffff; }
          #${MODAL_ID} button.arm-danger { color:#b91c1c; background:#ffffff; border-color:#fecaca; }
          #${MODAL_ID} button:disabled { opacity:.45; cursor:not-allowed; }
          #${MODAL_ID} .arm-condition {
            display:grid;
            grid-template-columns:100px minmax(320px,2.2fr) 170px minmax(220px,1.4fr) 42px;
            gap:10px; align-items:end; margin-top:10px; padding:12px;
            background:#ffffff; border:1px solid #d7dee8; border-radius:12px;
          }
          #${MODAL_ID} .arm-condition:first-child { margin-top:0; }
          #${MODAL_ID} .arm-condition-value,
          #${MODAL_ID} .arm-condition-value input,
          #${MODAL_ID} .arm-condition-value select {
            width:100%; max-width:none;
          }
          #${MODAL_ID} .arm-remove {
            width:42px; height:40px; padding:0; border-radius:10px; cursor:pointer;
            color:#b91c1c; background:#ffffff; border:1px solid #fecaca;
            font-weight:950;
          }
          #${MODAL_ID} .arm-query {
            direction:ltr; text-align:left; min-height:78px; max-height:160px; resize:vertical;
            font-family:Consolas,Monaco,monospace; color:#111827; background:#ffffff;
          }
          #${MODAL_ID} .arm-results-head {
            display:flex; align-items:center; justify-content:space-between;
            gap:10px; flex-wrap:wrap;
          }
          #${MODAL_ID} .arm-results-title { color:#0f172a; font-size:16px; font-weight:950; }
          #${MODAL_ID} .arm-results-subtitle { margin-top:3px; color:#64748b; font-size:12px; }
          #${MODAL_ID} .arm-results-toolbar {
            display:flex; align-items:center; justify-content:space-between;
            gap:10px; flex-wrap:wrap;
            padding:12px 14px; background:#f8fafc;
            border:1px solid #e2e8f0; border-radius:12px;
          }
          #${MODAL_ID} .arm-results {
            flex:1; min-height:420px; overflow:auto; background:#ffffff;
            border:1px solid #d7dee8; border-radius:12px;
          }
          #${MODAL_ID} table {
            width:max-content; min-width:100%; border-collapse:collapse;
            direction:ltr; text-align:left; font-size:12px;
          }
          #${MODAL_ID} th {
            position:sticky; top:0; z-index:2; padding:10px 9px;
            color:#ffffff; background:#111827; border-bottom:1px solid #111827;
            white-space:nowrap;
          }
          #${MODAL_ID} td {
            max-width:420px; padding:9px 10px; color:#111827;
            background:#ffffff; border-bottom:1px solid #e2e8f0;
            white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
          }
          #${MODAL_ID} tr:nth-child(even) td { background:#f8fafc; }
          #${MODAL_ID} tr:hover td { background:#eef2f7; }
          #${MODAL_ID} .arm-link { color:#0f172a; text-decoration:underline; font-weight:850; }
          #${MODAL_ID} .arm-empty { padding:42px; color:#64748b; text-align:center; }
          #${MODAL_ID} .arm-badge {
            display:inline-flex; padding:4px 8px; border-radius:999px;
            color:#ffffff; background:#111827; border:1px solid #111827;
            font-size:10px; font-weight:900;
          }
          #${MODAL_ID} * { scrollbar-color:#94a3b8 #f1f5f9; }
          #${MODAL_ID} ::-webkit-scrollbar { width:10px; height:10px; }
          #${MODAL_ID} ::-webkit-scrollbar-track { background:#f1f5f9; }
          #${MODAL_ID} ::-webkit-scrollbar-thumb { background:#94a3b8; border:2px solid #f1f5f9; border-radius:10px; }
          #${MODAL_ID} ::-webkit-scrollbar-thumb:hover { background:#64748b; }
          @media (max-width:1400px) {
            #${MODAL_ID} .arm-top-grid { grid-template-columns:1fr 1fr; }
            #${MODAL_ID} .arm-top-grid > :last-child { grid-column:1 / -1; }
          }
          @media (max-width:1100px) {
            #${MODAL_ID} .arm-top-grid,
            #${MODAL_ID} .arm-compact-grid { grid-template-columns:1fr 1fr; }
            #${MODAL_ID} .arm-condition {
              grid-template-columns:90px minmax(220px,1.6fr) 150px minmax(180px,1.1fr) 42px;
            }
          }
          @media (max-width:760px) {
            #${MODAL_ID} { padding:0; }
            #${MODAL_ID} .arm-dialog { width:100vw; height:100vh; border-radius:0; }
            #${MODAL_ID} .arm-top-grid,
            #${MODAL_ID} .arm-compact-grid,
            #${MODAL_ID} .arm-condition { grid-template-columns:1fr; }
            #${MODAL_ID} .arm-remove { width:100%; }
          }
        `;
        document.head.appendChild(style);

        const overlay = document.createElement("div");
        overlay.id = MODAL_ID;
        overlay.innerHTML = `
          <div class="arm-dialog">
            <div class="arm-header">
              <div>
                <div class="arm-title">Advanced Retrieve Multiple</div>
                <div class="arm-subtitle">Metadata-driven Dataverse search builder</div>
              </div>
              <button class="arm-close" id="armCloseTop">✕</button>
            </div>

            <div class="arm-body">
              <section class="arm-search-panel">
                <div class="arm-builder-head">
                  <div>
                    <div class="arm-builder-title">Search builder</div>
                    <div class="arm-builder-desc">Everything related to the search is concentrated here: select a table, choose columns, add filters, then run the query.</div>
                  </div>
                  <div class="arm-builder-meta">
                    <span class="arm-pill">Step 1: Table</span>
                    <span class="arm-pill">Step 2: Columns</span>
                    <span class="arm-pill">Step 3: Conditions</span>
                    <span class="arm-pill">Step 4: Run</span>
                  </div>
                </div>

                <div class="arm-top-grid">
                  <div class="arm-block">
                    <div class="arm-block-title">1. Table selection</div>
                    <div class="arm-block-subtitle">First choose the Dataverse table you want to search.</div>
                    <label style="margin-top:10px">Search table
                      <input id="armEntitySearch" placeholder="Display name / logical name" />
                    </label>
                    <label style="margin-top:10px">Select table
                      <select id="armEntitySelect" size="8"></select>
                    </label>
                  </div>

                  <div class="arm-block">
                    <div class="arm-block-title">2. Columns to return</div>
                    <div class="arm-block-subtitle">You can load and return many columns at once.</div>
                    <label style="margin-top:10px">Search column
                      <input id="armFieldSearch" placeholder="Display name / logical name" disabled />
                    </label>
                    <label style="margin-top:10px">Select columns to display
                      <select id="armColumns" multiple size="12" disabled></select>
                    </label>
                    <div class="arm-fields-help">
                      Search and select columns one by one. Selected columns stay loaded even when the search text changes.
                    </div>
                    <div id="armSelectedColumnsSummary" style="margin-top:8px;padding:9px 10px;background:#ffffff;border:1px solid #d7dee8;border-radius:10px;color:#334155;font-size:11px;font-weight:800">
                      0 columns selected
                    </div>
                    <div class="arm-actions" style="margin-top:10px">
                      <button class="arm-btn" id="armSelectRecommended">Recommended</button>
                      <button class="arm-btn" id="armClearColumns">Clear</button>
                    </div>
                  </div>

                  <div class="arm-block">
                    <div class="arm-block-title">3. Sort, status and page size</div>
                    <div class="arm-block-subtitle">Optional settings that affect the query results.</div>
                    <div class="arm-compact-grid" style="margin-top:10px">
                      <label>Order by
                        <select id="armOrderField" disabled></select>
                      </label>
                      <label>Direction
                        <select id="armOrderDirection">
                          <option value="asc">Ascending</option>
                          <option value="desc">Descending</option>
                        </select>
                      </label>
                      <label>Rows per page
                        <select id="armTop">
                          <option>25</option><option selected>100</option><option>250</option><option>500</option>
                        </select>
                      </label>
                      <label>Include inactive
                        <select id="armIncludeInactive">
                          <option value="yes">Yes</option>
                          <option value="no">No, statecode = 0</option>
                        </select>
                      </label>
                    </div>
                  </div>
                </div>

                <div class="arm-condition-section">
                  <div class="arm-block-title">4. Search conditions</div>
                  <div class="arm-block-subtitle">Add one or more filters and combine them with AND / OR.</div>
                  <div id="armConditions" style="margin-top:10px"></div>
                  <div class="arm-actions" style="margin-top:10px">
                    <button class="arm-btn" id="armAddCondition" disabled>+ Add condition</button>
                  </div>
                </div>

                <div class="arm-actions-bar">
                  <div class="arm-status-wrap">
                    <div class="arm-status-label">Current status</div>
                    <div class="arm-status" id="armStatus">Loading table metadata...</div>
                  </div>
                  <div class="arm-actions">
                    <button class="arm-btn arm-primary-soft" id="armCopyQuery">Copy query</button>
                    <button class="arm-btn arm-primary" id="armRun" disabled>Run search</button>
                  </div>
                </div>

                <div class="arm-query-wrap">
                  <div class="arm-query-title">Generated Web API query</div>
                  <textarea class="arm-query" id="armQuery" readonly></textarea>
                </div>
              </section>

              <section class="arm-results-panel">
                <div class="arm-results-head">
                  <div>
                    <div class="arm-results-title">Results</div>
                    <div class="arm-results-subtitle">The table below shows the retrieved Dataverse records.</div>
                  </div>
                  <span class="arm-badge" id="armResultBadge">0 rows</span>
                </div>
                <div class="arm-results-toolbar">
                  <div class="arm-status" id="armCount">0 rows</div>
                  <div class="arm-actions">
                    <button class="arm-btn" id="armPrev" disabled>Previous</button>
                    <button class="arm-btn" id="armNext" disabled>Next page</button>
                  </div>
                </div>
                <div class="arm-results" id="armResults">
                  <div class="arm-empty">Choose a table and build your search.</div>
                </div>
              </section>
            </div>
          </div>
        `;
        document.body.appendChild(overlay);

        const $ = (id) => overlay.querySelector(`#${id}`);
        const entitySearch = $("armEntitySearch");
        const entitySelect = $("armEntitySelect");
        const fieldSearch = $("armFieldSearch");
        const columnsSelect = $("armColumns");
        const selectedColumnsSummary = $("armSelectedColumnsSummary");
        const conditionsWrap = $("armConditions");
        const addConditionBtn = $("armAddCondition");
        const orderField = $("armOrderField");
        const orderDirection = $("armOrderDirection");
        const topSelect = $("armTop");
        const includeInactive = $("armIncludeInactive");
        const queryBox = $("armQuery");
        const resultsWrap = $("armResults");
        const status = $("armStatus");
        const count = $("armCount");
        const resultBadge = $("armResultBadge");
        const runBtn = $("armRun");
        const nextBtn = $("armNext");
        const prevBtn = $("armPrev");

        let allEntities = [];
        let filteredEntities = [];
        let selectedEntity = null;
        let allFields = [];
        let filteredFields = [];
        let nextLink = null;
        let history = [];
        let currentPageUrl = null;
        let currentRows = [];
        let loadToken = 0;
        const optionCache = new Map();
        const selectedColumnNames = new Set();

        const setStatus = (message, type = "normal") => {
          const colors = {
            normal: "#ffffff",
            success: "#bbf7d0",
            warning: "#fde68a",
            error: "#fecaca"
          };
          status.textContent = message;
          status.style.color = colors[type] || colors.normal;
        };

        const setCountText = (text) => {
          count.textContent = text;
          if (resultBadge) resultBadge.textContent = text;
        };

        const entityText = (entity) =>
          `${entity.displayName || "(No display name)"} | ${entity.logicalName}`;

        const fieldText = (field) =>
          `${field.displayName || "(No display name)"} | ${field.logicalName} [${field.type}]`;

        const renderEntities = () => {
          entitySelect.innerHTML = "";
          filteredEntities.forEach((entity, index) => {
            const option = document.createElement("option");
            option.value = String(index);
            option.textContent = entityText(entity);
            entitySelect.appendChild(option);
          });
        };

        const updateSelectedColumnsSummary = () => {
          const selected = allFields
            .filter((field) => selectedColumnNames.has(field.logicalName))
            .map((field) => field.displayName || field.logicalName);

          if (!selected.length) {
            selectedColumnsSummary.textContent = "0 columns selected";
            selectedColumnsSummary.title = "";
            return;
          }

          const preview = selected.slice(0, 5).join(", ");
          const more = selected.length > 5 ? ` +${selected.length - 5} more` : "";
          selectedColumnsSummary.textContent = `${selected.length} columns selected: ${preview}${more}`;
          selectedColumnsSummary.title = selected.join("\n");
        };

        const syncVisibleColumnSelections = () => {
          [...columnsSelect.options].forEach((option) => {
            if (option.selected) {
              selectedColumnNames.add(option.value);
            } else {
              selectedColumnNames.delete(option.value);
            }
          });
          updateSelectedColumnsSummary();
        };

        const renderFields = () => {
          columnsSelect.innerHTML = "";
          filteredFields.forEach((field) => {
            const option = document.createElement("option");
            option.value = field.logicalName;
            option.textContent = fieldText(field);
            option.selected = selectedColumnNames.has(field.logicalName);
            columnsSelect.appendChild(option);
          });
          updateSelectedColumnsSummary();
        };

        const renderOrderFields = () => {
          orderField.innerHTML = `<option value="">No sorting</option>`;
          allFields
            .filter((field) => field.isValidForRead && field.type !== "Virtual")
            .forEach((field) => {
              const option = document.createElement("option");
              option.value = field.logicalName;
              option.textContent = fieldText(field);
              orderField.appendChild(option);
            });
        };

        const normalizeType = (attribute) => {
          const raw =
            attribute.AttributeTypeName?.Value ||
            attribute.AttributeType ||
            "Unknown";
          return String(raw).replace(/Type$/i, "");
        };

        const isStringType = (type) =>
          ["String", "Memo", "EntityName"].includes(type);
        const isNumberType = (type) =>
          ["Integer", "BigInt", "Decimal", "Double", "Money"].includes(type);
        const isChoiceType = (type) =>
          ["Picklist", "State", "Status", "Boolean", "MultiSelectPicklist"].includes(type);
        const isDateType = (type) => type === "DateTime";
        const isLookupType = (type) => ["Lookup", "Customer", "Owner"].includes(type);

        const operatorsFor = (field) => {
          const baseNull = [
            ["eq", "Equals"],
            ["ne", "Not equal"],
            ["null", "Is null"],
            ["notnull", "Is not null"]
          ];

          if (isStringType(field.type)) {
            return [
              ...baseNull,
              ["contains", "Contains"],
              ["notcontains", "Does not contain"],
              ["startswith", "Starts with"],
              ["endswith", "Ends with"]
            ];
          }

          if (isNumberType(field.type) || isDateType(field.type)) {
            return [
              ...baseNull,
              ["gt", "Greater than"],
              ["ge", "Greater or equal"],
              ["lt", "Less than"],
              ["le", "Less or equal"]
            ];
          }

          if (isChoiceType(field.type) || isLookupType(field.type) || field.type === "Uniqueidentifier") {
            return baseNull;
          }

          return baseNull;
        };

        const getField = (logicalName) =>
          allFields.find((field) => field.logicalName === logicalName);

        const loadChoiceOptions = async (field) => {
          if (!selectedEntity || !field || !isChoiceType(field.type)) return [];
          const key = `${selectedEntity.logicalName}:${field.logicalName}`;
          if (optionCache.has(key)) return optionCache.get(key);

          let cast = "PicklistAttributeMetadata";
          if (field.type === "Boolean") cast = "BooleanAttributeMetadata";
          if (field.type === "State") cast = "StateAttributeMetadata";
          if (field.type === "Status") cast = "StatusAttributeMetadata";
          if (field.type === "MultiSelectPicklist") cast = "MultiSelectPicklistAttributeMetadata";

          try {
            const result = await apiGet(
              `EntityDefinitions(LogicalName='${selectedEntity.logicalName}')/Attributes(LogicalName='${field.logicalName}')/Microsoft.Dynamics.CRM.${cast}?$select=LogicalName&$expand=OptionSet`
            );

            const rawOptions = result?.OptionSet?.Options || [];
            const options = rawOptions.map((item) => ({
              value: item.Value,
              label: labelOf(item.Label) || String(item.Value)
            }));

            if (field.type === "Boolean" && !options.length) {
              const trueOption = result?.OptionSet?.TrueOption;
              const falseOption = result?.OptionSet?.FalseOption;
              if (falseOption) options.push({ value: falseOption.Value, label: labelOf(falseOption.Label) || "No" });
              if (trueOption) options.push({ value: trueOption.Value, label: labelOf(trueOption.Label) || "Yes" });
            }

            optionCache.set(key, options);
            return options;
          } catch (error) {
            console.warn("Could not load choice options", field, error);
            optionCache.set(key, []);
            return [];
          }
        };

        const makeValueControl = async (field, currentValue = "") => {
          if (!field) {
            const input = document.createElement("input");
            input.disabled = true;
            return input;
          }

          if (isChoiceType(field.type)) {
            const select = document.createElement("select");
            select.innerHTML = `<option value="">Select value...</option>`;
            const options = await loadChoiceOptions(field);
            options.forEach((item) => {
              const option = document.createElement("option");
              option.value = String(item.value);
              option.textContent = `${item.label} (${item.value})`;
              option.selected = String(item.value) === String(currentValue);
              select.appendChild(option);
            });
            if (!options.length) {
              const input = document.createElement("input");
              input.placeholder = "Numeric option value";
              input.value = currentValue;
              return input;
            }
            return select;
          }

          const input = document.createElement("input");
          input.value = currentValue;

          if (isDateType(field.type)) {
            input.type = "datetime-local";
          } else if (isNumberType(field.type)) {
            input.type = "number";
            input.step = "any";
          } else if (isLookupType(field.type) || field.type === "Uniqueidentifier") {
            input.placeholder = "GUID";
          } else {
            input.placeholder = "Value";
          }

          return input;
        };

        const updateConditionRow = async (row, preserveValue = false) => {
          const fieldSelect = row.querySelector(".arm-condition-field");
          const operatorSelect = row.querySelector(".arm-condition-operator");
          const valueHost = row.querySelector(".arm-condition-value");
          const field = getField(fieldSelect.value);
          const oldValue = preserveValue
            ? valueHost.querySelector("input,select")?.value || ""
            : "";

          operatorSelect.innerHTML = "";
          operatorsFor(field || { type: "Unknown" }).forEach(([value, text]) => {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = text;
            operatorSelect.appendChild(option);
          });

          valueHost.innerHTML = "";
          valueHost.appendChild(await makeValueControl(field, oldValue));
          updateConditionValueState(row);
          updateQueryPreview();
        };

        const updateConditionValueState = (row) => {
          const operator = row.querySelector(".arm-condition-operator")?.value;
          const control = row.querySelector(".arm-condition-value input, .arm-condition-value select");
          if (!control) return;
          const noValue = operator === "null" || operator === "notnull";
          control.disabled = noValue;
          control.style.opacity = noValue ? ".45" : "1";
        };

        const addCondition = async () => {
          const row = document.createElement("div");
          row.className = "arm-condition";
          row.innerHTML = `
            <label>Join
              <select class="arm-condition-join">
                <option value="and">AND</option>
                <option value="or">OR</option>
              </select>
            </label>
            <label>Column
              <select class="arm-condition-field"></select>
            </label>
            <label>Operator
              <select class="arm-condition-operator"></select>
            </label>
            <label>Value
              <span class="arm-condition-value"></span>
            </label>
            <button class="arm-remove" title="Remove">✕</button>
          `;

          const fieldSelect = row.querySelector(".arm-condition-field");
          allFields
            .filter((field) => field.isValidForRead)
            .forEach((field) => {
              const option = document.createElement("option");
              option.value = field.logicalName;
              option.textContent = fieldText(field);
              fieldSelect.appendChild(option);
            });

          row.querySelector(".arm-condition-join").disabled =
            conditionsWrap.children.length === 0;

          fieldSelect.addEventListener("change", () => updateConditionRow(row));
          row.querySelector(".arm-condition-operator").addEventListener("change", () => {
            updateConditionValueState(row);
            updateQueryPreview();
          });
          row.querySelector(".arm-condition-join").addEventListener("change", updateQueryPreview);
          row.querySelector(".arm-condition-value").addEventListener("input", updateQueryPreview);
          row.querySelector(".arm-condition-value").addEventListener("change", updateQueryPreview);
          row.querySelector(".arm-remove").addEventListener("click", () => {
            row.remove();
            [...conditionsWrap.children].forEach((child, index) => {
              child.querySelector(".arm-condition-join").disabled = index === 0;
            });
            updateQueryPreview();
          });

          conditionsWrap.appendChild(row);
          await updateConditionRow(row);
        };

        const escapeString = (value) => String(value).replace(/'/g, "''");
        const cleanGuid = (value) => String(value || "").replace(/[{}]/g, "").trim();

        const formatValue = (field, value) => {
          if (isNumberType(field.type)) return String(Number(value));
          if (field.type === "Boolean") return String(value).toLowerCase() === "true" || String(value) === "1" ? "true" : "false";
          if (isChoiceType(field.type)) return String(Number(value));
          if (isLookupType(field.type)) return cleanGuid(value);
          if (field.type === "Uniqueidentifier") return cleanGuid(value);
          if (isDateType(field.type)) {
            const date = new Date(value);
            if (Number.isNaN(date.getTime())) throw new Error(`Invalid date for ${field.logicalName}`);
            return date.toISOString();
          }
          return `'${escapeString(value)}'`;
        };

        const conditionExpression = (row) => {
          const field = getField(row.querySelector(".arm-condition-field").value);
          const operator = row.querySelector(".arm-condition-operator").value;
          const value = row.querySelector(".arm-condition-value input, .arm-condition-value select")?.value ?? "";
          if (!field) return null;

          const property = isLookupType(field.type)
            ? `_${field.logicalName}_value`
            : field.logicalName;

          if (operator === "null") return `${property} eq null`;
          if (operator === "notnull") return `${property} ne null`;
          if (!String(value).trim()) throw new Error(`Enter a value for ${field.displayName || field.logicalName}`);

          const formatted = formatValue(field, value);
          if (operator === "contains") return `contains(${property},${formatted})`;
          if (operator === "notcontains") return `not contains(${property},${formatted})`;
          if (operator === "startswith") return `startswith(${property},${formatted})`;
          if (operator === "endswith") return `endswith(${property},${formatted})`;
          return `${property} ${operator} ${formatted}`;
        };

        const selectedColumns = () =>
          allFields
            .filter((field) => selectedColumnNames.has(field.logicalName))
            .map((field) => field.logicalName);

        const buildRelativeQuery = () => {
          if (!selectedEntity) return "";
          const params = [];
          const cols = selectedColumns();
          if (cols.length) params.push(`$select=${cols.join(",")}`);

          const expressions = [];
          [...conditionsWrap.children].forEach((row, index) => {
            const expression = conditionExpression(row);
            if (!expression) return;
            const join = index === 0
              ? ""
              : ` ${row.querySelector(".arm-condition-join").value} `;
            expressions.push(`${join}${expression}`);
          });

          if (includeInactive.value === "no" && getField("statecode")) {
            expressions.push(`${expressions.length ? " and " : ""}statecode eq 0`);
          }

          if (expressions.length) params.push(`$filter=${expressions.join("")}`);
          if (orderField.value) params.push(`$orderby=${orderField.value} ${orderDirection.value}`);
          params.push(`$top=${Number(topSelect.value) || 100}`);

          return `${selectedEntity.entitySetName}?${params.join("&")}`;
        };

        const updateQueryPreview = () => {
          try {
            const relative = buildRelativeQuery();
            queryBox.value = relative ? `${API}/${relative}` : "";
            runBtn.disabled = !relative;
          } catch (error) {
            queryBox.value = `Query is incomplete: ${error.message}`;
          }
        };

        const pickRecommendedColumns = () => {
          const preferred = [
            selectedEntity?.primaryIdAttribute,
            selectedEntity?.primaryNameAttribute,
            "statecode",
            "statuscode",
            "createdon",
            "modifiedon",
            "ownerid"
          ].filter(Boolean);

          const existing = new Set(allFields.map((field) => field.logicalName));
          const final = preferred.filter((name) => existing.has(name));
          selectedColumnNames.clear();
          final.forEach((name) => selectedColumnNames.add(name));
          renderFields();
          updateQueryPreview();
        };

        const loadEntities = async () => {
          const response = await apiGet(
            "EntityDefinitions?$select=LogicalName,SchemaName,EntitySetName,DisplayName,PrimaryIdAttribute,PrimaryNameAttribute,IsPrivate,IsIntersect,IsActivity"
          );

          allEntities = (response.value || [])
            .filter((item) => item.LogicalName && item.EntitySetName && !item.IsPrivate)
            .map((item) => ({
              logicalName: item.LogicalName,
              schemaName: item.SchemaName || "",
              entitySetName: item.EntitySetName,
              displayName: labelOf(item.DisplayName),
              primaryIdAttribute: item.PrimaryIdAttribute,
              primaryNameAttribute: item.PrimaryNameAttribute,
              isIntersect: Boolean(item.IsIntersect),
              isActivity: Boolean(item.IsActivity)
            }))
            .sort((a, b) => entityText(a).localeCompare(entityText(b)));

          filteredEntities = [...allEntities];
          renderEntities();
          setStatus(`Loaded ${allEntities.length} tables. Select one.`, "success");
        };

        const loadFields = async (entity) => {
          const token = ++loadToken;
          selectedEntity = entity;
          setStatus(`Loading columns for ${entity.displayName || entity.logicalName}...`);
          runBtn.disabled = true;
          fieldSearch.disabled = true;
          columnsSelect.disabled = true;
          addConditionBtn.disabled = true;
          orderField.disabled = true;
          conditionsWrap.innerHTML = "";
          fieldSearch.value = "";
          selectedColumnNames.clear();
          updateSelectedColumnsSummary();

          const response = await apiGet(
            `EntityDefinitions(LogicalName='${entity.logicalName}')/Attributes?` +
            "$select=LogicalName,SchemaName,DisplayName,AttributeType,AttributeTypeName,IsValidForRead,IsPrimaryId,IsPrimaryName,IsLogical,IsSecured"
          );
          if (token !== loadToken) return;

          allFields = (response.value || [])
            .filter((item) => item.LogicalName && item.IsValidForRead !== false)
            .map((item) => ({
              logicalName: item.LogicalName,
              schemaName: item.SchemaName || "",
              displayName: labelOf(item.DisplayName),
              type: normalizeType(item),
              isValidForRead: item.IsValidForRead !== false,
              isPrimaryId: Boolean(item.IsPrimaryId),
              isPrimaryName: Boolean(item.IsPrimaryName),
              isLogical: Boolean(item.IsLogical),
              isSecured: Boolean(item.IsSecured)
            }))
            .sort((a, b) => fieldText(a).localeCompare(fieldText(b)));

          filteredFields = [...allFields];
          renderFields();
          renderOrderFields();
          fieldSearch.disabled = false;
          columnsSelect.disabled = false;
          addConditionBtn.disabled = false;
          orderField.disabled = false;
          pickRecommendedColumns();
          await addCondition();
          setStatus(`Loaded ${allFields.length} columns for ${entityText(entity)}.`, "success");
          updateQueryPreview();
        };

        const displayValue = (row, fieldName) => {
          const formattedKey = `${fieldName}@OData.Community.Display.V1.FormattedValue`;
          if (row[formattedKey] !== undefined) return row[formattedKey];
          const field = getField(fieldName);
          if (field && isLookupType(field.type)) {
            const lookupKey = `_${fieldName}_value`;
            return row[`${lookupKey}@OData.Community.Display.V1.FormattedValue`] ?? row[lookupKey] ?? "";
          }
          const value = row[fieldName];
          if (value === null || value === undefined) return "";
          if (typeof value === "object") return JSON.stringify(value);
          return String(value);
        };

        const recordId = (row) => row[selectedEntity.primaryIdAttribute];

        const renderResults = (rows) => {
          currentRows = rows;
          const columns = selectedColumns().length
            ? selectedColumns()
            : Object.keys(rows[0] || {}).filter((key) => !key.includes("@"));

          if (!rows.length) {
            resultsWrap.innerHTML = `<div class="arm-empty">No records found.</div>`;
            setCountText("0 rows");
            return;
          }

          const table = document.createElement("table");
          const thead = document.createElement("thead");
          const trh = document.createElement("tr");

          columns.forEach((column) => {
            const field = getField(column);
            const th = document.createElement("th");
            th.textContent = field?.displayName
              ? `${field.displayName} (${column})`
              : column;
            trh.appendChild(th);
          });
          thead.appendChild(trh);
          table.appendChild(thead);

          const tbody = document.createElement("tbody");
          rows.forEach((row) => {
            const tr = document.createElement("tr");
            columns.forEach((column) => {
              const td = document.createElement("td");
              const value = displayValue(row, column);
              td.title = value;

              if (column === selectedEntity.primaryIdAttribute && recordId(row)) {
                const link = document.createElement("a");
                link.className = "arm-link";
                link.href = "#";
                link.textContent = value;
                link.addEventListener("click", async (event) => {
                  event.preventDefault();
                  try {
                    await window.Xrm.Navigation.openForm({
                      entityName: selectedEntity.logicalName,
                      entityId: recordId(row),
                      openInNewWindow: true
                    });
                  } catch (error) {
                    alert(error.message);
                  }
                });
                td.appendChild(link);
              } else {
                td.textContent = value;
              }
              tr.appendChild(td);
            });
            tbody.appendChild(tr);
          });
          table.appendChild(tbody);
          resultsWrap.innerHTML = "";
          resultsWrap.appendChild(table);
          setCountText(`${rows.length} rows on this page`);
        };

        const runUrl = async (url, pushHistory = false) => {
          try {
            runBtn.disabled = true;
            nextBtn.disabled = true;
            prevBtn.disabled = true;
            setStatus("Running query...");
            resultsWrap.innerHTML = `<div class="arm-empty">Loading...</div>`;

            if (pushHistory && currentPageUrl) history.push(currentPageUrl);
            const response = await requestJson(url);
            currentPageUrl = url;
            nextLink = response?.["@odata.nextLink"] || null;
            renderResults(response?.value || []);
            nextBtn.disabled = !nextLink;
            prevBtn.disabled = history.length === 0;
            setStatus(
              `Query completed. ${response?.value?.length || 0} rows returned${nextLink ? "; more rows available" : ""}.`,
              "success"
            );
          } catch (error) {
            console.error(error);
            setStatus(error.message, "error");
            resultsWrap.innerHTML = `<div class="arm-empty" style="color:#fca5a5">${html(error.message)}</div>`;
          } finally {
            runBtn.disabled = false;
          }
        };

        entitySearch.addEventListener("input", () => {
          const term = entitySearch.value.trim().toLowerCase();
          filteredEntities = allEntities.filter((entity) =>
            [entity.displayName, entity.logicalName, entity.schemaName, entity.entitySetName]
              .some((value) => String(value || "").toLowerCase().includes(term))
          );
          renderEntities();
        });

        entitySelect.addEventListener("change", () => {
          const entity = filteredEntities[Number(entitySelect.value)];
          if (entity) loadFields(entity);
        });

        entitySelect.addEventListener("dblclick", () => {
          const entity = filteredEntities[Number(entitySelect.value)];
          if (entity) loadFields(entity);
        });

        fieldSearch.addEventListener("input", () => {
          const term = fieldSearch.value.trim().toLowerCase();
          filteredFields = allFields.filter((field) =>
            [field.displayName, field.logicalName, field.schemaName, field.type]
              .some((value) => String(value || "").toLowerCase().includes(term))
          );
          renderFields();
        });

        columnsSelect.addEventListener("change", () => {
          syncVisibleColumnSelections();
          updateQueryPreview();
        });
        orderField.addEventListener("change", updateQueryPreview);
        orderDirection.addEventListener("change", updateQueryPreview);
        topSelect.addEventListener("change", updateQueryPreview);
        includeInactive.addEventListener("change", updateQueryPreview);
        addConditionBtn.addEventListener("click", addCondition);
        $("armSelectRecommended").addEventListener("click", pickRecommendedColumns);
        $("armClearColumns").addEventListener("click", () => {
          selectedColumnNames.clear();
          renderFields();
          updateQueryPreview();
        });

        runBtn.addEventListener("click", async () => {
          try {
            const relative = buildRelativeQuery();
            if (!relative) return;
            history = [];
            currentPageUrl = null;
            await runUrl(`${API}/${relative}`);
          } catch (error) {
            setStatus(error.message, "error");
          }
        });

        nextBtn.addEventListener("click", async () => {
          if (nextLink) await runUrl(nextLink, true);
        });

        prevBtn.addEventListener("click", async () => {
          const previous = history.pop();
          if (previous) {
            currentPageUrl = null;
            await runUrl(previous, false);
            prevBtn.disabled = history.length === 0;
          }
        });

        $("armCopyQuery").addEventListener("click", async () => {
          try {
            await navigator.clipboard.writeText(queryBox.value);
            setStatus("Query copied to clipboard.", "success");
          } catch {
            queryBox.focus();
            queryBox.select();
            document.execCommand("copy");
            setStatus("Query copied to clipboard.", "success");
          }
        });

        const close = () => {
          style.remove();
          overlay.remove();
        };
        $("armCloseTop").addEventListener("click", close);

        try {
          await loadEntities();
        } catch (error) {
          console.error(error);
          setStatus(error.message, "error");
        }
      }
    });
  });








// popup.js  (FetchXML UI button - FULL CODE, pretty output as CSV)
// 1) Add a button in popup.html: <button id="fetchXmlUi">FetchXML</button>
// 2) Paste this whole block into popup.js

document.getElementById("fetchXmlUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      // --- helpers ---
      const remove = () => document.getElementById("__d365_fetchxml_modal")?.remove();

      const escapeHtml = (s) =>
        String(s ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");

      const toCsv = (rows, cols) => {
        const esc = (v) => {
          const s = String(v ?? "");
          const needs = /[",\n]/.test(s);
          const out = s.replaceAll('"', '""');
          return needs ? `"${out}"` : out;
        };
        const header = cols.map(esc).join(",");
        const lines = rows.map(r => cols.map(c => esc(r?.[c])).join(","));
        return [header, ...lines].join("\n");
      };

      const buildModal = () => {
        remove();

        const overlay = document.createElement("div");
        overlay.id = "__d365_fetchxml_modal";
        overlay.style.cssText = `
          position: fixed; inset: 0; background: rgba(0,0,0,.35);
          z-index: 2147483647; display:flex; align-items:center; justify-content:center; padding:16px;
        `;

        const box = document.createElement("div");
        box.style.cssText = `
          width: min(1200px, 96vw);
          height: min(760px, 92vh);
          background:#fff;
          border-radius:16px;
          box-shadow:0 18px 50px rgba(0,0,0,.35);
          overflow:hidden;
          font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
          direction: rtl;
          text-align: right;
          display:flex;
          flex-direction:column;
        `;

        const header = document.createElement("div");
        header.style.cssText = `padding:12px 14px; font-weight:900; border-bottom:1px solid #e5e7eb;`;
        header.textContent = "FetchXml Tester (table view)";

        const body = document.createElement("div");
        body.style.cssText = `padding:12px 14px; display:grid; gap:10px; flex:1; min-height:0;`;

        const fetchTa = document.createElement("textarea");
        fetchTa.placeholder = "Paste full FetchXML here…";
        fetchTa.style.cssText = `
          width:100%; height:140px; resize:vertical;
          border:1px solid #cbd5e1; border-radius:10px; padding:10px;
          font-size:12px; line-height:1.4; box-sizing:border-box;
          font-family: Consolas, Monaco, "Courier New", monospace;
          direction:ltr; text-align:left; white-space:pre;
        `;

        const status = document.createElement("div");
        status.style.cssText = `font-size:12px; color:#374151;`;

        const tableWrap = document.createElement("div");
        tableWrap.style.cssText = `
          border:1px solid #cbd5e1;
          border-radius:10px;
          overflow:auto;
          height: 100%;
          min-height: 260px;
        `;

        const table = document.createElement("table");
        table.style.cssText = `
          width:100%;
          border-collapse:collapse;
          font-size:12px;
          direction:ltr;
          text-align:left;
        `;
        tableWrap.appendChild(table);

        const rawTa = document.createElement("textarea");
        rawTa.readOnly = true;
        rawTa.placeholder = "Raw JSON (for copy) will appear here…";
        rawTa.style.cssText = `
          width:100%; height:140px; resize:vertical;
          border:1px solid #cbd5e1; border-radius:10px; padding:10px;
          font-size:12px; line-height:1.4; box-sizing:border-box;
          font-family: Consolas, Monaco, "Courier New", monospace;
          direction:ltr; text-align:left; white-space:pre;
        `;

        body.appendChild(fetchTa);
        body.appendChild(status);
        body.appendChild(tableWrap);
        body.appendChild(rawTa);

        const footer = document.createElement("div");
        footer.style.cssText = `
          display:flex; gap:10px; justify-content:flex-end;
          padding:12px 14px; border-top:1px solid #e5e7eb;
        `;

        const mkBtn = (text) => {
          const b = document.createElement("button");
          b.textContent = text;
          b.style.cssText = `
            border:1px solid #cbd5e1; padding:10px 14px; border-radius:10px;
            cursor:pointer; background:#fff; font-weight:900;
          `;
          return b;
        };

        const btnClose = mkBtn("Close");
        const btnCopy = mkBtn("Copy Raw");
        btnCopy.style.border = "none";
        btnCopy.style.background = "#2563eb";
        btnCopy.style.color = "#fff";

        const btnCsv = mkBtn("Copy CSV");
        btnCsv.style.border = "none";
        btnCsv.style.background = "#059669";
        btnCsv.style.color = "#fff";

        const btnRun = mkBtn("Run");
        btnRun.style.border = "none";
        btnRun.style.background = "#111827";
        btnRun.style.color = "#fff";

        const close = () => overlay.remove();
        btnClose.onclick = close;

        let lastRows = [];
        let lastCols = [];

        btnCopy.onclick = async () => {
          const text = rawTa.value || "";
          try { await navigator.clipboard.writeText(text); }
          catch { rawTa.focus(); rawTa.select(); document.execCommand("copy"); }
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy Raw"), 900);
        };

        btnCsv.onclick = async () => {
          if (!lastRows.length || !lastCols.length) return;
          const csv = toCsv(lastRows, lastCols);
          try { await navigator.clipboard.writeText(csv); }
          catch {
            rawTa.value = csv;
            rawTa.focus(); rawTa.select(); document.execCommand("copy");
          }
          btnCsv.textContent = "CSV ✅";
          setTimeout(() => (btnCsv.textContent = "Copy CSV"), 900);
        };

        const renderTable = (rows) => {
          table.innerHTML = "";
          if (!rows.length) return;

          // columns: union of keys from first 25 rows
          const colSet = new Set();
          rows.slice(0, 25).forEach(r => Object.keys(r || {}).forEach(k => colSet.add(k)));
          const cols = Array.from(colSet);
          lastCols = cols;
          lastRows = rows;

          // thead
          const thead = document.createElement("thead");
          const trh = document.createElement("tr");
          cols.forEach(c => {
            const th = document.createElement("th");
            th.innerHTML = escapeHtml(c);
            th.style.cssText = `
              position: sticky; top: 0;
              background: #0b1220;
              color: #fff;
              padding: 8px;
              border-bottom: 1px solid rgba(255,255,255,.15);
              white-space: nowrap;
            `;
            trh.appendChild(th);
          });
          thead.appendChild(trh);
          table.appendChild(thead);

          // tbody
          const tbody = document.createElement("tbody");
          rows.forEach((r, idx) => {
            const tr = document.createElement("tr");
            tr.style.background = idx % 2 === 0 ? "#ffffff" : "#f8fafc";
            cols.forEach(c => {
              const td = document.createElement("td");
              const v = r?.[c];

              let cell = v;
              if (typeof cell === "object" && cell !== null) {
                try { cell = JSON.stringify(cell); } catch { cell = String(cell); }
              }

              td.innerHTML = escapeHtml(cell ?? "");
              td.style.cssText = `
                padding: 8px;
                border-bottom: 1px solid #e5e7eb;
                max-width: 420px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
              `;
              tr.appendChild(td);
            });
            tbody.appendChild(tr);
          });
          table.appendChild(tbody);
        };

        btnRun.onclick = async () => {
          status.textContent = "";
          table.innerHTML = "";
          rawTa.value = "";
          lastRows = [];
          lastCols = [];

          const fetchXml = (fetchTa.value || "").trim();
          if (!fetchXml) {
            status.textContent = "❌ Paste FetchXML first.";
            return;
          }

          const Xrm = window.Xrm;
          const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;
          if (!webApi?.retrieveMultipleRecords) {
            status.textContent = "❌ Xrm.WebApi.retrieveMultipleRecords not available.";
            return;
          }

          // Parse entity name from fetchxml (simple regex)
          const m = fetchXml.match(/<entity\s+name="([^"]+)"/i);
          const entity = m?.[1];
          if (!entity) {
            status.textContent = "❌ Could not detect entity name from <entity name=\"...\">";
            return;
          }

          status.textContent = "⏳ Running…";

          try {
            const encoded = encodeURIComponent(fetchXml);
            const res = await webApi.retrieveMultipleRecords(entity, `?fetchXml=${encoded}`);
            const rows = res?.entities || [];

            status.textContent = `✅ Entity: ${entity} | Returned: ${rows.length}`;
            rawTa.value = JSON.stringify(rows, null, 2);

            // render table (up to 5000 returned anyway)
            renderTable(rows);

            // auto select raw for easy copy if you want:
            // rawTa.focus(); rawTa.select();
          } catch (err) {
            status.textContent = "❌ Failed.";
            rawTa.value = (err?.message || err?.toString?.() || "Unknown error");
          }
        };

        footer.appendChild(btnClose);
        footer.appendChild(btnCopy);
        footer.appendChild(btnCsv);
        footer.appendChild(btnRun);

        box.appendChild(header);
        box.appendChild(body);
        box.appendChild(footer);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        fetchTa.focus();
      };

      buildModal();
    }
  });
});











document.getElementById("findLogicalByLabel").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      // remove existing modal
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(900px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 800; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "Find Logical Name by Display Name";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const labelInput = document.createElement("input");
      labelInput.placeholder = "Enter Display Name (Label) from the form...";
      labelInput.style.cssText = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Results will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 280px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: rtl;
        text-align: right;
      `;

      body.appendChild(labelInput);
      body.appendChild(status);
      body.appendChild(resultTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 14px;
        border-top: 1px solid #e5e7eb;
      `;

      const btn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 800;
        `;
        return b;
      };

      const btnClose = btn("Close");

      const btnCopy = btn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnSearch = btn("Search");
      btnSearch.style.border = "none";
      btnSearch.style.background = "#111827";
      btnSearch.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch (e) {
          resultTa.focus();
          resultTa.select();
          document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      const runSearch = () => {
        const q = (labelInput.value || "").trim().toLowerCase();
        resultTa.value = "";
        status.textContent = "";

        if (!q) {
          status.textContent = "❌ Enter display name.";
          return;
        }

        const Xrm = window.Xrm;
        const page = Xrm?.Page;

        if (!Xrm || !page?.ui?.controls) {
          status.textContent = "❌ Xrm not found. Open a record form first.";
          return;
        }

        const matches = [];

        page.ui.controls.forEach(function (c) {
          try {
            if (!c || !c.getLabel || !c.getName) return;

            const label = (c.getLabel() || "").trim();
            const name = (c.getName() || "").trim();
            if (!label || !name) return;

            if (label.toLowerCase().includes(q)) {
              matches.push({
                label,
                logicalName: name,
                controlType: c.getControlType ? c.getControlType() : ""
              });
            }
          } catch (e) {}
        });

        if (!matches.length) {
          status.textContent = "⚠️ No matches found.";
          return;
        }

        // pretty align
        const labelMax = Math.max(...matches.map(m => m.label.length), 5);
        const nameMax = Math.max(...matches.map(m => m.logicalName.length), 5);

        const pad = (s, n) => (s + " ".repeat(n)).slice(0, n);

        const lines = [];
        lines.push(`Found ${matches.length} match(es)\n`);
        lines.push(`${pad("Display Name", labelMax)}  =>  ${pad("Logical Name", nameMax)}  | Type`);
        lines.push(`${"-".repeat(labelMax)}  =>  ${"-".repeat(nameMax)}  | ----`);

        matches.forEach(m => {
          lines.push(`${pad(m.label, labelMax)}  =>  ${pad(m.logicalName, nameMax)}  | ${m.controlType}`);
        });

        resultTa.value = lines.join("\n");
        status.textContent = "✅ Done. You can copy the results.";
        resultTa.focus();
        resultTa.select();
      };

      btnSearch.onclick = runSearch;

      // ENTER = search
      labelInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runSearch();
      });

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnSearch);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      labelInput.focus();
    }
  });
});
document.getElementById("findLabelByLogical").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(900px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 800; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "Display Name (Label) by Logical Name";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const input = document.createElement("input");
      input.placeholder = "Enter logical name(s), comma separated... e.g. firstname,lastname,emailaddress1";
      input.style.cssText = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Results will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 280px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: rtl;
        text-align: right;
      `;

      body.appendChild(input);
      body.appendChild(status);
      body.appendChild(resultTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 14px;
        border-top: 1px solid #e5e7eb;
      `;

      const btn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 800;
        `;
        return b;
      };

      const btnClose = btn("Close");

      const btnCopy = btn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnSearch = btn("Search");
      btnSearch.style.border = "none";
      btnSearch.style.background = "#111827";
      btnSearch.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch (e) {
          resultTa.focus();
          resultTa.select();
          document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      const runSearch = () => {
        status.textContent = "";
        resultTa.value = "";

        const raw = (input.value || "").trim();
        if (!raw) {
          status.textContent = "❌ Enter at least 1 logical name.";
          return;
        }

        const logicalNames = raw
          .split(",")
          .map(s => s.trim())
          .filter(Boolean);

        const Xrm = window.Xrm;
        const page = Xrm?.Page;

        if (!Xrm || !page?.ui?.controls) {
          status.textContent = "❌ Xrm not found. Open a record form first.";
          return;
        }

        const results = [];

        for (const ln of logicalNames) {
          try {
            const ctrl = page.getControl?.(ln) || null;

            if (!ctrl) {
              results.push({ logicalName: ln, label: "(not on form)" });
              continue;
            }

            const label = (ctrl.getLabel?.() || "").trim();
            results.push({
              logicalName: ln,
              label: label || "(no label)"
            });
          } catch (e) {
            results.push({ logicalName: ln, label: "(error)" });
          }
        }

        const maxLabel = Math.max(...results.map(r => r.label.length), 10);
        const maxLn = Math.max(...results.map(r => r.logicalName.length), 10);
        const pad = (s, n) => (s + " ".repeat(n)).slice(0, n);

        const lines = [];
        lines.push(`Found ${results.length} field(s)\n`);
        lines.push(`${pad("Logical Name", maxLn)}  =>  ${pad("Display Name", maxLabel)}`);
        lines.push(`${"-".repeat(maxLn)}  =>  ${"-".repeat(maxLabel)}`);

        results.forEach(r => {
          lines.push(`${pad(r.logicalName, maxLn)}  =>  ${pad(r.label, maxLabel)}`);
        });

        resultTa.value = lines.join("\n");
        status.textContent = "✅ Done. You can copy.";
        resultTa.focus();
        resultTa.select();
      };

      btnSearch.onclick = runSearch;

      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runSearch();
      });

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnSearch);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      input.focus();
    }
  });
});
document.getElementById("getSystemParam").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(900px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 800; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "Get System Param (ey_system_params)";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const nameInput = document.createElement("input");
      nameInput.placeholder = "Enter ey_name (example: MyParamName)";
      nameInput.style.cssText = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Result will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 240px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: ltr;
        text-align: left;
      `;

      body.appendChild(nameInput);
      body.appendChild(status);
      body.appendChild(resultTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 14px;
        border-top: 1px solid #e5e7eb;
      `;

      const btn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 800;
        `;
        return b;
      };

      const btnClose = btn("Close");

      const btnCopy = btn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnGet = btn("Get");
      btnGet.style.border = "none";
      btnGet.style.background = "#111827";
      btnGet.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch (e) {
          resultTa.focus();
          resultTa.select();
          document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      const runGet = async () => {
        status.textContent = "";
        resultTa.value = "";

        const eyName = (nameInput.value || "").trim();
        if (!eyName) {
          status.textContent = "❌ ey_name is required.";
          return;
        }

        const Xrm = window.Xrm;
        const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;

        if (!webApi?.retrieveMultipleRecords) {
          status.textContent = "❌ Xrm.WebApi.retrieveMultipleRecords not available.";
          return;
        }

        status.textContent = "⏳ Loading…";

        try {
          // escape single quotes for OData
          const safeName = eyName.replace(/'/g, "''");

          const query =
            `?$select=ey_name,ey_value` +
            `&$filter=ey_name eq '${safeName}'` +
            `&$top=5`;

          const res = await webApi.retrieveMultipleRecords("ey_system_params", query);
          const rows = res?.entities || [];

          if (!rows.length) {
            status.textContent = "⚠️ Not found.";
            resultTa.value = `No system param found for ey_name = "${eyName}"`;
            return;
          }

          // show all matches (sometimes there are duplicates)
          const lines = [];
          lines.push(`Entity: ey_system_params`);
          lines.push(`Filter: ey_name = "${eyName}"`);
          lines.push(`Found: ${rows.length}`);
          lines.push("");

          rows.forEach((r, i) => {
            const val = r.ey_value ?? "";
            const name = r.ey_name ?? "";
            const id = r.ey_system_paramsid || r.ey_system_paramid || r.ey_system_params_id || "(id not returned)";
            lines.push(`${i + 1}) ey_name  = ${name}`);
            lines.push(`   ey_value = ${val}`);
            lines.push(`   id       = ${id}`);
            lines.push("");
          });

          resultTa.value = lines.join("\n");
          status.textContent = "✅ Done.";
          resultTa.focus();
          resultTa.select();
        } catch (err) {
          status.textContent = "❌ Failed.";
          resultTa.value =
            "ERROR:\n" +
            (err?.message || err?.toString?.() || "Unknown error") +
            "\n\nTip: verify the entity is correct: ey_system_params and fields ey_name / ey_value.";
        }
      };

      btnGet.onclick = runGet;
      nameInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runGet();
      });

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnGet);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      nameInput.focus();
    }
  });
});
document.getElementById("shareExt").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      const GITHUB_URL = "https://github.com/roivaldman1/D365-Extention/tree/main"; // <-- put your link

      // remove existing modal
      document.getElementById("__d365helper_share_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_share_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(780px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
        direction: rtl;
        text-align: right;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 900; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "Share / Install Instructions";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const text = [
        "📌 איך להתקין את התוסף:",
        "1) פתח Chrome והיכנס ל: chrome://extensions",
        "2) הפעל Developer mode (בפינה למעלה)",
        "3) לחץ Load unpacked",
        "4) בחר את תיקיית הפרויקט של התוסף",
        "",
        "✅ דרך GitHub:",
        `- פתח את הריפו: ${GITHUB_URL}`,
        "- או clone:",
        `  git clone ${GITHUB_URL}`,
        "",
        "לאחר מכן: Load unpacked על התיקייה שנוצרה."
      ].join("\n");

      const ta = document.createElement("textarea");
      ta.readOnly = true;
      ta.value = text;
      ta.style.cssText = `
        width: 100%;
        height: 260px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.5;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: ltr;
        text-align: left;
      `;

      const linkRow = document.createElement("div");
      linkRow.style.cssText = `display:flex; gap:10px; justify-content: space-between; align-items:center;`;

      const a = document.createElement("a");
      a.href = GITHUB_URL;
      a.textContent = "Open GitHub Repo";
      a.target = "_blank";
      a.rel = "noreferrer";
      a.style.cssText = `
        font-weight: 900;
        color: #2563eb;
        text-decoration: none;
      `;

      const hint = document.createElement("div");
      hint.textContent = "Tip: Ctrl+A ואז Ctrl+C כדי להעתיק הכול";
      hint.style.cssText = `font-size:12px; color:#6b7280;`;

      linkRow.appendChild(hint);
      linkRow.appendChild(a);

      body.appendChild(linkRow);
      body.appendChild(ta);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex; gap: 10px; justify-content: flex-end;
        padding: 12px 14px; border-top: 1px solid #e5e7eb;
      `;

      const mkBtn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 900;
        `;
        return b;
      };

      const btnClose = mkBtn("Close");

      const btnCopy = mkBtn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnOpen = mkBtn("Open GitHub");
      btnOpen.style.border = "none";
      btnOpen.style.background = "#111827";
      btnOpen.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(ta.value);
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch {
          ta.focus(); ta.select(); document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      btnOpen.onclick = () => window.open(GITHUB_URL, "_blank", "noreferrer");

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnOpen);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      // Auto-select for easy copy
      ta.focus();
      ta.select();
    }
  });
});
document.getElementById("openAdvancedFind").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id || !tab?.url) return;

  try {
    const u = new URL(tab.url);

    // ✅ org base: https://mhcsd.crm4.dynamics.com
    const base = `${u.protocol}//${u.host}`;

    // ✅ get appid from current URL
    const appid = u.searchParams.get("appid");
    if (!appid) {
      alert("appid not found in current URL.\nOpen a D365 record with appid=... and try again.");
      return;
    }

    // ✅ hard-coded rest
    const advancedUrl = `${base}/main.aspx?appid=${appid}&pagetype=AdvancedFind#292681398`;

    chrome.tabs.create({ url: advancedUrl });
  } catch (e) {
    alert("Failed to open Advanced Find.\n" + String(e));
  }
});
document.getElementById("searchSystemParams").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(1000px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
        direction: rtl;
        text-align: right;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 900; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "Search ey_system_params by string inside ey_value";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

      const input = document.createElement("input");
      input.placeholder = "Enter text to search inside ey_value (example: DirectDebit)";
      input.style.cssText = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Results will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 420px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: ltr;
        text-align: left;
      `;

      body.appendChild(input);
      body.appendChild(status);
      body.appendChild(resultTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 14px;
        border-top: 1px solid #e5e7eb;
      `;

      const mkBtn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 900;
        `;
        return b;
      };

      const btnClose = mkBtn("Close");

      const btnCopy = mkBtn("Copy");
      btnCopy.style.border = "none";
      btnCopy.style.background = "#2563eb";
      btnCopy.style.color = "#fff";

      const btnSearch = mkBtn("Search");
      btnSearch.style.border = "none";
      btnSearch.style.background = "#111827";
      btnSearch.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch {
          resultTa.focus();
          resultTa.select();
          document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      const prettyJsonIfPossible = (valRaw) => {
        if (valRaw == null) return "";
        const s = String(valRaw);

        try {
          const obj = JSON.parse(s);
          return JSON.stringify(obj, null, 2);
        } catch {
          return s;
        }
      };

      const runSearch = async () => {
        status.textContent = "";
        resultTa.value = "";

        const text = (input.value || "").trim();
        if (!text) {
          status.textContent = "❌ Enter a string to search.";
          return;
        }

        const Xrm = window.Xrm;
        const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;

        if (!webApi?.retrieveMultipleRecords) {
          status.textContent = "❌ Xrm.WebApi.retrieveMultipleRecords not available.";
          return;
        }

        status.textContent = "⏳ Searching…";

        try {
          // contains(field,'text') needs single quotes escaped
          const safeText = text.replace(/'/g, "''");

          // NOTE: Dataverse string contains is case-insensitive in most environments,
          // depends on DB collation, but usually works fine.
          const query =
            `?$select=ey_name,ey_value` +
            `&$filter=contains(ey_value,'${safeText}')` +
            `&$top=5000`;

          const res = await webApi.retrieveMultipleRecords("ey_system_params", query);
          const rows = res?.entities || [];

          if (!rows.length) {
            status.textContent = "⚠️ No matches.";
            resultTa.value = `No ey_system_params found where ey_value contains: "${text}"`;
            return;
          }

          const lines = [];
          lines.push(`Entity: ey_system_params`);
          lines.push(`Search: ey_value contains "${text}"`);
          lines.push(`Found: ${rows.length}`);
          lines.push("");

          rows.forEach((r, i) => {
            const eyName = r.ey_name ?? "";
            const eyValueRaw = r.ey_value ?? "";

            lines.push(`${i + 1}) ey_name  = ${eyName}`);
            lines.push(`   ey_value =`);

            const pretty = prettyJsonIfPossible(eyValueRaw);
            pretty.split("\n").forEach(line => lines.push("   " + line));

            lines.push("");
          });

          resultTa.value = lines.join("\n");
          status.textContent = `✅ Done (${rows.length}).`;
          resultTa.focus();
          resultTa.select();
        } catch (err) {
          status.textContent = "❌ Failed.";
          resultTa.value =
            "ERROR:\n" +
            (err?.message || err?.toString?.() || "Unknown error") +
            "\n\nTip: If contains() is blocked in your environment, tell me and I’ll switch to FetchXML search.";
        }
      };

      btnSearch.onclick = runSearch;

      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runSearch();
      });

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
      footer.appendChild(btnSearch);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      input.focus();
    }
  });
});
document.getElementById("showDirtyFields").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      // ---------- modal ----------
      const remove = () => document.getElementById("__d365_dirty_modal")?.remove();

      const openModal = (text, title) => {
        remove();

        const overlay = document.createElement("div");
        overlay.id = "__d365_dirty_modal";
        overlay.style.cssText = `
          position: fixed; inset: 0; background: rgba(0,0,0,.35);
          z-index: 2147483647; display:flex; align-items:center; justify-content:center; padding:16px;
        `;

        const box = document.createElement("div");
        box.style.cssText = `
          width: min(900px, 96vw);
          background:#fff;
          border-radius:14px;
          box-shadow:0 18px 50px rgba(0,0,0,.35);
          overflow:hidden;
          font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
          direction: rtl;
          text-align: right;
        `;

        const header = document.createElement("div");
        header.style.cssText = `padding:12px 14px; font-weight:900; border-bottom:1px solid #e5e7eb;`;
        header.textContent = title || "Dirty Fields";

        const body = document.createElement("div");
        body.style.cssText = `padding:12px 14px; display:grid; gap:10px;`;

        const ta = document.createElement("textarea");
        ta.readOnly = true;
        ta.value = text || "";
        ta.style.cssText = `
          width:100%;
          height:420px;
          resize:vertical;
          border:1px solid #cbd5e1;
          border-radius:10px;
          padding:10px;
          font-size:12px;
          line-height:1.4;
          white-space:pre;
          box-sizing:border-box;
          font-family: Consolas, Monaco, "Courier New", monospace;
          direction:ltr;
          text-align:left;
        `;

        body.appendChild(ta);

        const footer = document.createElement("div");
        footer.style.cssText = `
          display:flex; gap:10px; justify-content:flex-end;
          padding:12px 14px; border-top:1px solid #e5e7eb;
        `;

        const mkBtn = (txt) => {
          const b = document.createElement("button");
          b.textContent = txt;
          b.style.cssText = `
            border:1px solid #cbd5e1;
            padding:10px 14px;
            border-radius:10px;
            cursor:pointer;
            background:#fff;
            font-weight:900;
          `;
          return b;
        };

        const btnClose = mkBtn("Close");
        const btnCopy = mkBtn("Copy");
        btnCopy.style.border = "none";
        btnCopy.style.background = "#2563eb";
        btnCopy.style.color = "#fff";

        const close = () => overlay.remove();
        btnClose.onclick = close;

        btnCopy.onclick = async () => {
          try {
            await navigator.clipboard.writeText(ta.value || "");
            btnCopy.textContent = "Copied ✅";
            setTimeout(() => (btnCopy.textContent = "Copy"), 900);
          } catch {
            ta.focus();
            ta.select();
            document.execCommand("copy");
            btnCopy.textContent = "Copied ✅";
            setTimeout(() => (btnCopy.textContent = "Copy"), 900);
          }
        };

        footer.appendChild(btnClose);
        footer.appendChild(btnCopy);

        box.appendChild(header);
        box.appendChild(body);
        box.appendChild(footer);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        ta.focus();
        ta.select();
      };

      // ---------- core ----------
      try {
        const Xrm = window.Xrm;
        const page = Xrm?.Page;

        if (!Xrm || !page?.ui?.controls?.forEach) {
          openModal("Xrm not found. Open a record form first.", "Dirty Fields");
          return;
        }

        window.__d365_dirty_highlight_on = window.__d365_dirty_highlight_on || false;

        const dirty = [];
        const els = [];

        page.ui.controls.forEach((c) => {
          try {
            if (!c?.getName || !c?.getAttribute) return;

            const attr = c.getAttribute();
            if (!attr?.getIsDirty || !attr.getIsDirty()) return;

            const logical = c.getName();
            const label = (c.getLabel && c.getLabel()) || logical;

            dirty.push({ label, logical });

            const el =
              document.getElementById(logical) ||
              document.querySelector(`[data-id="${logical}"]`) ||
              document.querySelector(`[data-id="${logical}.fieldControl"]`);

            if (el) els.push(el);
          } catch {}
        });

        // toggle off
        if (window.__d365_dirty_highlight_on) {
          (window.__d365_dirty_highlight_els || []).forEach((el) => {
            try { el.style.outline = ""; el.style.background = ""; } catch {}
          });
          window.__d365_dirty_highlight_els = [];
          window.__d365_dirty_highlight_on = false;
        } else {
          // toggle on
          els.forEach((el) => {
            try {
              el.style.outline = "3px solid #facc15";
              el.style.background = "rgba(250, 204, 21, .12)";
            } catch {}
          });
          window.__d365_dirty_highlight_els = els;
          window.__d365_dirty_highlight_on = true;
        }

        // ---------- build aligned text ----------
        const mode = window.__d365_dirty_highlight_on ? "HIGHLIGHTED" : "CLEARED";

        const maxLabel = Math.max(5, ...dirty.map(d => (d.label || "").length));
        const lines = [];
        lines.push(`Dirty Fields (${mode})`);
        lines.push(`Count: ${dirty.length}`);
        lines.push("");

        if (!dirty.length) {
          lines.push("No dirty fields found.");
        } else {
          lines.push(`${"Label".padEnd(maxLabel)} | Logical Name`);
          lines.push(`${"-".repeat(maxLabel)}-+------------`);
          dirty.forEach(d => {
            lines.push(`${(d.label || "").padEnd(maxLabel)} | ${d.logical}`);
          });
          lines.push("");
          lines.push("Tip: Click the button again to clear highlights.");
        }

        openModal(lines.join("\n"), "Dirty Fields");
      } catch (e) {
        openModal(String(e), "Dirty Fields (Error)");
      }
    }
  });
});



document.getElementById("openDefaultView")?.addEventListener("click", async () => {

  const [tab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });

  if (!tab?.id || !tab.url) {
    alert("No active Dynamics tab found.");
    return;
  }

  // OPEN POPUP IMMEDIATELY
  const popup = showEntityPickerPopup([], tab.url, true);

  try {

    const frameResults = await chrome.scripting.executeScript({
      target: {
        tabId: tab.id,
        allFrames: true
      },
      world: "MAIN",

      func: async () => {

        try {

          const Xrm = window.Xrm;

          if (!Xrm?.Utility) {

            return {
              ok: false,
              hasXrm: false
            };
          }

          const clientUrl =
            Xrm.Utility.getGlobalContext().getClientUrl();

          const url =
            clientUrl +
            "/api/data/v9.2/EntityDefinitions" +
            "?$select=LogicalName,DisplayName,ObjectTypeCode,IsCustomEntity,IsActivity,IsIntersect" +
            "&$filter=IsIntersect eq false";

          const cacheKey = "__d365_entity_defs_v2:" + clientUrl;
          const ttlMs = 10 * 60 * 1000;
          let entities = null;

          try {
            const cached = JSON.parse(sessionStorage.getItem(cacheKey) || "null");
            if (cached && Date.now() - cached.t < ttlMs && Array.isArray(cached.entities)) {
              entities = cached.entities;
            }
          } catch {}

          if (!entities) {
            const response = await fetch(url, {
              method: "GET",
              headers: {
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
                "Prefer": "odata.maxpagesize=5000"
              }
            });

            if (!response.ok) {

              const text = await response.text();

              throw new Error(text);
            }

            const json = await response.json();

            entities = json.value
              .map(e => ({
                logicalName: e.LogicalName,
                displayName:
                  e.DisplayName?.UserLocalizedLabel?.Label ||
                  e.LogicalName,
                objectTypeCode: e.ObjectTypeCode,
                isCustom: e.IsCustomEntity,
                isActivity: e.IsActivity
              }))
              .filter(e => e.logicalName)
              .sort((a, b) =>
                a.displayName.localeCompare(b.displayName)
              );

            try { sessionStorage.setItem(cacheKey, JSON.stringify({ t: Date.now(), entities })); } catch {}
          }

          return {
            ok: true,
            hasXrm: true,
            entities
          };

        } catch (e) {

          return {
            ok: false,
            hasXrm: true,
            message: e.message || String(e)
          };
        }
      }
    });

    const crmFrameResult =
      frameResults.find(r => r.result?.hasXrm);

    if (!crmFrameResult) {

      popup.setError(
        "Could not find Xrm. Open this on a Dynamics 365 page."
      );

      return;
    }

    const data = crmFrameResult.result;

    if (!data.ok) {

      popup.setError(
        data.message || "Failed to retrieve entities."
      );

      return;
    }

    popup.setEntities(data.entities || []);

  } catch (e) {

    popup.setError(
      e.message || String(e)
    );
  }
});


function showEntityPickerPopup(
  initialEntities,
  currentUrl,
  loading = false
) {

  document.getElementById("entityPickerOverlay")?.remove();
  document.getElementById("entityPickerStyle")?.remove();

  const overlay = document.createElement("div");

  overlay.id = "entityPickerOverlay";

  overlay.innerHTML = `
    <div class="entity-picker-modal">

      <div class="entity-picker-header">
        <div>
          <h2>Open Entity View</h2>
          <p>Search entity from current CRM</p>
        </div>

        <button
          id="entityPickerClose"
          class="entity-picker-close"
        >
          ×
        </button>
      </div>

      <input
        id="entityPickerSearch"
        class="entity-picker-search"
        type="text"
        placeholder="Search display name or logical name..."
      />

      <div
        id="entityPickerLoading"
        class="entity-picker-loading"
        style="${loading ? "" : "display:none"}"
      >
        Loading entities...
      </div>

      <div
        id="entityPickerError"
        class="entity-picker-error"
        style="display:none"
      ></div>

      <div
        id="entityPickerList"
        class="entity-picker-list"
      ></div>

      <div class="entity-picker-footer">

        <input
          id="entityPickerManual"
          class="entity-picker-manual"
          type="text"
          placeholder="Manual logical name..."
        />

        <button
          id="entityPickerOpenManual"
          class="entity-picker-open"
        >
          Open
        </button>

      </div>

    </div>
  `;

  document.body.appendChild(overlay);

  const style = document.createElement("style");

  style.id = "entityPickerStyle";

  style.textContent = `
    #entityPickerOverlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.65);
      z-index: 999999;
      display: flex;
      align-items: center;
      justify-content: center;
      direction: rtl;
      font-family: Arial, sans-serif;
    }

    .entity-picker-modal {
      width: 540px;
      max-width: calc(100vw - 24px);
      max-height: calc(100vh - 24px);
      background: #1f1f1f;
      color: white;
      border-radius: 14px;
      border: 1px solid #444;
      overflow: hidden;
      box-shadow: 0 12px 40px rgba(0,0,0,0.45);
      display: flex;
      flex-direction: column;
    }

    .entity-picker-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 16px;
      border-bottom: 1px solid #333;
    }

    .entity-picker-header h2 {
      margin: 0;
      font-size: 18px;
    }

    .entity-picker-header p {
      margin: 4px 0 0;
      color: #aaa;
      font-size: 12px;
    }

    .entity-picker-close {
      background: transparent;
      border: 0;
      color: white;
      font-size: 28px;
      cursor: pointer;
    }

    .entity-picker-search,
    .entity-picker-manual {
      width: calc(100% - 32px);
      margin: 14px 16px;
      padding: 11px;
      border-radius: 8px;
      border: 1px solid #555;
      background: #121212;
      color: white;
      box-sizing: border-box;
      outline: none;
    }

    .entity-picker-loading {
      padding: 24px;
      text-align: center;
      color: #8cc8ff;
    }

    .entity-picker-error {
      padding: 20px;
      text-align: center;
      color: #ff8c8c;
    }

    .entity-picker-list {
      flex: 1;
      overflow-y: auto;
      padding: 0 12px 12px;
      min-height: 120px;
      max-height: 420px;
    }

    .entity-picker-row {
      width: 100%;
      text-align: start;
      padding: 11px 12px;
      margin-bottom: 6px;
      border-radius: 8px;
      border: 1px solid #333;
      background: #2a2a2a;
      color: white;
      cursor: pointer;
      display: flex;
      flex-direction: column;
      gap: 3px;
    }

    .entity-picker-row:hover {
      background: #333;
      border-color: #6aa9ff;
    }

    .entity-picker-name {
      font-weight: bold;
      font-size: 14px;
    }

    .entity-picker-logical {
      color: #aaa;
      font-size: 12px;
      direction: ltr;
      text-align: left;
    }

    .entity-picker-tags {
      color: #8cc8ff;
      font-size: 11px;
    }

    .entity-picker-empty {
      color: #aaa;
      text-align: center;
      padding: 28px;
    }

    .entity-picker-footer {
      border-top: 1px solid #333;
      display: flex;
      align-items: center;
      gap: 8px;
      padding-bottom: 14px;
    }

    .entity-picker-footer .entity-picker-manual {
      flex: 1;
      margin-left: 0;
    }

    .entity-picker-open {
      margin-left: 16px;
      padding: 11px 18px;
      border-radius: 8px;
      border: 0;
      background: #6aa9ff;
      color: #000;
      font-weight: bold;
      cursor: pointer;
    }
  `;

  document.head.appendChild(style);

  const searchInput =
    document.getElementById("entityPickerSearch");

  const list =
    document.getElementById("entityPickerList");

  const loadingEl =
    document.getElementById("entityPickerLoading");

  const errorEl =
    document.getElementById("entityPickerError");

  const closeBtn =
    document.getElementById("entityPickerClose");

  const manualInput =
    document.getElementById("entityPickerManual");

  const manualBtn =
    document.getElementById("entityPickerOpenManual");

  let entities = initialEntities || [];

  function closePopup() {

    document.getElementById("entityPickerOverlay")?.remove();

    document.getElementById("entityPickerStyle")?.remove();
  }

  function openEntity(logicalName) {

    const cleanEntity =
      logicalName?.trim().toLowerCase();

    if (!cleanEntity) {
      alert("Enter entity logical name.");
      return;
    }

    const u = new URL(currentUrl);

    const orgUrl =
      `${u.protocol}//${u.host}`;

    const appid =
      u.searchParams.get("appid");

    const url =
      `${orgUrl}/main.aspx?` +
      (appid
        ? `appid=${encodeURIComponent(appid)}&`
        : "") +
      `pagetype=entitylist&etn=${encodeURIComponent(cleanEntity)}`;

    chrome.tabs.create({ url });

    closePopup();
  }

  function escapeHtml(value) {

    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function render(filterText = "") {

    const q =
      filterText.trim().toLowerCase();

    const filtered = entities.filter(e => {

      const text =
        `${e.displayName} ${e.logicalName}`.toLowerCase();

      return !q || text.includes(q);
    });

    list.innerHTML = "";

    if (!filtered.length) {

      list.innerHTML =
        `<div class="entity-picker-empty">
          No entities found
        </div>`;

      return;
    }

    filtered.slice(0, 500).forEach(e => {

      const btn =
        document.createElement("button");

      btn.className =
        "entity-picker-row";

      const tags = [];

      if (e.isCustom) tags.push("Custom");

      if (e.isActivity) tags.push("Activity");

      if (e.objectTypeCode)
        tags.push(`OTC ${e.objectTypeCode}`);

      btn.innerHTML = `
        <span class="entity-picker-name">
          ${escapeHtml(e.displayName)}
        </span>

        <span class="entity-picker-logical">
          ${escapeHtml(e.logicalName)}
        </span>

        <span class="entity-picker-tags">
          ${escapeHtml(tags.join(" • "))}
        </span>
      `;

      btn.addEventListener("click", () =>
        openEntity(e.logicalName)
      );

      list.appendChild(btn);
    });
  }

  function setEntities(newEntities) {

    entities = newEntities || [];

    loadingEl.style.display = "none";

    errorEl.style.display = "none";

    render(searchInput.value);
  }

  function setError(message) {

    loadingEl.style.display = "none";

    errorEl.style.display = "block";

    errorEl.textContent = message;
  }

  closeBtn.addEventListener("click", closePopup);

  overlay.addEventListener("click", e => {

    if (e.target === overlay) {
      closePopup();
    }
  });

  let renderTimer = 0;
  searchInput.addEventListener("input", () => {
    clearTimeout(renderTimer);
    renderTimer = setTimeout(() => render(searchInput.value), 80);
  });

  manualBtn.addEventListener("click", () => {

    openEntity(manualInput.value);
  });

  manualInput.addEventListener("keydown", e => {

    if (e.key === "Enter") {
      openEntity(manualInput.value);
    }
  });

  render();

  searchInput.focus();

  return {
    setEntities,
    setError
  };
}




document.getElementById("getRolePermissions").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position:fixed; inset:0; background:rgba(15,23,42,.45);
        z-index:2147483647; display:flex; align-items:center; justify-content:center; padding:16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width:min(1450px,98vw); height:min(880px,94vh); background:#fff;
        border-radius:16px; box-shadow:0 24px 70px rgba(0,0,0,.35);
        overflow:hidden; font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
        display:flex; flex-direction:column; direction:rtl;
      `;

      const header = document.createElement("div");
      header.style.cssText = `
        padding:16px 20px; font-size:18px; font-weight:900;
        border-bottom:1px solid #e5e7eb; color:#111827;
      `;
      header.textContent = "Role Permissions By Entity";

      const body = document.createElement("div");
      body.style.cssText = `
        padding:16px; display:grid; gap:12px; flex:1; min-height:0;
        grid-template-rows:auto auto auto auto 1fr auto;
      `;

      const inputStyle = `
        width:100%; border:1px solid #cbd5e1; border-radius:12px;
        padding:12px 14px; font-size:13px; box-sizing:border-box; outline:none;
        background:#fff; color:#111827;
      `;

      const mkRow = (label, el) => {
        const wrap = document.createElement("div");
        wrap.style.cssText = "display:grid; gap:6px; position:relative;";
        const l = document.createElement("div");
        l.textContent = label;
        l.style.cssText = "font-size:12px;font-weight:800;color:#111827;";
        wrap.appendChild(l);
        wrap.appendChild(el);
        return wrap;
      };

      const status = document.createElement("div");
      status.style.cssText = "font-size:12px;color:#374151;";

      const summary = document.createElement("div");
      summary.style.cssText = `
        font-size:12px; color:#111827; background:#f8fafc;
        border:1px solid #e5e7eb; border-radius:12px; padding:10px;
      `;
      summary.textContent = "No data yet.";

      const tableWrap = document.createElement("div");
      tableWrap.style.cssText = `
        border:1px solid #cbd5e1; border-radius:12px; overflow:auto;
        min-height:260px; background:#fff; direction:ltr;
      `;

      const table = document.createElement("table");
      table.style.cssText = `
        width:100%; border-collapse:collapse; font-size:12px; direction:ltr; text-align:left;
      `;
      tableWrap.appendChild(table);

      const rawTa = document.createElement("textarea");
      rawTa.readOnly = true;
      rawTa.placeholder = "Raw JSON...";
      rawTa.style.cssText = `
        width:100%; height:140px; resize:vertical; border:1px solid #cbd5e1;
        border-radius:12px; padding:10px; font-size:12px; box-sizing:border-box;
        font-family:Consolas,monospace; direction:ltr; text-align:left;
      `;

      const roleInput = document.createElement("input");
      roleInput.placeholder = "Search role / business unit...";
      roleInput.style.cssText = inputStyle;

      const roleDrop = document.createElement("div");
      roleDrop.style.cssText = `
        display:none; position:absolute; top:74px; left:0; right:0;
        max-height:330px; overflow:auto; background:#fff; border:1px solid #cbd5e1;
        border-radius:12px; box-shadow:0 16px 40px rgba(0,0,0,.18);
        z-index:2147483647; direction:ltr; text-align:left;
      `;

      const entityInput = document.createElement("input");
      entityInput.placeholder = "Search entity logical name...";
      entityInput.style.cssText = inputStyle;

      const entityDrop = document.createElement("div");
      entityDrop.style.cssText = roleDrop.style.cssText;

      const roleRow = mkRow("Role Name / Business Unit", roleInput);
      roleRow.appendChild(roleDrop);

      const entityRow = mkRow("Entity Logical Name", entityInput);
      entityRow.appendChild(entityDrop);

      body.appendChild(roleRow);
      body.appendChild(entityRow);
      body.appendChild(status);
      body.appendChild(summary);
      body.appendChild(tableWrap);
      body.appendChild(rawTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display:flex; gap:10px; justify-content:flex-end;
        padding:14px 16px; border-top:1px solid #e5e7eb;
      `;

      const btn = (text, bg, color = "#111827") => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border:none; padding:10px 14px; border-radius:10px; cursor:pointer;
          font-weight:800; background:${bg}; color:${color};
        `;
        return b;
      };

      const btnClose = btn("Close", "#e5e7eb");
      const btnCopyJson = btn("Copy JSON", "#2563eb", "#fff");
      const btnCopyCsv = btn("Copy CSV", "#059669", "#fff");
      const btnRun = btn("Run", "#111827", "#fff");

      footer.append(btnClose, btnCopyJson, btnCopyCsv, btnRun);
      box.append(header, body, footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      let rolesOptions = [];
      let entitiesOptions = [];
      let selectedRole = null;
      let lastRows = [];

      const getWebApi = () => window.Xrm?.WebApi?.online || window.Xrm?.WebApi;

      const retrieveAll = async (entityName, query) => {
        const webApi = getWebApi();
        let result = await webApi.retrieveMultipleRecords(entityName, query);
        let rows = [...(result.entities || [])];

        while (result.nextLink) {
          result = await webApi.retrieveMultipleRecords(entityName, result.nextLink);
          rows.push(...(result.entities || []));
        }

        return rows;
      };

      const escapeHtml = s =>
        String(s ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");

      const escapeXml = s =>
        String(s ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&apos;");

      const closeDrops = () => {
        roleDrop.style.display = "none";
        entityDrop.style.display = "none";
      };

      overlay.addEventListener("click", e => {
        if (!roleRow.contains(e.target) && !entityRow.contains(e.target)) closeDrops();
      });

      btnClose.onclick = () => overlay.remove();

      const renderDropdown = (drop, items, onSelect, emptyText) => {
        drop.innerHTML = "";

        if (!items.length) {
          const empty = document.createElement("div");
          empty.textContent = emptyText || "No results";
          empty.style.cssText = "padding:12px;color:#6b7280;font-size:13px;";
          drop.appendChild(empty);
          drop.style.display = "block";
          return;
        }

        items.slice(0, 80).forEach(item => {
          const row = document.createElement("div");
          row.style.cssText = `
            padding:10px 12px; cursor:pointer; border-bottom:1px solid #f1f5f9;
            display:grid; gap:3px; background:#fff;
          `;

          row.onmouseenter = () => row.style.background = "#f8fafc";
          row.onmouseleave = () => row.style.background = "#fff";
          row.onclick = () => onSelect(item);

          const main = document.createElement("div");
          main.innerHTML = escapeHtml(item.main);
          main.style.cssText = "font-size:13px;font-weight:800;color:#111827;";

          row.appendChild(main);

          if (item.sub) {
            const sub = document.createElement("div");
            sub.innerHTML = escapeHtml(item.sub);
            sub.style.cssText = "font-size:12px;color:#64748b;";
            row.appendChild(sub);
          }

          drop.appendChild(row);
        });

        drop.style.display = "block";
      };

      const filterItems = (items, text) => {
        const q = String(text || "").trim().toLowerCase();
        if (!q) return items.slice(0, 80);

        return items.filter(x =>
          `${x.main} ${x.sub || ""} ${x.search || ""}`.toLowerCase().includes(q)
        );
      };

      roleInput.addEventListener("input", () => {
        selectedRole = null;
        renderDropdown(
          roleDrop,
          filterItems(rolesOptions, roleInput.value),
          item => {
            selectedRole = item.raw;
            roleInput.value = `${item.raw.roleName} | ${item.raw.businessUnitName}`;
            roleDrop.style.display = "none";
          },
          "No roles found"
        );
      });

      roleInput.addEventListener("focus", () => {
        renderDropdown(
          roleDrop,
          filterItems(rolesOptions, roleInput.value),
          item => {
            selectedRole = item.raw;
            roleInput.value = `${item.raw.roleName} | ${item.raw.businessUnitName}`;
            roleDrop.style.display = "none";
          },
          "No roles found"
        );
      });

      entityInput.addEventListener("input", () => {
        renderDropdown(
          entityDrop,
          filterItems(entitiesOptions, entityInput.value),
          item => {
            entityInput.value = item.raw.logicalName;
            entityDrop.style.display = "none";
          },
          "No entities found"
        );
      });

      entityInput.addEventListener("focus", () => {
        renderDropdown(
          entityDrop,
          filterItems(entitiesOptions, entityInput.value),
          item => {
            entityInput.value = item.raw.logicalName;
            entityDrop.style.display = "none";
          },
          "No entities found"
        );
      });

      const renderTable = rows => {
        table.innerHTML = "";

        const cols = [
          { key: "roleName", label: "Role" },
          { key: "businessUnitName", label: "Business Unit" },
          { key: "entityName", label: "Entity" },
          { key: "permission", label: "Permission" },
          { key: "depth", label: "Depth" },
          { key: "privilegeName", label: "Privilege" }
        ];

        const thead = document.createElement("thead");
        const trh = document.createElement("tr");

        cols.forEach(c => {
          const th = document.createElement("th");
          th.textContent = c.label;
          th.style.cssText = `
            position:sticky; top:0; background:#0f172a; color:#fff;
            padding:10px; text-align:left; white-space:nowrap;
          `;
          trh.appendChild(th);
        });

        thead.appendChild(trh);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");

        if (!rows.length) {
          const tr = document.createElement("tr");
          const td = document.createElement("td");
          td.colSpan = cols.length;
          td.textContent = "No permissions found.";
          td.style.cssText = "padding:14px;text-align:center;color:#6b7280;";
          tr.appendChild(td);
          tbody.appendChild(tr);
          table.appendChild(tbody);
          return;
        }

        rows.forEach((r, idx) => {
          const tr = document.createElement("tr");
          tr.style.background = idx % 2 ? "#f8fafc" : "#fff";

          cols.forEach(c => {
            const td = document.createElement("td");
            td.textContent = r[c.key] || "";
            td.style.cssText = `
              padding:8px 10px; border-bottom:1px solid #e5e7eb;
              white-space:nowrap; max-width:420px; overflow:hidden; text-overflow:ellipsis;
            `;
            tr.appendChild(td);
          });

          tbody.appendChild(tr);
        });

        table.appendChild(tbody);
      };

      const loadOptions = async () => {
  const Xrm = window.Xrm;
  const webApi = getWebApi();

  if (!Xrm || !webApi?.retrieveMultipleRecords) {
    status.textContent = "❌ Xrm.WebApi not found. Open this on a D365 page.";
    return;
  }

  try {
    status.textContent = "⏳ Loading parent BU roles and entities...";

    const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
    const cachePrefix = "__d365_role_permissions_options_v2:" + clientUrl + ":";
    const ttlMs = 10 * 60 * 1000;
    const readCache = (k) => {
      try {
        const hit = JSON.parse(sessionStorage.getItem(cachePrefix + k) || "null");
        return hit && Date.now() - hit.t < ttlMs ? hit.v : null;
      } catch { return null; }
    };
    const writeCache = (k, v) => {
      try { sessionStorage.setItem(cachePrefix + k, JSON.stringify({ t: Date.now(), v })); } catch {}
      return v;
    };

    const buRows = readCache("businessunits") || writeCache("businessunits", await retrieveAll(
      "businessunit",
      "?$select=businessunitid,name,_parentbusinessunitid_value"
    ));

    const parentBu = buRows.find(b => !b._parentbusinessunitid_value);

    if (!parentBu) {
      status.textContent = "❌ Parent BU not found.";
      return;
    }

    const parentBuId = parentBu.businessunitid.replace(/[{}]/g, "");

    const roles = readCache("roles:" + parentBuId) || writeCache("roles:" + parentBuId, await retrieveAll(
      "role",
      `?$select=roleid,name,_businessunitid_value
       &$filter=_businessunitid_value eq ${parentBuId}
       &$orderby=name asc`
    ));

    rolesOptions = roles
      .filter(r => r.name && r.roleid)
      .map(r => {
        const buName =
          r["_businessunitid_value@OData.Community.Display.V1.FormattedValue"] ||
          parentBu.name ||
          "";

        const raw = {
          roleId: r.roleid.replace(/[{}]/g, ""),
          roleName: r.name,
          businessUnitName: buName
        };

        return {
          main: r.name,
          sub: buName,
          search: `${r.name} ${buName}`,
          raw
        };
      });

    const cachedEntities = readCache("entities");
    const data = cachedEntities || await (async () => {
      const response = await fetch(
        `${clientUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName`,
        {
          method: "GET",
          headers: {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=*,odata.maxpagesize=5000"
          }
        }
      );
      const json = await response.json();
      return writeCache("entities", json);
    })();

    entitiesOptions = (data.value || [])
      .filter(e => e.LogicalName)
      .sort((a, b) => a.LogicalName.localeCompare(b.LogicalName))
      .map(e => {
        const display =
          e.DisplayName?.UserLocalizedLabel?.Label ||
          e.DisplayName?.LocalizedLabels?.[0]?.Label ||
          "";

        return {
          main: e.LogicalName,
          sub: display,
          search: `${e.LogicalName} ${display}`,
          raw: { logicalName: e.LogicalName }
        };
      });

    status.textContent =
      `✅ Loaded ${rolesOptions.length} roles from parent BU: ${parentBu.name}, and ${entitiesOptions.length} entities`;

  } catch (err) {
    console.error(err);
    status.textContent = "❌ Failed loading options: " + (err?.message || err);
  }
};

      const toCsv = rows => {
        const cols = ["roleName", "businessUnitName", "entityName", "permission", "depth", "privilegeName"];
        const esc = v => {
          const s = String(v ?? "");
          const out = s.replaceAll('"', '""');
          return /[",\n]/.test(s) ? `"${out}"` : out;
        };
        return [cols.join(","), ...rows.map(r => cols.map(c => esc(r[c])).join(","))].join("\n");
      };

      const copyText = async (text, btnRef, originalText) => {
        try {
          await navigator.clipboard.writeText(text || "");
        } catch {
          rawTa.value = text || "";
          rawTa.focus();
          rawTa.select();
          document.execCommand("copy");
        }

        btnRef.textContent = "Copied ✅";
        setTimeout(() => btnRef.textContent = originalText, 900);
      };

      const runGet = async () => {
        closeDrops();
        status.textContent = "";
        summary.textContent = "Loading...";
        rawTa.value = "";
        lastRows = [];
        renderTable([]);

        if (!selectedRole) {
          const exact = rolesOptions.find(x =>
            `${x.raw.roleName} | ${x.raw.businessUnitName}` === roleInput.value
          );
          selectedRole = exact?.raw || null;
        }

        if (!selectedRole) {
          status.textContent = "❌ Select exact role from the dropdown.";
          summary.textContent = "Role missing.";
          return;
        }

        const entityLogicalName = entityInput.value.trim().toLowerCase();

        if (!entityLogicalName) {
          status.textContent = "❌ Entity logical name is required.";
          summary.textContent = "Entity missing.";
          return;
        }

        try {
          status.textContent = "⏳ Loading permissions...";

          const depthMap = {
            0: "None",
            1: "User",
            2: "Business Unit",
            4: "Parent:Child Business Unit",
            8: "Organization"
          };

          const rights = ["Create", "Read", "Write", "Delete", "Append", "AppendTo", "Assign", "Share"];

          const fetchXml = `
<fetch distinct="true">
  <entity name="role">
    <attribute name="roleid" />
    <attribute name="name" />
    <filter>
      <condition attribute="roleid" operator="eq" value="${escapeXml(selectedRole.roleId)}" />
    </filter>
    <link-entity name="roleprivileges" from="roleid" to="roleid" intersect="true" alias="rp">
      <attribute name="privilegedepthmask" />
      <link-entity name="privilege" from="privilegeid" to="privilegeid" alias="p">
        <attribute name="name" />
      </link-entity>
    </link-entity>
  </entity>
</fetch>`;

          const privilegeResult = await getWebApi().retrieveMultipleRecords(
            "role",
            "?fetchXml=" + encodeURIComponent(fetchXml)
          );

          const finalRows = [];

          privilegeResult.entities.forEach(row => {
            const privilegeName = row["p.name"];
            const depth = row["rp.privilegedepthmask"];
            if (!privilegeName) return;

            const matchedRight = rights.find(r =>
              privilegeName.toLowerCase() === `prv${r}${entityLogicalName}`.toLowerCase()
            );

            if (!matchedRight) return;

            finalRows.push({
              roleName: selectedRole.roleName,
              businessUnitName: selectedRole.businessUnitName,
              entityName: entityLogicalName,
              permission: matchedRight,
              depth: depthMap[depth] || String(depth),
              privilegeName
            });
          });

          finalRows.sort((a, b) => a.permission.localeCompare(b.permission));

          lastRows = finalRows;
          rawTa.value = JSON.stringify(finalRows, null, 2);
          renderTable(finalRows);

          summary.innerHTML = `
            <b>Role:</b> ${escapeHtml(selectedRole.roleName)}
            &nbsp; | &nbsp;
            <b>BU:</b> ${escapeHtml(selectedRole.businessUnitName)}
            &nbsp; | &nbsp;
            <b>Entity:</b> ${escapeHtml(entityLogicalName)}
            &nbsp; | &nbsp;
            <b>Permissions:</b> ${finalRows.length}
          `;

          status.textContent = finalRows.length ? `✅ Done (${finalRows.length})` : "⚠️ No permissions found.";
        } catch (err) {
          console.error(err);
          status.textContent = "❌ Failed";
          rawTa.value = err?.message || String(err);
          summary.textContent = "Error";
          renderTable([]);
        }
      };

      btnCopyJson.onclick = () => copyText(rawTa.value, btnCopyJson, "Copy JSON");
      btnCopyCsv.onclick = () => copyText(toCsv(lastRows), btnCopyCsv, "Copy CSV");
      btnRun.onclick = runGet;

      roleInput.addEventListener("keydown", e => {
        if (e.key === "Enter") runGet();
        if (e.key === "Escape") closeDrops();
      });

      entityInput.addEventListener("keydown", e => {
        if (e.key === "Enter") runGet();
        if (e.key === "Escape") closeDrops();
      });

      renderTable([]);
      roleInput.focus();
      loadOptions();
    }
  });
});


document.getElementById("getSystemRolesByPrivilege").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(1200px, 96vw);
        height: min(780px, 92vh);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
        display: flex;
        flex-direction: column;
      `;

      const header = document.createElement("div");
      header.style.cssText = `
        padding: 12px 14px;
        font-weight: 800;
        border-bottom: 1px solid #e5e7eb;
      `;
      header.textContent = "System Roles By Privilege";

      const body = document.createElement("div");
      body.style.cssText = `
        padding: 12px 14px;
        display: grid;
        gap: 10px;
        min-height: 0;
        flex: 1;
      `;

      const inputStyle = `
        width: 100%;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 13px;
        box-sizing: border-box;
      `;

      const mkRow = (label, el) => {
        const wrap = document.createElement("div");
        wrap.style.cssText = `display:grid; gap:6px;`;
        const l = document.createElement("div");
        l.textContent = label;
        l.style.cssText = `font-size: 12px; font-weight: 700; color: #111827;`;
        wrap.appendChild(l);
        wrap.appendChild(el);
        return wrap;
      };

      const entityInput = document.createElement("input");
      entityInput.placeholder = "Entity logical name (e.g. contact)";
      entityInput.style.cssText = inputStyle;

      const actionSelect = document.createElement("select");
      actionSelect.style.cssText = inputStyle;
      [
        { value: "", text: "Select privilege action..." },
        { value: "Create", text: "Create" },
        { value: "Read", text: "Read" },
        { value: "Write", text: "Write" },
        { value: "Delete", text: "Delete" },
        { value: "Append", text: "Append" },
        { value: "AppendTo", text: "Append To" },
        { value: "Assign", text: "Assign" },
        { value: "Share", text: "Share" }
      ].forEach(o => {
        const opt = document.createElement("option");
        opt.value = o.value;
        opt.textContent = o.text;
        actionSelect.appendChild(opt);
      });

      const depthSelect = document.createElement("select");
      depthSelect.style.cssText = inputStyle;
      [
        { value: "", text: "Select depth..." },
        { value: "1", text: "User / Basic (1)" },
        { value: "2", text: "Business Unit / Local (2)" },
        { value: "4", text: "Parent:Child BU / Deep (4)" },
        { value: "8", text: "Organization / Global (8)" }
      ].forEach(o => {
        const opt = document.createElement("option");
        opt.value = o.value;
        opt.textContent = o.text;
        depthSelect.appendChild(opt);
      });

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const summary = document.createElement("div");
      summary.style.cssText = `
        font-size: 12px;
        color: #111827;
        background: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 10px;
      `;
      summary.textContent = "No data yet.";

      const tableWrap = document.createElement("div");
      tableWrap.style.cssText = `
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        overflow: auto;
        min-height: 220px;
        height: 100%;
        background: #fff;
      `;

      const table = document.createElement("table");
      table.style.cssText = `
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        direction: ltr;
        text-align: left;
      `;
      tableWrap.appendChild(table);

      const rawTa = document.createElement("textarea");
      rawTa.readOnly = true;
      rawTa.placeholder = "Raw JSON will appear here…";
      rawTa.style.cssText = `
        width: 100%;
        height: 150px;
        resize: vertical;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        font-size: 12px;
        line-height: 1.4;
        white-space: pre;
        box-sizing: border-box;
        font-family: Consolas, Monaco, "Courier New", monospace;
        direction: ltr;
        text-align: left;
      `;

      body.appendChild(mkRow("Entity Logical Name", entityInput));
      body.appendChild(mkRow("Privilege Action", actionSelect));
      body.appendChild(mkRow("Required Depth", depthSelect));
      body.appendChild(status);
      body.appendChild(summary);
      body.appendChild(tableWrap);
      body.appendChild(rawTa);

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex; gap: 10px; justify-content: flex-end;
        padding: 12px 14px; border-top: 1px solid #e5e7eb;
      `;

      const btn = (text) => {
        const b = document.createElement("button");
        b.textContent = text;
        b.style.cssText = `
          border: 1px solid #cbd5e1;
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: #fff;
          font-weight: 800;
        `;
        return b;
      };

      const btnClose = btn("Close");

      const btnCopyJson = btn("Copy JSON");
      btnCopyJson.style.border = "none";
      btnCopyJson.style.background = "#2563eb";
      btnCopyJson.style.color = "#fff";

      const btnCopyCsv = btn("Copy CSV");
      btnCopyCsv.style.border = "none";
      btnCopyCsv.style.background = "#059669";
      btnCopyCsv.style.color = "#fff";

      const btnRun = btn("Run");
      btnRun.style.border = "none";
      btnRun.style.background = "#111827";
      btnRun.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;

      let lastRows = [];

      const escapeHtml = (s) =>
        String(s ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");

      const copyText = async (text, btnRef, originalLabel) => {
        try {
          await navigator.clipboard.writeText(text || "");
        } catch {
          rawTa.value = text || "";
          rawTa.focus();
          rawTa.select();
          document.execCommand("copy");
        }
        btnRef.textContent = "Copied ✅";
        setTimeout(() => (btnRef.textContent = originalLabel), 900);
      };

      const toCsv = (rows) => {
        if (!rows.length) return "";
        const cols = ["Distinct Role Name", "Entity", "Action", "Depth"];
        const esc = (v) => {
          const s = String(v ?? "");
          const need = /[",\n]/.test(s);
          const out = s.replaceAll('"', '""');
          return need ? `"${out}"` : out;
        };
        return [
          cols.join(","),
          ...rows.map(r => cols.map(c => esc(r[c])).join(","))
        ].join("\n");
      };

      const renderTable = (rows) => {
        table.innerHTML = "";

        const cols = [
          { key: "Distinct Role Name", label: "Distinct Role Name" },
          { key: "Entity", label: "Entity" },
          { key: "Action", label: "Action" },
          { key: "Depth", label: "Depth" }
        ];

        const thead = document.createElement("thead");
        const trh = document.createElement("tr");

        cols.forEach(c => {
          const th = document.createElement("th");
          th.innerHTML = escapeHtml(c.label);
          th.style.cssText = `
            position: sticky;
            top: 0;
            background: #0f172a;
            color: #fff;
            padding: 10px 8px;
            border-bottom: 1px solid rgba(255,255,255,.15);
            white-space: nowrap;
            z-index: 1;
          `;
          trh.appendChild(th);
        });

        thead.appendChild(trh);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");

        if (!rows.length) {
          const tr = document.createElement("tr");
          const td = document.createElement("td");
          td.colSpan = cols.length;
          td.textContent = "No roles found.";
          td.style.cssText = `
            padding: 14px 10px;
            color: #6b7280;
            text-align: center;
          `;
          tr.appendChild(td);
          tbody.appendChild(tr);
          table.appendChild(tbody);
          return;
        }

        rows.forEach((row, idx) => {
          const tr = document.createElement("tr");
          tr.style.background = idx % 2 === 0 ? "#ffffff" : "#f8fafc";

          cols.forEach(c => {
            const td = document.createElement("td");
            td.innerHTML = escapeHtml(row[c.key] ?? "");
            td.style.cssText = `
              padding: 8px 8px;
              border-bottom: 1px solid #e5e7eb;
              white-space: nowrap;
              max-width: 420px;
              overflow: hidden;
              text-overflow: ellipsis;
            `;
            tr.appendChild(td);
          });

          tbody.appendChild(tr);
        });

        table.appendChild(tbody);
      };

      btnCopyJson.onclick = async () => {
        await copyText(rawTa.value || "", btnCopyJson, "Copy JSON");
      };

      btnCopyCsv.onclick = async () => {
        await copyText(toCsv(lastRows), btnCopyCsv, "Copy CSV");
      };

      const runGet = async () => {
        status.textContent = "";
        summary.textContent = "Loading…";
        rawTa.value = "";
        renderTable([]);
        lastRows = [];

        const entityName = (entityInput.value || "").trim().toLowerCase();
        const privilegeAction = (actionSelect.value || "").trim();
        const requiredDepth = (depthSelect.value || "").trim();

        if (!entityName || !privilegeAction || !requiredDepth) {
          status.textContent = "❌ Entity, action and depth are required.";
          summary.textContent = "Missing input.";
          return;
        }

        const Xrm = window.Xrm;
        const context = Xrm?.Utility?.getGlobalContext?.();
        const clientUrl = context?.getClientUrl?.();

        if (!Xrm || !clientUrl) {
          status.textContent = "❌ D365 context not available. Open a D365 page.";
          summary.textContent = "D365 context not found.";
          return;
        }

        status.textContent = "⏳ Loading…";

        try {
          const normalizeGuid = (id) => (id || "").replace(/[{}]/g, "").toLowerCase();
          const capitalize = (s) => s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : "";
          const getDepthValue = (depth) => {
            const d = String(depth || "").trim().toLowerCase();
            const map = { "basic": 1, "user": 1, "local": 2, "bu": 2, "deep": 4, "global": 8, "org": 8 };
            return isNaN(d) ? (map[d] || 0) : Number(d);
          };

          const wantedDepthValue = getDepthValue(requiredDepth);
          const actionName = capitalize(privilegeAction);

          const variations = [
            `prv${actionName}${entityName}`,
            `prv${actionName}${entityName.toLowerCase()}`,
            `prv${actionName}${capitalize(entityName)}`
          ];

          const privFilter = variations.map(v => `name eq '${v}'`).join(" or ");
          const privUrl = `${clientUrl}/api/data/v9.2/privileges?$select=privilegeid,name&$filter=${privFilter}`;

          const privRes = await fetch(privUrl, {
            method: "GET",
            headers: {
              "Accept": "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0"
            },
            credentials: "same-origin"
          });

          if (!privRes.ok) {
            throw new Error(`Privilege lookup failed (${privRes.status})`);
          }

          const privData = await privRes.json();

          if (!privData.value?.length) {
            status.textContent = "⚠️ Privilege not found.";
            summary.textContent = `No privilege found for ${actionName} on ${entityName}`;
            rawTa.value = JSON.stringify([], null, 2);
            renderTable([]);
            return;
          }

          const targetPrivId = normalizeGuid(privData.value[0].privilegeid);

          // Performance: the previous implementation loaded every role and then called
          // RetrieveRolePrivilegesRole once per role. In environments with many roles this
          // can mean hundreds/thousands of requests and popup hangs. This FetchXML does the
          // join server-side and returns only roles that have the selected privilege/depth.
          const fetchXml = `
<fetch distinct="true" mapping="logical">
  <entity name="role">
    <attribute name="name" />
    <order attribute="name" />
    <link-entity name="roleprivileges" from="roleid" to="roleid" intersect="true" alias="rp">
      <filter>
        <condition attribute="privilegeid" operator="eq" value="${targetPrivId}" />
        <condition attribute="privilegedepthmask" operator="eq" value="${wantedDepthValue}" />
      </filter>
    </link-entity>
  </entity>
</fetch>`;

          const roleUrl = `${clientUrl}/api/data/v9.2/roles?fetchXml=${encodeURIComponent(fetchXml)}`;
          const roleRes = await fetch(roleUrl, {
            method: "GET",
            headers: {
              "Accept": "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0",
              "Prefer": "odata.maxpagesize=5000"
            },
            credentials: "same-origin"
          });

          if (!roleRes.ok) {
            throw new Error(`Role lookup failed (${roleRes.status})`);
          }

          const roleData = await roleRes.json();
          const uniqueRoleNames = [...new Set((roleData.value || []).map(r => r.name).filter(Boolean))].sort();

          const depthLabels = {
            "1": "USER",
            "2": "BU",
            "4": "DEEP",
            "8": "ORG"
          };

          const tableData = uniqueRoleNames.map(name => ({
            "Distinct Role Name": name,
            "Entity": entityName,
            "Action": privilegeAction,
            "Depth": depthLabels[String(requiredDepth)] || String(requiredDepth).toUpperCase()
          }));

          lastRows = tableData;
          rawTa.value = JSON.stringify(tableData, null, 2);
          renderTable(tableData);

          summary.innerHTML = `
            <b>Entity:</b> ${escapeHtml(entityName)}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Action:</b> ${escapeHtml(privilegeAction)}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Depth:</b> ${escapeHtml(depthLabels[String(requiredDepth)] || String(requiredDepth).toUpperCase())}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Roles:</b> ${tableData.length}
          `;

          status.textContent = tableData.length
            ? `✅ Done (${tableData.length} roles).`
            : "⚠️ No roles found with that exact criteria.";

          if (tableData.length) {
            console.table(tableData);
          } else {
            console.warn("No roles found with that exact criteria.");
          }
        } catch (err) {
          status.textContent = "❌ Failed.";
          summary.textContent = "Error";
          rawTa.value =
            "ERROR:\n" +
            (err?.message || err?.toString?.() || "Unknown error");
          renderTable([]);
        }
      };

      btnRun.onclick = runGet;

      entityInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runGet();
      });
      actionSelect.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runGet();
      });
      depthSelect.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runGet();
      });

      footer.appendChild(btnClose);
      footer.appendChild(btnCopyJson);
      footer.appendChild(btnCopyCsv);
      footer.appendChild(btnRun);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      entityInput.focus();
    }
  });
});




document.getElementById("apiTesterUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: async () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,.35);
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
        direction: ltr;
      `;

      const modal = document.createElement("div");
      modal.style.cssText = `
        width: min(1350px, 96vw);
        max-height: 94vh;
        overflow: auto;
        background: #fff;
        border-radius: 16px;
        box-shadow: 0 20px 60px rgba(0,0,0,.25);
        padding: 20px;
        font-family: Segoe UI, Arial, sans-serif;
      `;

      modal.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px;">
          <h2 style="margin:0;font-size:22px;">API Tester / Custom Action Runner</h2>
          <button id="__apiTesterClose" style="
            border:none;
            background:#f3f3f3;
            border-radius:10px;
            padding:8px 12px;
            cursor:pointer;
            font-size:16px;
          ">✖</button>
        </div>

        <div style="display:grid;grid-template-columns: 480px 1fr; gap:20px; align-items:start;">
          <div>
            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
              margin-bottom:14px;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Request</div>

              <div style="padding:14px;">
                <label style="display:block;font-weight:600;margin-bottom:6px;">Method</label>
                <select id="__apiTesterMethod"
                  style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;">
                  <option value="GET">GET</option>
                  <option value="POST">POST</option>
                  <option value="PATCH">PATCH</option>
                  <option value="PUT">PUT</option>
                  <option value="DELETE">DELETE</option>
                </select>

                <label style="display:block;font-weight:600;margin-bottom:6px;">URL</label>
                <input id="__apiTesterUrl" type="text" placeholder="/api/data/v9.2/accounts?$top=5  OR  https://server/api/..."
                  style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:8px;box-sizing:border-box;" />

                <div style="font-size:12px;color:#666;margin-bottom:12px;line-height:1.5;">
                  Supports both relative and absolute URLs.<br/>
                  Examples:<br/>
                  <code>/api/data/v9.2/WhoAmI</code><br/>
                  <code>/api/data/v9.2/accounts?$top=5&$select=name</code><br/>
                  <code>/api/data/v9.2/new_MyCustomApi</code><br/>
                  <code>https://wsgend01:9642/crmapi/mac/v1/targets/DigitalService/0/63884548/1/activeaccount</code>
                </div>

                <label style="display:block;font-weight:600;margin-bottom:6px;">Headers (JSON)</label>
                <textarea id="__apiTesterHeaders" spellcheck="false"
                  style="width:100%;min-height:140px;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;resize:vertical;font-family:Consolas,monospace;font-size:13px;">{
  "Accept": "application/json;odata.include-annotations=*"
}</textarea>

                <label style="display:block;font-weight:600;margin-bottom:6px;">Body (JSON)</label>
                <textarea id="__apiTesterBody" spellcheck="false"
                  style="width:100%;min-height:220px;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;resize:vertical;font-family:Consolas,monospace;font-size:13px;"
                  placeholder='{
  "name": "New Name"
}'></textarea>

                <div style="display:flex;gap:10px;flex-wrap:wrap;">
                  <button id="__apiTesterExecute" style="
                    border:none;
                    background:#107c10;
                    color:white;
                    border-radius:10px;
                    padding:10px 16px;
                    cursor:pointer;
                    font-size:14px;
                  ">Execute</button>

                  <button id="__apiTesterUseCurrent" style="
                    border:none;
                    background:#0f6cbd;
                    color:white;
                    border-radius:10px;
                    padding:10px 16px;
                    cursor:pointer;
                    font-size:14px;
                  ">Use Current Record</button>

                  <button id="__apiTesterPretty" style="
                    border:none;
                    background:#5c2d91;
                    color:white;
                    border-radius:10px;
                    padding:10px 16px;
                    cursor:pointer;
                    font-size:14px;
                  ">Pretty JSON</button>

                  <button id="__apiTesterOpenGet" style="
                    border:none;
                    background:#eaeaea;
                    color:#222;
                    border-radius:10px;
                    padding:10px 16px;
                    cursor:pointer;
                    font-size:14px;
                  ">Open GET</button>

                  <button id="__apiTesterClear" style="
                    border:none;
                    background:#eaeaea;
                    color:#222;
                    border-radius:10px;
                    padding:10px 16px;
                    cursor:pointer;
                    font-size:14px;
                  ">Clear</button>
                </div>
              </div>
            </div>

            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Quick Presets</div>

              <div style="padding:14px;display:flex;gap:8px;flex-wrap:wrap;">
                <button data-preset="whoami" class="__apiPresetBtn" style="border:none;background:#f3f3f3;border-radius:8px;padding:8px 12px;cursor:pointer;">WhoAmI</button>
                <button data-preset="top5accounts" class="__apiPresetBtn" style="border:none;background:#f3f3f3;border-radius:8px;padding:8px 12px;cursor:pointer;">Top 5 Accounts</button>
                <button data-preset="retrieveCurrent" class="__apiPresetBtn" style="border:none;background:#f3f3f3;border-radius:8px;padding:8px 12px;cursor:pointer;">Retrieve Current</button>
                <button data-preset="patchCurrentName" class="__apiPresetBtn" style="border:none;background:#f3f3f3;border-radius:8px;padding:8px 12px;cursor:pointer;">Patch Current Name</button>
                <button data-preset="externalSample" class="__apiPresetBtn" style="border:none;background:#f3f3f3;border-radius:8px;padding:8px 12px;cursor:pointer;">External Sample</button>
              </div>
            </div>
          </div>

          <div>
            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
              margin-bottom:14px;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Status</div>
              <div id="__apiTesterStatus" style="
                padding:14px;
                min-height:70px;
                white-space:pre-wrap;
                font-size:14px;
              ">Ready</div>
            </div>

            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
              margin-bottom:14px;
            ">
              <div style="
                display:flex;justify-content:space-between;align-items:center;gap:12px;
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">
                <div>Response Headers / Meta</div>
                <button id="__apiTesterCopyMeta" style="
                  border:none;background:#eaeaea;border-radius:8px;padding:8px 12px;cursor:pointer;">Copy Meta</button>
              </div>
              <div id="__apiTesterMeta" style="
                padding:14px;
                min-height:120px;
                white-space:pre-wrap;
                word-break:break-word;
                font-family:Consolas, monospace;
                font-size:13px;
              ">{}</div>
            </div>

            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
            ">
              <div style="
                display:flex;justify-content:space-between;align-items:center;gap:12px;
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">
                <div>Response Body</div>
                <div style="display:flex;gap:8px;">
                  <button id="__apiTesterCopyResponse" style="
                    border:none;background:#eaeaea;border-radius:8px;padding:8px 12px;cursor:pointer;">Copy Response</button>
                  <button id="__apiTesterCopyFetch" style="
                    border:none;background:#eaeaea;border-radius:8px;padding:8px 12px;cursor:pointer;">Copy fetch</button>
                </div>
              </div>
              <div id="__apiTesterResponse" style="
                padding:14px;
                min-height:320px;
                white-space:pre-wrap;
                word-break:break-word;
                font-family:Consolas, monospace;
                font-size:13px;
              ">{}</div>
            </div>
          </div>
        </div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      const closeBtn = document.getElementById("__apiTesterClose");
      const methodSelect = document.getElementById("__apiTesterMethod");
      const urlInput = document.getElementById("__apiTesterUrl");
      const headersInput = document.getElementById("__apiTesterHeaders");
      const bodyInput = document.getElementById("__apiTesterBody");
      const executeBtn = document.getElementById("__apiTesterExecute");
      const useCurrentBtn = document.getElementById("__apiTesterUseCurrent");
      const prettyBtn = document.getElementById("__apiTesterPretty");
      const openGetBtn = document.getElementById("__apiTesterOpenGet");
      const clearBtn = document.getElementById("__apiTesterClear");
      const statusBox = document.getElementById("__apiTesterStatus");
      const metaBox = document.getElementById("__apiTesterMeta");
      const responseBox = document.getElementById("__apiTesterResponse");
      const copyResponseBtn = document.getElementById("__apiTesterCopyResponse");
      const copyMetaBtn = document.getElementById("__apiTesterCopyMeta");
      const copyFetchBtn = document.getElementById("__apiTesterCopyFetch");

      closeBtn.onclick = () => overlay.remove();

      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const BASE_API = `${clientUrl}/api/data/v9.2`;

      function normalizeGuid(value) {
        return String(value || "").replace(/[{}]/g, "").trim().toLowerCase();
      }

      function tryParseJson(text, fallback) {
        const trimmed = String(text || "").trim();
        if (!trimmed) return fallback;
        return JSON.parse(trimmed);
      }

      function safePretty(value) {
        try {
          return JSON.stringify(value, null, 2);
        } catch (_) {
          return String(value);
        }
      }

      async function copyText(text) {
        try {
          await navigator.clipboard.writeText(text);
          return true;
        } catch (_) {
          const temp = document.createElement("textarea");
          temp.value = text;
          document.body.appendChild(temp);
          temp.select();
          try {
            document.execCommand("copy");
            return true;
          } catch (_) {
            return false;
          } finally {
            temp.remove();
          }
        }
      }

      async function getEntitySetName(logicalName) {
        const url =
          `${BASE_API}/EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName`;

        const res = await fetch(url, {
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json"
          },
          credentials: "same-origin"
        });

        if (!res.ok) {
          let message = `Failed to get EntitySetName for ${logicalName}`;
          try {
            const err = await res.json();
            message = err?.error?.message || message;
          } catch (_) {}
          throw new Error(message);
        }

        const data = await res.json();
        return data.EntitySetName;
      }

      function tryGetCurrentRecordContext() {
        try {
          if (window.Xrm?.Page?.data?.entity) {
            return {
              entityLogicalName: Xrm.Page.data.entity.getEntityName?.() || null,
              id: normalizeGuid(Xrm.Page.data.entity.getId?.() || ""),
              formType: Xrm.Page.ui?.getFormType?.() || null
            };
          }
        } catch (_) {}

        try {
          const pageEntity = window.parent?.Xrm?.Page?.data?.entity;
          if (pageEntity) {
            return {
              entityLogicalName: pageEntity.getEntityName?.() || null,
              id: normalizeGuid(pageEntity.getId?.() || ""),
              formType: window.parent?.Xrm?.Page?.ui?.getFormType?.() || null
            };
          }
        } catch (_) {}

        return null;
      }

      function resolveUrl(inputUrl) {
        const trimmed = String(inputUrl || "").trim();
        if (!trimmed) throw new Error("URL is required");

        if (/^https?:\/\//i.test(trimmed)) {
          return trimmed;
        }

        if (trimmed.startsWith("/")) return `${clientUrl}${trimmed}`;
        if (trimmed.startsWith("api/")) return `${clientUrl}/${trimmed}`;

        return `${clientUrl}/api/data/v9.2/${trimmed}`;
      }

      function buildFetchSnippet(method, inputUrl, headers, bodyText) {
        const absUrl = resolveUrl(inputUrl);
        const options = {
          method,
          headers
        };

        if (!["GET", "DELETE"].includes(method) && String(bodyText || "").trim()) {
          options.body = bodyText;
        }

        return `fetch(${JSON.stringify(absUrl)}, ${safePretty(options)})
  .then(async (res) => {
    const text = await res.text();
    let data = text;
    try { data = JSON.parse(text); } catch (_) {}
    console.log("status", res.status, res.statusText);
    console.log(data);
  });`;
      }

      async function executeRequest() {
        const method = methodSelect.value;
        const inputUrl = String(urlInput.value || "").trim();
        const absoluteUrl = resolveUrl(inputUrl);

        let customHeaders = {};
        try {
          customHeaders = tryParseJson(headersInput.value, {});
          if (typeof customHeaders !== "object" || Array.isArray(customHeaders) || customHeaders === null) {
            throw new Error("Headers must be a JSON object");
          }
        } catch (err) {
          throw new Error(`Invalid headers JSON: ${err.message}`);
        }

        const isExternal = /^https?:\/\//i.test(inputUrl);

        const defaultHeaders = isExternal
          ? {}
          : {
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
              "Accept": "application/json;odata.include-annotations=*"
            };

        const headers = {
          ...defaultHeaders,
          ...customHeaders
        };

        let requestBody = null;
        const rawBody = String(bodyInput.value || "").trim();

        if (!["GET", "DELETE"].includes(method) && rawBody) {
          try {
            requestBody = JSON.parse(rawBody);
          } catch (err) {
            throw new Error(`Invalid body JSON: ${err.message}`);
          }

          if (!headers["Content-Type"] && !headers["content-type"]) {
            headers["Content-Type"] = "application/json; charset=utf-8";
          }
        }

        const startedAt = performance.now();
const res = await fetch(absoluteUrl, {
  method,
  headers,
  body: requestBody !== null ? JSON.stringify(requestBody) : undefined,
  credentials: isExternal ? "omit" : "include"
});

        const endedAt = performance.now();

        const responseHeaders = {};
        try {
          for (const [k, v] of res.headers.entries()) {
            responseHeaders[k] = v;
          }
        } catch (_) {}

        const rawText = await res.text();

        let responseBody;
        if (!rawText) {
          responseBody = null;
        } else {
          try {
            responseBody = JSON.parse(rawText);
          } catch (_) {
            responseBody = rawText;
          }
        }

        const meta = {
          ok: res.ok,
          status: res.status,
          statusText: res.statusText,
          durationMs: Math.round(endedAt - startedAt),
          isExternal,
          request: {
            method,
            inputUrl,
            absoluteUrl
          },
          responseHeaders
        };

        return { meta, responseBody, ok: res.ok };
      }

      async function fillCurrentRecordUrl() {
        const ctx = tryGetCurrentRecordContext();

        if (!ctx?.entityLogicalName) {
          throw new Error("Could not detect current record context");
        }

        if (!ctx?.id) {
          throw new Error("Current form has no saved record id");
        }

        const entitySetName = await getEntitySetName(ctx.entityLogicalName);

        methodSelect.value = "GET";
        urlInput.value = `/api/data/v9.2/${entitySetName}(${ctx.id})`;
        bodyInput.value = "";
        statusBox.textContent = `Loaded current record URL for ${ctx.entityLogicalName}`;
      }

      function prettyJsonEditors() {
        try {
          const headers = tryParseJson(headersInput.value, {});
          headersInput.value = JSON.stringify(headers, null, 2);
        } catch (_) {}

        try {
          const body = tryParseJson(bodyInput.value, null);
          if (body !== null) {
            bodyInput.value = JSON.stringify(body, null, 2);
          }
        } catch (_) {}
      }

      function clearAll() {
        methodSelect.value = "GET";
        urlInput.value = "";
        headersInput.value = `{
  "Accept": "application/json;odata.include-annotations=*"
}`;
        bodyInput.value = "";
        metaBox.textContent = "{}";
        responseBox.textContent = "{}";
        statusBox.textContent = "Ready";
      }

      function applyPreset(name) {
        const ctx = tryGetCurrentRecordContext();

        if (name === "whoami") {
          methodSelect.value = "GET";
          urlInput.value = "/api/data/v9.2/WhoAmI";
          bodyInput.value = "";
          return;
        }

        if (name === "top5accounts") {
          methodSelect.value = "GET";
          urlInput.value = "/api/data/v9.2/accounts?$top=5&$select=name,accountid";
          bodyInput.value = "";
          return;
        }

        if (name === "retrieveCurrent") {
          if (!ctx?.entityLogicalName || !ctx?.id) {
            statusBox.textContent = "❌ No saved current record detected";
            return;
          }

          getEntitySetName(ctx.entityLogicalName)
            .then((entitySetName) => {
              methodSelect.value = "GET";
              urlInput.value = `/api/data/v9.2/${entitySetName}(${ctx.id})`;
              bodyInput.value = "";
              statusBox.textContent = "Loaded preset: Retrieve Current";
            })
            .catch((err) => {
              statusBox.textContent = `❌ ${err.message}`;
            });

          return;
        }

        if (name === "patchCurrentName") {
          if (!ctx?.entityLogicalName || !ctx?.id) {
            statusBox.textContent = "❌ No saved current record detected";
            return;
          }

          getEntitySetName(ctx.entityLogicalName)
            .then((entitySetName) => {
              methodSelect.value = "PATCH";
              urlInput.value = `/api/data/v9.2/${entitySetName}(${ctx.id})`;
              bodyInput.value = `{
  "name": "Updated from API Tester"
}`;
              statusBox.textContent = "Loaded preset: Patch Current Name";
            })
            .catch((err) => {
              statusBox.textContent = `❌ ${err.message}`;
            });

          return;
        }

        if (name === "externalSample") {
          methodSelect.value = "GET";
          urlInput.value = "https://wsgend01:9642/crmapi/mac/v1/targets/DigitalService/0/63884548/1/activeaccount";
          headersInput.value = `{
  "Accept": "application/json"
}`;
          bodyInput.value = "";
          statusBox.textContent = "Loaded preset: External Sample";
          return;
        }
      }

      executeBtn.addEventListener("click", async () => {
        executeBtn.disabled = true;
        statusBox.textContent = "Executing...";
        metaBox.textContent = "{}";
        responseBox.textContent = "{}";

        try {
          const { meta, responseBody, ok } = await executeRequest();

          metaBox.textContent = safePretty(meta);
          responseBox.textContent = safePretty(responseBody);

          statusBox.textContent =
            `${ok ? "✅" : "❌"} ${meta.status} ${meta.statusText}\n` +
            `Duration: ${meta.durationMs} ms\n` +
            `${meta.request.method} ${meta.request.inputUrl}`;
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
          responseBox.textContent = safePretty({
            error: err.message
          });
        } finally {
          executeBtn.disabled = false;
        }
      });

      useCurrentBtn.addEventListener("click", async () => {
        useCurrentBtn.disabled = true;

        try {
          await fillCurrentRecordUrl();
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
        } finally {
          useCurrentBtn.disabled = false;
        }
      });

      prettyBtn.addEventListener("click", () => {
        prettyJsonEditors();
        statusBox.textContent = "JSON pretty-printed";
      });

      openGetBtn.addEventListener("click", () => {
        try {
          const method = methodSelect.value;
          if (method !== "GET") {
            throw new Error("Open GET works only for GET requests");
          }

          const absoluteUrl = resolveUrl(urlInput.value);
          window.open(absoluteUrl, "_blank");
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
        }
      });

      clearBtn.addEventListener("click", clearAll);

      copyResponseBtn.addEventListener("click", async () => {
        const ok = await copyText(responseBox.textContent || "");
        statusBox.textContent = ok ? "Response copied" : "Failed to copy response";
      });

      copyMetaBtn.addEventListener("click", async () => {
        const ok = await copyText(metaBox.textContent || "");
        statusBox.textContent = ok ? "Meta copied" : "Failed to copy meta";
      });

      copyFetchBtn.addEventListener("click", async () => {
        try {
          const method = methodSelect.value;
          const inputUrl = String(urlInput.value || "").trim();

          let headers = {};
          try {
            headers = tryParseJson(headersInput.value, {});
          } catch (_) {}

          const snippet = buildFetchSnippet(method, inputUrl, headers, bodyInput.value);
          const ok = await copyText(snippet);
          statusBox.textContent = ok ? "fetch snippet copied" : "Failed to copy fetch snippet";
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
        }
      });

      document.querySelectorAll(".__apiPresetBtn").forEach((btn) => {
        btn.addEventListener("click", () => applyPreset(btn.dataset.preset));
      });

      const ctx = tryGetCurrentRecordContext();
      if (ctx?.entityLogicalName && ctx?.id) {
        statusBox.textContent = `Ready\nCurrent record: ${ctx.entityLogicalName} (${ctx.id})`;
      }
    }
  });
});

document.getElementById("securityRolesUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: async () => {
      document.getElementById("__d365helper_modal")?.remove();

      const CACHE_TTL_MS = 10 * 60 * 1000;
      const USER_SEARCH_MIN_LENGTH = 2;
      const USER_SEARCH_TOP = 50;
      const ROLE_SEARCH_TOP = 300;

      const collator = new Intl.Collator("he");

      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const BASE_URL = `${clientUrl}/api/data/v9.2`;

      let allRoles = [];
      let businessUnitsMap = {};
      let currentSelectedUserId = null;
      let currentUserRoles = [];
      let userSearchAbortController = null;

      function normalizeGuid(id) {
        return String(id || "").replace(/[{}]/g, "").trim();
      }

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");
      }

      function escapeODataString(str) {
        return String(str ?? "").replace(/'/g, "''");
      }

      function debounce(fn, delay = 300) {
        let timer = null;
        return (...args) => {
          clearTimeout(timer);
          timer = setTimeout(() => fn(...args), delay);
        };
      }

      async function fetchJSON(url, options = {}) {
        const res = await fetch(url, {
          ...options,
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Accept: "application/json",
            ...(options.headers || {})
          },
          credentials: "same-origin"
        });

        if (!res.ok) {
          let message = `${res.status}`;
          try {
            const err = await res.json();
            message = err?.error?.message || message;
          } catch (_) {}
          throw new Error(message);
        }

        if (res.status === 204) return null;
        return await res.json();
      }

      async function fetchAllPages(url) {
        const all = [];
        let nextUrl = url;

        while (nextUrl) {
          const data = await fetchJSON(nextUrl);
          all.push(...(data?.value || []));
          nextUrl = data?.["@odata.nextLink"] || null;
        }

        return all;
      }

      function getCurrentUserId() {
        return normalizeGuid(Xrm.Utility.getGlobalContext().userSettings.userId);
      }

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,.35);
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
        direction: rtl;
      `;

      const modal = document.createElement("div");
      modal.style.cssText = `
        width: min(1200px, 96vw);
        max-height: 92vh;
        overflow: auto;
        background: #fff;
        border-radius: 16px;
        box-shadow: 0 20px 60px rgba(0,0,0,.25);
        padding: 20px;
        font-family: Segoe UI, Arial, sans-serif;
      `;

      modal.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px;">
          <h2 style="margin:0;font-size:22px;">ניהול תפקידי אבטחה למשתמש</h2>
          <button id="__rolesClose" style="border:none;background:#f3f3f3;border-radius:10px;padding:8px 12px;cursor:pointer;font-size:16px;">✖</button>
        </div>

        <div style="margin-bottom:14px;padding:10px 12px;background:#f7f7f7;border-radius:10px;">
          <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:600;">
            <input id="__rolesUseCurrentUser" type="checkbox" />
            בצע עליי - המשתמש המחובר
          </label>
        </div>

        <div style="display:grid;grid-template-columns:380px 1fr;gap:20px;align-items:start;">
          <div>
            <label style="display:block;font-weight:600;margin-bottom:8px;">חיפוש משתמש</label>
            <input id="__rolesUserSearch" type="text" placeholder="הקלד לפחות 2 תווים"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:10px;box-sizing:border-box;" />

            <select id="__rolesUserSelect" size="15"
              style="width:100%;padding:8px;border:1px solid #ccc;border-radius:10px;box-sizing:border-box;"></select>

            <label style="display:block;font-weight:600;margin:16px 0 8px;">חיפוש תפקיד להוספה</label>
            <input id="__rolesRoleSearch" type="text" placeholder="חפש תפקיד"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:10px;box-sizing:border-box;" />

            <label style="display:block;font-weight:600;margin-bottom:8px;">סינון לפי יחידה עסקית</label>
            <select id="__rolesBuFilter"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:10px;box-sizing:border-box;">
              <option value="">הכל</option>
            </select>

            <select id="__rolesRoleSelect" size="6"
              style="width:100%;padding:8px;border:1px solid #ccc;border-radius:10px;box-sizing:border-box;"></select>

         <div id="__rolesActions" style="
  display:grid;
  grid-template-columns:repeat(3, 1fr);
  gap:10px;
  margin-top:14px;
  position:sticky;
  bottom:0;
  background:#fff;
  padding-top:10px;
  z-index:2;
">
  <button id="__rolesShow" style="border:none;background:#0f6cbd;color:white;border-radius:10px;padding:10px 16px;cursor:pointer;font-size:14px;">הצג תפקידים</button>
  <button id="__rolesAdd" style="border:none;background:#107c10;color:white;border-radius:10px;padding:10px 16px;cursor:pointer;font-size:14px;">הוסף תפקיד</button>
  <button id="__rolesClear" style="border:none;background:#eaeaea;color:#222;border-radius:10px;padding:10px 16px;cursor:pointer;font-size:14px;">נקה</button>
</div>
          </div>

          <div>
            <div id="__rolesSelectedUser" style="margin-bottom:12px;padding:10px 12px;background:#f7f7f7;border-radius:10px;min-height:20px;font-weight:600;">
              לא נבחר משתמש
            </div>

            <div style="border:1px solid #ddd;border-radius:12px;overflow:hidden;background:#fff;">
              <div style="display:grid;grid-template-columns:1.5fr 1fr 140px;background:#f3f3f3;font-weight:700;padding:12px;border-bottom:1px solid #ddd;">
                <div>תפקיד אבטחה</div>
                <div>יחידה עסקית</div>
                <div>פעולה</div>
              </div>
              <div id="__rolesRows" style="max-height:560px;overflow:auto;"></div>
            </div>
          </div>
        </div>

        <div id="__rolesStatus" style="margin-top:14px;padding:10px 12px;background:#f7f7f7;border-radius:10px;min-height:22px;white-space:pre-wrap;font-size:14px;"></div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      const userSearchInput = document.getElementById("__rolesUserSearch");
      const userSelect = document.getElementById("__rolesUserSelect");
      const roleSearchInput = document.getElementById("__rolesRoleSearch");
      const roleSelect = document.getElementById("__rolesRoleSelect");
      const buFilterSelect = document.getElementById("__rolesBuFilter");
      const showBtn = document.getElementById("__rolesShow");
      const addBtn = document.getElementById("__rolesAdd");
      const clearBtn = document.getElementById("__rolesClear");
      const statusBox = document.getElementById("__rolesStatus");
      const selectedUserBox = document.getElementById("__rolesSelectedUser");
      const rowsBox = document.getElementById("__rolesRows");
      const useCurrentUserCheckbox = document.getElementById("__rolesUseCurrentUser");

      document.getElementById("__rolesClose").onclick = () => overlay.remove();

      function setStatus(text) {
        statusBox.textContent = text || "";
      }

      function clearUsers(message = "הקלד לפחות 2 תווים לחיפוש משתמש.") {
        userSelect.innerHTML = "";
        const option = document.createElement("option");
        option.textContent = message;
        option.disabled = true;
        userSelect.appendChild(option);
      }

      function getSelectedUserId() {
        if (useCurrentUserCheckbox.checked) return getCurrentUserId();

        const selectedUser = userSelect.options[userSelect.selectedIndex];
        if (!selectedUser || selectedUser.disabled || !selectedUser.dataset.userid) {
          throw new Error("צריך לבחור משתמש מהרשימה או לסמן 'בצע עליי'.");
        }

        return selectedUser.dataset.userid;
      }

      function getSelectedUserLabel() {
        if (useCurrentUserCheckbox.checked) return "משתמש נבחר: המשתמש המחובר";

        const selectedUser = userSelect.options[userSelect.selectedIndex];
        if (!selectedUser || selectedUser.disabled) return "לא נבחר משתמש";

        return `משתמש נבחר: ${selectedUser.dataset.fullname || ""} | ${selectedUser.dataset.domainname || "ללא domain"} | ${selectedUser.dataset.isdisabled === "true" ? "לא פעיל" : "פעיל"}`;
      }

      function toggleUserSelectionState() {
        const disabled = useCurrentUserCheckbox.checked;

        userSearchInput.disabled = disabled;
        userSelect.disabled = disabled;
        userSearchInput.style.opacity = disabled ? "0.6" : "1";
        userSelect.style.opacity = disabled ? "0.6" : "1";

        if (disabled) {
          selectedUserBox.textContent = "משתמש נבחר: המשתמש המחובר";
        }
      }

      async function searchUsers(searchText) {
        const q = searchText.trim();

        if (q.length < USER_SEARCH_MIN_LENGTH) {
          clearUsers();
          return;
        }

        if (userSearchAbortController) {
          userSearchAbortController.abort();
        }

        userSearchAbortController = new AbortController();

      const safeQ = escapeODataString(q);

const filterParts = [
  `contains(fullname,'${safeQ}')`,
  `contains(domainname,'${safeQ}')`,
  `contains(internalemailaddress,'${safeQ}')`
];

        const url =
          `${BASE_URL}/systemusers` +
          `?$select=systemuserid,fullname,domainname,isdisabled,internalemailaddress` +
          `&$filter=${encodeURIComponent(filterParts.join(" or "))}` +
          `&$orderby=fullname asc` +
          `&$top=${USER_SEARCH_TOP}`;

        userSelect.innerHTML = "";
        const loading = document.createElement("option");
        loading.disabled = true;
        loading.textContent = "מחפש משתמשים...";
        userSelect.appendChild(loading);

        try {
          const data = await fetchJSON(url, { signal: userSearchAbortController.signal });
          renderUsers(data?.value || []);
        } catch (err) {
          if (err.name === "AbortError") return;
          clearUsers(`שגיאה בחיפוש משתמשים: ${err.message}`);
        }
      }

      function renderUsers(users) {
        userSelect.innerHTML = "";

        if (!users.length) {
          clearUsers("לא נמצאו משתמשים.");
          return;
        }

        const fragment = document.createDocumentFragment();

        for (const u of users) {
          const domain = u.domainname || "";
          const username = domain ? domain.replace(/@mac\.org\.il$/i, "") : "";

          const option = document.createElement("option");
          option.value = u.systemuserid;
          option.textContent = `${u.fullname || "(ללא שם)"} | ${username || domain || "ללא domain"} | ${u.isdisabled ? "לא פעיל" : "פעיל"}`;
          option.dataset.userid = u.systemuserid;
          option.dataset.fullname = u.fullname || "";
          option.dataset.domainname = domain;
          option.dataset.email = u.internalemailaddress || "";
          option.dataset.isdisabled = String(!!u.isdisabled);

          fragment.appendChild(option);
        }

        userSelect.appendChild(fragment);
        userSelect.selectedIndex = 0;
      }

      async function loadBusinessUnitsMapCached() {
        const cache = window.__d365SecurityRolesBuCache;
        const now = Date.now();

        if (cache?.createdAt && now - cache.createdAt < CACHE_TTL_MS) {
          businessUnitsMap = cache.businessUnitsMap || {};
          return cache.businessUnits || [];
        }

        const businessUnits = await fetchAllPages(
          `${BASE_URL}/businessunits?$select=businessunitid,name&$orderby=name asc`
        );

        businessUnitsMap = {};
        for (const bu of businessUnits) {
          businessUnitsMap[normalizeGuid(bu.businessunitid)] = bu.name || "";
        }

        window.__d365SecurityRolesBuCache = {
          createdAt: now,
          businessUnits,
          businessUnitsMap
        };

        return businessUnits;
      }

      async function loadRolesCached() {
        const cache = window.__d365SecurityRolesCache;
        const now = Date.now();

        if (cache?.createdAt && now - cache.createdAt < CACHE_TTL_MS) {
          allRoles = cache.allRoles || [];
          businessUnitsMap = cache.businessUnitsMap || {};
          return;
        }

        await loadBusinessUnitsMapCached();

        const roles = await fetchAllPages(
          `${BASE_URL}/roles?$select=roleid,name,_businessunitid_value&$orderby=name asc`
        );

        allRoles = roles
          .filter(r => r.name)
          .map(r => {
            const buId = normalizeGuid(r._businessunitid_value);
            return {
              id: normalizeGuid(r.roleid),
              name: r.name || "",
              businessunitid: buId,
              buName: businessUnitsMap[buId] || ""
            };
          })
          .sort((a, b) => {
            const nameCmp = collator.compare(a.name || "", b.name || "");
            if (nameCmp !== 0) return nameCmp;
            return collator.compare(a.buName || "", b.buName || "");
          });

        window.__d365SecurityRolesCache = {
          createdAt: now,
          allRoles,
          businessUnitsMap
        };
      }

      function populateBuFilter() {
        buFilterSelect.innerHTML = `<option value="">הכל</option>`;

        const buList = [...new Map(
          allRoles
            .filter(r => r.businessunitid && r.buName)
            .map(r => [r.businessunitid, { id: r.businessunitid, name: r.buName }])
        ).values()].sort((a, b) => collator.compare(a.name, b.name));

        const fragment = document.createDocumentFragment();

        for (const bu of buList) {
          const opt = document.createElement("option");
          opt.value = bu.id;
          opt.textContent = bu.name;
          fragment.appendChild(opt);
        }

        buFilterSelect.appendChild(fragment);
      }

      function renderRoles(searchText = "", buFilter = "") {
        const q = searchText.trim().toLowerCase();
        roleSelect.innerHTML = "";

        const filtered = [];
        for (const role of allRoles) {
          if (buFilter && role.businessunitid !== buFilter) continue;
          if (q && !role.name.toLowerCase().includes(q)) continue;

          filtered.push(role);
          if (filtered.length >= ROLE_SEARCH_TOP) break;
        }

        if (!filtered.length) {
          const option = document.createElement("option");
          option.textContent = "לא נמצאו תפקידים.";
          option.disabled = true;
          roleSelect.appendChild(option);
          return;
        }

        const fragment = document.createDocumentFragment();

        for (const role of filtered) {
          const option = document.createElement("option");
          option.value = role.name;
          option.textContent = role.buName ? `${role.name} | ${role.buName}` : role.name;
          option.dataset.roleid = role.id;
          option.dataset.rolename = role.name;
          option.dataset.businessunitid = role.businessunitid;
          option.dataset.buname = role.buName || "";
          fragment.appendChild(option);
        }

        roleSelect.appendChild(fragment);
        roleSelect.selectedIndex = 0;
      }

      async function getDirectUserRoles(userId) {
        const url =
          `${BASE_URL}/systemusers(${userId})` +
          `?$select=fullname,domainname,isdisabled` +
          `&$expand=systemuserroles_association($select=roleid,name,_businessunitid_value)`;

        const userData = await fetchJSON(url);

        return {
          user: {
            fullname: userData.fullname || "",
            domainname: userData.domainname || "",
            isdisabled: !!userData.isdisabled
          },
          roles: (userData.systemuserroles_association || []).map(r => {
            const buId = normalizeGuid(r._businessunitid_value);
            return {
              roleid: normalizeGuid(r.roleid),
              name: r.name || "",
              businessunitid: buId,
              buName: businessUnitsMap[buId] || "",
              source: "ישיר"
            };
          })
        };
      }

      async function getTeamRolesViaFetchXml(userId) {
        const fetchXml = `
          <fetch distinct="true">
            <entity name="role">
              <attribute name="roleid" />
              <attribute name="name" />
              <attribute name="businessunitid" />
              <link-entity name="teamroles" from="roleid" to="roleid" intersect="true">
                <link-entity name="team" from="teamid" to="teamid" alias="team">
                  <attribute name="name" />
                  <link-entity name="teammembership" from="teamid" to="teamid" intersect="true">
                    <link-entity name="systemuser" from="systemuserid" to="systemuserid">
                      <filter>
                        <condition attribute="systemuserid" operator="eq" value="${userId}" />
                      </filter>
                    </link-entity>
                  </link-entity>
                </link-entity>
              </link-entity>
            </entity>
          </fetch>`.trim();

        const result = await Xrm.WebApi.retrieveMultipleRecords(
          "role",
          `?fetchXml=${encodeURIComponent(fetchXml)}`
        );

        return (result.entities || []).map(r => {
          const roleId = normalizeGuid(r.roleid);
          const buId = normalizeGuid(
            r._businessunitid_value ||
            r.businessunitid ||
            r["businessunitid"] ||
            ""
          );

          const teamName =
            r["team.name"] ||
            r["team.name@OData.Community.Display.V1.FormattedValue"] ||
            "";

          return {
            roleid: roleId,
            name: r.name || "",
            businessunitid: buId,
            buName:
              businessUnitsMap[buId] ||
              r["_businessunitid_value@OData.Community.Display.V1.FormattedValue"] ||
              "",
            source: teamName ? `צוות: ${teamName}` : "צוות"
          };
        });
      }

      async function getUserRoles(userId) {
        const [directResult, teamRoles] = await Promise.all([
          getDirectUserRoles(userId),
          getTeamRolesViaFetchXml(userId)
        ]);

        const roleMap = new Map();

        for (const role of directResult.roles) {
          roleMap.set(role.roleid, role);
        }

        for (const role of teamRoles) {
          if (!roleMap.has(role.roleid)) {
            roleMap.set(role.roleid, role);
          }
        }

        return {
          user: directResult.user,
          roles: [...roleMap.values()]
        };
      }

      async function addUserRole(userId, roleId) {
        await fetchJSON(`${BASE_URL}/systemusers(${userId})/systemuserroles_association/$ref`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            "@odata.id": `${BASE_URL}/roles(${roleId})`
          })
        });
      }

      async function removeUserRole(userId, roleId) {
        await fetchJSON(`${BASE_URL}/systemusers(${userId})/systemuserroles_association(${roleId})/$ref`, {
          method: "DELETE",
          headers: { "If-Match": "*" }
        });
      }

      function renderRolesRows(roles, userId) {
        rowsBox.innerHTML = "";

        if (!roles.length) {
          const empty = document.createElement("div");
          empty.style.cssText = "padding:14px;";
          empty.textContent = "לא נמצאו תפקידי אבטחה למשתמש זה.";
          rowsBox.appendChild(empty);
          return;
        }

        const sortedRoles = [...roles].sort((a, b) => {
          const nameCompare = collator.compare(a.name || "", b.name || "");
          if (nameCompare !== 0) return nameCompare;
          return collator.compare(a.buName || "", b.buName || "");
        });

        const fragment = document.createDocumentFragment();

        for (const role of sortedRoles) {
          const row = document.createElement("div");
          row.style.cssText = `
            display:grid;
            grid-template-columns:1.5fr 1fr 140px;
            padding:12px;
            border-bottom:1px solid #eee;
            align-items:center;
          `;

          const sourceLabel =
            role.source && role.source !== "ישיר"
              ? ` <span style="font-size:12px;color:#666;">(${escapeHtml(role.source)})</span>`
              : "";

          row.innerHTML = `
            <div>${escapeHtml(role.name || "")}${sourceLabel}</div>
            <div>${escapeHtml(role.buName || businessUnitsMap[role.businessunitid] || role.businessunitid || "")}</div>
            <div>
              ${
                role.source === "ישיר"
                  ? `<button 
                      type="button"
                      data-action="remove-role"
                      data-roleid="${escapeHtml(role.roleid)}"
                      data-rolename="${escapeHtml(role.name)}"
                      style="border:none;background:#d13438;color:white;border-radius:8px;padding:8px 12px;cursor:pointer;font-size:14px;">
                      הסר
                    </button>`
                  : `<span style="font-size:12px;color:#888;">דרך צוות</span>`
              }
            </div>
          `;

          fragment.appendChild(row);
        }

        rowsBox.appendChild(fragment);
      }

      async function refreshUserRoles(userId, userLabel) {
        currentSelectedUserId = userId;
        currentUserRoles = [];

        rowsBox.innerHTML = "";
        selectedUserBox.textContent = userLabel || getSelectedUserLabel();
        setStatus("טוען תפקידי אבטחה...");

        const result = await getUserRoles(userId);
        currentUserRoles = result.roles;

        renderRolesRows(currentUserRoles, userId);

        setStatus(
          `✅ נמצאו ${currentUserRoles.length} תפקידי אבטחה עבור ${result.user.fullname}` +
          (result.user.domainname ? ` (${result.user.domainname})` : "")
        );
      }

      rowsBox.addEventListener("click", async event => {
        const btn = event.target.closest("[data-action='remove-role']");
        if (!btn) return;

        const roleId = btn.dataset.roleid;
        const roleName = btn.dataset.rolename || "";

        if (!currentSelectedUserId) {
          setStatus("לא נבחר משתמש.");
          return;
        }

        if (!confirm(`להסיר את התפקיד "${roleName}" מהמשתמש?`)) return;

        btn.disabled = true;
        setStatus(`מסיר את התפקיד "${roleName}"...`);

        try {
          await removeUserRole(currentSelectedUserId, roleId);

          currentUserRoles = currentUserRoles.filter(r => r.roleid !== roleId);
          renderRolesRows(currentUserRoles, currentSelectedUserId);

          setStatus(`✅ התפקיד "${roleName}" הוסר בהצלחה`);
        } catch (err) {
          setStatus(`❌ שגיאה בהסרה: ${err.message}`);
        } finally {
          btn.disabled = false;
        }
      });

      useCurrentUserCheckbox.addEventListener("change", toggleUserSelectionState);

      userSearchInput.addEventListener(
        "input",
        debounce(() => searchUsers(userSearchInput.value), 300)
      );

      roleSearchInput.addEventListener(
        "input",
        debounce(() => renderRoles(roleSearchInput.value, buFilterSelect.value), 150)
      );

      buFilterSelect.addEventListener("change", () => {
        renderRoles(roleSearchInput.value, buFilterSelect.value);
      });

      clearBtn.addEventListener("click", () => {
        userSearchInput.value = "";
        roleSearchInput.value = "";
        buFilterSelect.value = "";
        currentSelectedUserId = null;
        currentUserRoles = [];

        clearUsers();
        renderRoles();
        selectedUserBox.textContent = "לא נבחר משתמש";
        rowsBox.innerHTML = "";
        setStatus(`✅ נטענו ${allRoles.length} תפקידים. משתמשים נטענים לפי חיפוש בלבד.`);
      });

      showBtn.addEventListener("click", async () => {
        showBtn.disabled = true;

        try {
          const userId = getSelectedUserId();
          await refreshUserRoles(userId, getSelectedUserLabel());
        } catch (err) {
          setStatus(`❌ שגיאה: ${err.message}`);
        } finally {
          showBtn.disabled = false;
        }
      });

      addBtn.addEventListener("click", async () => {
        const selectedRole = roleSelect.options[roleSelect.selectedIndex];

        if (!selectedRole || selectedRole.disabled) {
          setStatus("צריך לבחור תפקיד מהרשימה.");
          return;
        }

        addBtn.disabled = true;

        try {
          const userId = getSelectedUserId();
          const roleId = normalizeGuid(selectedRole.dataset.roleid);
          const roleName = selectedRole.dataset.rolename || selectedRole.value;
          const buId = normalizeGuid(selectedRole.dataset.businessunitid);
          const buName = selectedRole.dataset.buname || businessUnitsMap[buId] || "";

          setStatus(`מקצה את התפקיד "${roleName}"...`);

          await addUserRole(userId, roleId);

          const exists = currentUserRoles.some(r => r.roleid === roleId);

          if (currentSelectedUserId === userId && !exists) {
            currentUserRoles.push({
              roleid: roleId,
              name: roleName,
              businessunitid: buId,
              buName,
              source: "ישיר"
            });

            renderRolesRows(currentUserRoles, userId);
          } else if (currentSelectedUserId !== userId) {
            await refreshUserRoles(userId, getSelectedUserLabel());
          }

          setStatus(`✅ התפקיד "${roleName}" הוקצה בהצלחה`);
        } catch (err) {
          setStatus(`❌ שגיאה בהוספה: ${err.message}`);
        } finally {
          addBtn.disabled = false;
        }
      });

      try {
        setStatus("טוען תפקידים ויחידות עסקיות...");
        clearUsers();

        await loadRolesCached();

        renderRoles();
        populateBuFilter();
        toggleUserSelectionState();

        setStatus(`✅ נטענו ${allRoles.length} תפקידים. חיפוש משתמשים מתבצע לפי הקלדה בלבד.`);
      } catch (err) {
        setStatus(`❌ שגיאה בטעינה: ${err.message}`);
      }
    }
  });
});

document.getElementById("quickUpdateFieldUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: async () => {
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,.35);
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
        direction: ltr;
      `;

      const modal = document.createElement("div");
      modal.style.cssText = `
        width: min(1200px, 96vw);
        max-height: 92vh;
        overflow: auto;
        background: #fff;
        border-radius: 16px;
        box-shadow: 0 20px 60px rgba(0,0,0,.25);
        padding: 20px;
        font-family: Segoe UI, Arial, sans-serif;
      `;

      modal.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:16px;">
          <h2 style="margin:0;font-size:22px;">Quick Update Field</h2>
          <button id="__quickUpdateClose" style="
            border:none;
            background:#f3f3f3;
            border-radius:10px;
            padding:8px 12px;
            cursor:pointer;
            font-size:16px;
          ">✖</button>
        </div>

        <div style="display:grid;grid-template-columns: 440px 1fr; gap:20px; align-items:start;">
          <div>
            <label style="display:block;font-weight:600;margin-bottom:6px;">Entity logical name</label>
            <input id="__quickUpdateEntity" type="text" placeholder="e.g. account"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;" />

            <label style="display:block;font-weight:600;margin-bottom:6px;">Record GUID</label>
            <input id="__quickUpdateId" type="text" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;" />

            <label style="display:block;font-weight:600;margin-bottom:6px;">Field logical name</label>
            <input id="__quickUpdateField" type="text" placeholder="e.g. name / ey_xxx / ownerid"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;" />

            <label style="display:block;font-weight:600;margin-bottom:6px;">Field type</label>
            <select id="__quickUpdateType"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;">
              <option value="string">String</option>
              <option value="memo">Memo</option>
              <option value="integer">Whole Number</option>
              <option value="decimal">Decimal</option>
              <option value="double">Double</option>
              <option value="boolean">Two Options</option>
              <option value="datetime">Date Time</option>
              <option value="optionset">Option Set (integer)</option>
              <option value="lookup">Lookup (@odata.bind)</option>
            </select>

            <label style="display:block;font-weight:600;margin-bottom:6px;">New value</label>
            <textarea id="__quickUpdateValue" spellcheck="false"
              style="width:100%;min-height:120px;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:12px;box-sizing:border-box;resize:vertical;"
              placeholder="Enter new value"></textarea>

            <div id="__quickUpdateLookupHelp" style="
              display:none;
              margin-bottom:12px;
              padding:12px;
              background:#f7f7f7;
              border-radius:10px;
              font-size:13px;
              line-height:1.5;
              white-space:pre-wrap;
              color:#444;
            ">For lookup use JSON:
{
  "entitySetName": "systemusers",
  "id": "GUID"
}</div>

            <div style="display:flex;gap:10px;flex-wrap:wrap;">
              <button id="__quickUpdateLoadCurrent" style="
                border:none;
                background:#0f6cbd;
                color:white;
                border-radius:10px;
                padding:10px 16px;
                cursor:pointer;
                font-size:14px;
              ">Load Current Value</button>

              <button id="__quickUpdateSubmit" style="
                border:none;
                background:#107c10;
                color:white;
                border-radius:10px;
                padding:10px 16px;
                cursor:pointer;
                font-size:14px;
              ">Update</button>

              <button id="__quickUpdateClear" style="
                border:none;
                background:#eaeaea;
                color:#222;
                border-radius:10px;
                padding:10px 16px;
                cursor:pointer;
                font-size:14px;
              ">Clear</button>
            </div>
          </div>

          <div>
            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
              margin-bottom:14px;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Current Value</div>
              <div id="__quickUpdateCurrentValue" style="
                padding:14px;
                min-height:120px;
                white-space:pre-wrap;
                word-break:break-word;
                font-family:Consolas, monospace;
                font-size:13px;
              ">Not loaded</div>
            </div>

            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
              margin-bottom:14px;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Payload Preview</div>
              <div id="__quickUpdatePayloadPreview" style="
                padding:14px;
                min-height:120px;
                white-space:pre-wrap;
                word-break:break-word;
                font-family:Consolas, monospace;
                font-size:13px;
              ">{}</div>
            </div>

            <div style="
              border:1px solid #ddd;
              border-radius:12px;
              overflow:hidden;
              background:#fff;
            ">
              <div style="
                background:#f3f3f3;
                font-weight:700;
                padding:12px;
                border-bottom:1px solid #ddd;
              ">Status</div>
              <div id="__quickUpdateStatus" style="
                padding:14px;
                min-height:80px;
                white-space:pre-wrap;
                font-size:14px;
              ">Ready</div>
            </div>
          </div>
        </div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      const closeModal = () => overlay.remove();
      document.getElementById("__quickUpdateClose").onclick = closeModal;

      const entityInput = document.getElementById("__quickUpdateEntity");
      const idInput = document.getElementById("__quickUpdateId");
      const fieldInput = document.getElementById("__quickUpdateField");
      const typeSelect = document.getElementById("__quickUpdateType");
      const valueInput = document.getElementById("__quickUpdateValue");
      const lookupHelp = document.getElementById("__quickUpdateLookupHelp");
      const loadCurrentBtn = document.getElementById("__quickUpdateLoadCurrent");
      const submitBtn = document.getElementById("__quickUpdateSubmit");
      const clearBtn = document.getElementById("__quickUpdateClear");
      const currentValueBox = document.getElementById("__quickUpdateCurrentValue");
      const payloadPreviewBox = document.getElementById("__quickUpdatePayloadPreview");
      const statusBox = document.getElementById("__quickUpdateStatus");

      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const BASE_URL = `${clientUrl}/api/data/v9.2`;

      function normalizeGuid(value) {
        return String(value || "").replace(/[{}]/g, "").trim();
      }

      function isGuid(value) {
        return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value);
      }

      async function fetchJSON(url, options = {}) {
        const res = await fetch(url, {
          ...options,
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json;odata.include-annotations=*",
            "Content-Type": "application/json; charset=utf-8",
            ...(options.headers || {})
          },
          credentials: "same-origin"
        });

        if (!res.ok) {
          let message = `${res.status}`;
          try {
            const err = await res.json();
            message = err?.error?.message || message;
          } catch (_) {}
          throw new Error(message);
        }

        if (res.status === 204) return null;
        return await res.json();
      }

      async function getEntitySetName(logicalName) {
        const url =
          `${BASE_URL}/EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName`;

        const data = await fetchJSON(url);
        if (!data || !data.EntitySetName) {
          throw new Error(`EntitySetName not found for '${logicalName}'`);
        }

        return data.EntitySetName;
      }

      function tryGetCurrentRecordContext() {
        try {
          if (window.Xrm?.Page?.data?.entity) {
            const entityName = Xrm.Page.data.entity.getEntityName?.();
            const id = normalizeGuid(Xrm.Page.data.entity.getId?.());
            return { entityName, id };
          }
        } catch (_) {}

        try {
          const pageEntity = window.parent?.Xrm?.Page?.data?.entity;
          if (pageEntity) {
            const entityName = pageEntity.getEntityName?.();
            const id = normalizeGuid(pageEntity.getId?.());
            return { entityName, id };
          }
        } catch (_) {}

        return null;
      }

      function updateLookupHelpVisibility() {
        lookupHelp.style.display = typeSelect.value === "lookup" ? "block" : "none";
      }

      function parseInputValue(type, rawValue, fieldName) {
        const value = rawValue ?? "";

        switch (type) {
          case "string":
          case "memo":
            return { [fieldName]: String(value) };

          case "integer": {
            if (String(value).trim() === "") throw new Error("Integer value is required");
            const parsed = parseInt(value, 10);
            if (Number.isNaN(parsed)) throw new Error("Invalid integer value");
            return { [fieldName]: parsed };
          }

          case "decimal":
          case "double": {
            if (String(value).trim() === "") throw new Error("Numeric value is required");
            const parsed = Number(value);
            if (Number.isNaN(parsed)) throw new Error("Invalid numeric value");
            return { [fieldName]: parsed };
          }

          case "boolean": {
            const normalized = String(value).trim().toLowerCase();
            if (["true", "1", "yes"].includes(normalized)) return { [fieldName]: true };
            if (["false", "0", "no"].includes(normalized)) return { [fieldName]: false };
            throw new Error("Boolean must be true/false or 1/0");
          }

          case "datetime": {
            if (String(value).trim() === "") throw new Error("Datetime value is required");
            const dt = new Date(value);
            if (isNaN(dt.getTime())) throw new Error("Invalid datetime value");
            return { [fieldName]: dt.toISOString() };
          }

          case "optionset": {
            if (String(value).trim() === "") throw new Error("Option Set integer value is required");
            const parsed = parseInt(value, 10);
            if (Number.isNaN(parsed)) throw new Error("Invalid Option Set value");
            return { [fieldName]: parsed };
          }

          case "lookup": {
            let parsed;
            try {
              parsed = JSON.parse(value);
            } catch (_) {
              throw new Error('Lookup value must be JSON like {"entitySetName":"systemusers","id":"GUID"}');
            }

            const entitySetName = String(parsed.entitySetName || "").trim();
            const id = normalizeGuid(parsed.id || "");

            if (!entitySetName) throw new Error("Lookup JSON missing entitySetName");
            if (!isGuid(id)) throw new Error("Lookup JSON contains invalid GUID");

            return {
              [`${fieldName}@odata.bind`]: `/${entitySetName}(${id})`
            };
          }

          default:
            throw new Error(`Unsupported field type: ${type}`);
        }
      }

      function refreshPayloadPreview() {
          try {
            const fieldName = fieldInput.value.trim();
            const type = typeSelect.value;
            const rawValue = valueInput.value;

            if (!fieldName) {
              payloadPreviewBox.textContent = "{}";
              return;
            }

            if (String(rawValue).trim() === "") {
              payloadPreviewBox.textContent = "{}";
              return;
            }

            const payload = parseInputValue(type, rawValue, fieldName);
            payloadPreviewBox.textContent = JSON.stringify(payload, null, 2);
          } catch (err) {
            payloadPreviewBox.textContent = `Invalid payload: ${err.message}`;
          }
        }
async function getLookupTargetEntity(entityLogicalName, fieldName) {
  const url =
    `${BASE_URL}/EntityDefinitions(LogicalName='${entityLogicalName}')` +
    `/Attributes(LogicalName='${fieldName}')/Microsoft.Dynamics.CRM.LookupAttributeMetadata` +
    `?$select=LogicalName,Targets`;

  const data = await fetchJSON(url);

  if (Array.isArray(data?.Targets) && data.Targets.length > 0) {
    return data.Targets[0];
  }

  return null;
}
     async function loadCurrentValue() {
  const entityLogicalName = entityInput.value.trim();
  const recordId = normalizeGuid(idInput.value);
  const fieldName = fieldInput.value.trim();
  const fieldType = typeSelect.value;

  if (!entityLogicalName) throw new Error("Entity logical name is required");
  if (!fieldName) throw new Error("Field logical name is required");
  if (!isGuid(recordId)) throw new Error("Valid GUID is required");

  const entitySetName = await getEntitySetName(entityLogicalName);

  let selectFieldName = fieldName;
  let responseFieldName = fieldName;

  if (fieldType === "lookup") {
    selectFieldName = `_${fieldName}_value`;
    responseFieldName = selectFieldName;
  }

  const url =
    `${BASE_URL}/${entitySetName}(${recordId})?$select=${encodeURIComponent(selectFieldName)}`;

  const data = await fetchJSON(url);

  const raw = data[responseFieldName] ?? null;
  const formatted =
    data[`${responseFieldName}@OData.Community.Display.V1.FormattedValue`] ?? null;

  const result = {
    raw,
    formatted
  };

  if (fieldType === "lookup") {
    let logicalName =
      data[`${responseFieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`] ?? null;

    if (!logicalName) {
      try {
        logicalName = await getLookupTargetEntity(entityLogicalName, fieldName);
      } catch (_) {}
    }

    result.logicalName = logicalName ?? null;

    if (raw && logicalName) {
  try {
    const lookupEntitySetName = await getEntitySetName(logicalName);

    valueInput.value = JSON.stringify({
      entitySetName: lookupEntitySetName,
      id: raw
    }, null, 2);

    // 🔥 זה החדש — להביא שם
    const name = await getLookupFormattedValue(logicalName, raw);

    result.formatted = name;
  } catch (_) {}
}
  }

  currentValueBox.textContent = JSON.stringify(result, null, 2);
  refreshPayloadPreview();
}
async function getLookupFormattedValue(entityLogicalName, id) {
  const entitySetName = await getEntitySetName(entityLogicalName);
  const primaryName = await getPrimaryNameField(entityLogicalName);

  const url =
    `${BASE_URL}/${entitySetName}(${id})?$select=${primaryName}`;

  const data = await fetchJSON(url);

  return data[primaryName] ?? null;
}
async function getPrimaryNameField(entityLogicalName) {
  const url =
    `${BASE_URL}/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=PrimaryNameAttribute`;

  const data = await fetchJSON(url);
  return data.PrimaryNameAttribute;
}
      async function submitUpdate() {
        const entityLogicalName = entityInput.value.trim();
        const recordId = normalizeGuid(idInput.value);
        const fieldName = fieldInput.value.trim();
        const type = typeSelect.value;
        const rawValue = valueInput.value;

        if (!entityLogicalName) throw new Error("Entity logical name is required");
        if (!fieldName) throw new Error("Field logical name is required");
        if (!isGuid(recordId)) throw new Error("Valid GUID is required");

        const payload = parseInputValue(type, rawValue, fieldName);
        const entitySetName = await getEntitySetName(entityLogicalName);
        const url = `${BASE_URL}/${entitySetName}(${recordId})`;

        await fetchJSON(url, {
          method: "PATCH",
          body: JSON.stringify(payload),
          headers: {
            "If-Match": "*"
          }
        });

        payloadPreviewBox.textContent = JSON.stringify(payload, null, 2);
      }

      function clearForm() {
        const detected = tryGetCurrentRecordContext();

        fieldInput.value = "";
        typeSelect.value = "string";
        valueInput.value = "";
        currentValueBox.textContent = "Not loaded";
        payloadPreviewBox.textContent = "{}";
        statusBox.textContent = "Ready";

        if (detected?.entityName) entityInput.value = detected.entityName;
        if (detected?.id) idInput.value = detected.id;

        updateLookupHelpVisibility();
        refreshPayloadPreview();
      }

      const detectedContext = tryGetCurrentRecordContext();
      if (detectedContext?.entityName) entityInput.value = detectedContext.entityName;
      if (detectedContext?.id) idInput.value = detectedContext.id;

      updateLookupHelpVisibility();
      refreshPayloadPreview();

      typeSelect.addEventListener("change", () => {
        updateLookupHelpVisibility();
        refreshPayloadPreview();
      });

      fieldInput.addEventListener("input", refreshPayloadPreview);
      valueInput.addEventListener("input", refreshPayloadPreview);

      loadCurrentBtn.addEventListener("click", async () => {
        loadCurrentBtn.disabled = true;
        statusBox.textContent = "Loading current value...";

        try {
          await loadCurrentValue();
          statusBox.textContent = "✅ Current value loaded";
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
        } finally {
          loadCurrentBtn.disabled = false;
        }
      });

      submitBtn.addEventListener("click", async () => {
        submitBtn.disabled = true;
        statusBox.textContent = "Updating...";

        try {
          await submitUpdate();
          statusBox.textContent = "✅ Record updated successfully";

          try {
            await loadCurrentValue();
          } catch (_) {}
        } catch (err) {
          statusBox.textContent = `❌ ${err.message}`;
        } finally {
          submitBtn.disabled = false;
        }
      });

      clearBtn.addEventListener("click", clearForm);
    }
  });
});





document.getElementById("formBeautifier")?.addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();

  if (!tab?.id) {
    alert("No active tab found.");
    return;
  }

  const frameResults = await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: true },
    world: "MAIN",
    func: () => {
      return {
        hasXrm: !!window.Xrm?.Page?.ui,
        href: location.href
      };
    }
  });

  const xrmFrame = frameResults.find(r => r.result?.hasXrm);

  if (!xrmFrame) {
    alert("Could not find Xrm.Page. Open this on a Dynamics 365 form.");
    return;
  }

  const [{ result }] = await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      frameIds: [xrmFrame.frameId]
    },
    world: "MAIN",
    func: () => {
      try {
        const Xrm = window.Xrm;

        if (!Xrm?.Page?.ui?.controls) {
          return {
            ok: false,
            message: "Xrm.Page.ui.controls not found."
          };
        }

        const STYLE_ID = "rv-form-beautifier-style";

        if (!document.getElementById(STYLE_ID)) {
          const style = document.createElement("style");
          style.id = STYLE_ID;
          style.textContent = `
            .rv-beautifier-badge {
              display: inline-block;
              margin-inline-start: 6px;
              padding: 2px 6px;
              border-radius: 6px;
              font-size: 11px;
              font-weight: 600;
              background: #1f6feb;
              color: #fff;
              direction: ltr;
              vertical-align: middle;
            }

            .rv-beautifier-required {
              outline: 2px solid #ff9800 !important;
              outline-offset: 2px !important;
            }

            .rv-beautifier-dirty {
              outline: 2px solid #e91e63 !important;
              outline-offset: 2px !important;
            }

            .rv-beautifier-disabled {
              outline: 2px dashed #9e9e9e !important;
              outline-offset: 2px !important;
              opacity: 0.85;
            }

            .rv-beautifier-hidden-row {
              border: 1px dashed #607d8b !important;
              background: rgba(96, 125, 139, 0.15) !important;
            }

            .rv-beautifier-tooltip {
              cursor: help;
            }
          `;

          document.head.appendChild(style);
        }

        let totalControls = 0;
        let decorated = 0;
        let dirty = 0;
        let required = 0;
        let disabled = 0;
        let hidden = 0;

        Xrm.Page.ui.controls.forEach(control => {
          try {
            const name = control.getName?.();
            const label = control.getLabel?.();
            const attr = control.getAttribute?.();

            if (!name || !attr) return;

            totalControls++;

            const attrName = attr.getName?.() || name;
            const value = attr.getValue?.();
            const type = attr.getAttributeType?.();
            const requiredLevel = attr.getRequiredLevel?.();
            const isDirty = attr.getIsDirty?.();
            const isDisabled = control.getDisabled?.();
            const isVisible = control.getVisible?.();

            if (isDirty) dirty++;
            if (requiredLevel === "required") required++;
            if (isDisabled) disabled++;
            if (isVisible === false) hidden++;

            const controlEl =
              document.querySelector(`[data-id="${name}.fieldControl"]`) ||
              document.querySelector(`[data-id="${name}"]`) ||
              document.querySelector(`[aria-label="${label}"]`);

            const labelEl =
              document.querySelector(`[data-id="${name}.fieldControl-label"]`) ||
              document.querySelector(`label[for*="${name}"]`) ||
              [...document.querySelectorAll("label, span")]
                .find(x => x.textContent?.trim() === label);

            if (labelEl && !labelEl.querySelector(".rv-beautifier-badge")) {
              const badge = document.createElement("span");
              badge.className = "rv-beautifier-badge";
              badge.textContent = attrName;
              badge.title =
                `Logical: ${attrName}\n` +
                `Type: ${type}\n` +
                `Required: ${requiredLevel}\n` +
                `Dirty: ${isDirty}\n` +
                `Disabled: ${isDisabled}\n` +
                `Visible: ${isVisible}\n` +
                `Value: ${JSON.stringify(value)}`;

              labelEl.appendChild(badge);
              labelEl.classList.add("rv-beautifier-tooltip");
              decorated++;
            }

            if (controlEl) {
              controlEl.title =
                `Logical: ${attrName}\n` +
                `Display: ${label || ""}\n` +
                `Type: ${type}\n` +
                `Required: ${requiredLevel}\n` +
                `Dirty: ${isDirty}\n` +
                `Disabled: ${isDisabled}\n` +
                `Visible: ${isVisible}\n` +
                `Value: ${JSON.stringify(value)}`;

              if (requiredLevel === "required") {
                controlEl.classList.add("rv-beautifier-required");
              }

              if (isDirty) {
                controlEl.classList.add("rv-beautifier-dirty");
              }

              if (isDisabled) {
                controlEl.classList.add("rv-beautifier-disabled");
              }

              if (isVisible === false) {
                controlEl.classList.add("rv-beautifier-hidden-row");
              }
            }
          } catch (e) {
            console.warn("Beautifier control error", e);
          }
        });

        return {
          ok: true,
          totalControls,
          decorated,
          dirty,
          required,
          disabled,
          hidden
        };

      } catch (e) {
        return {
          ok: false,
          message: e.message || String(e)
        };
      }
    }
  });

  if (result?.ok) {
    alert(
      `Form Beautifier applied.\n\n` +
      `Controls: ${result.totalControls}\n` +
      `Labels decorated: ${result.decorated}\n` +
      `Required: ${result.required}\n` +
      `Dirty: ${result.dirty}\n` +
      `Disabled: ${result.disabled}\n` +
      `Hidden: ${result.hidden}`
    );
  } else {
    alert(`Form Beautifier failed:\n${result?.message || "Unknown error"}`);
  }
});


document.getElementById("elementInspector")?.addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });

  if (!tab?.id) {
    alert("No active tab found.");
    return;
  }

  const frameResults = await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      allFrames: true
    },
    world: "MAIN",
    func: () => ({
      hasXrm: !!window.Xrm,
      href: location.href
    })
  });

  const targetFrame = frameResults.find(r => r.result?.hasXrm);

  if (!targetFrame) {
    alert("Could not find Dynamics frame.");
    return;
  }

  await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      frameIds: [targetFrame.frameId]
    },
    world: "MAIN",
    func: () => {
      if (window.__rvElementInspectorInstalled) {
        alert("Element Inspector already installed.\n\nHover a ribbon button and press ALT + SHIFT.");
        return;
      }

      window.__rvElementInspectorInstalled = true;

      let currentHoveredRibbonButton = null;

      // --- Deactivate / teardown ---
      function deactivate() {
        removeRibbonHighlight();
        removePopup();

        document.removeEventListener("mouseover", onMouseOver, true);
        document.removeEventListener("keydown",   onKeyDown,   true);

        document.getElementById("rv-element-helper")?.remove();
        document.getElementById("rv-element-helper-style")?.remove();
        document.getElementById("rv-element-inspector-style")?.remove();

        window.__rvElementInspectorInstalled = false;
        window.__rvCachedApplicationRibbonXml   = null;
      }

      showElementInspectorHelper();

      async function retrieveEntityRibbonXml(entityName) {
        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        const url =
          `${clientUrl}/api/data/v9.2/RetrieveEntityRibbon` +
          `(EntityName='${encodeURIComponent(entityName)}',RibbonLocationFilter='All')`;

        const response = await fetch(url, {
          method: "GET",
          headers: {
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
          }
        });

        const text = await response.text();

        if (!response.ok) {
          throw new Error(text);
        }

        const json = JSON.parse(text);

        const raw =
          json.RibbonXml ||
          json.RibbonXmlString ||
          json.EntityRibbonXml ||
          json.CompressedEntityXml ||
          json.value ||
          "";

        if (!raw) {
          throw new Error("Entity Ribbon XML is empty.");
        }

        return await decodeRibbonXml(raw);
      }

      async function retrieveApplicationRibbonXml() {
        if (window.__rvCachedApplicationRibbonXml) {
          return window.__rvCachedApplicationRibbonXml;
        }

        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        const url = `${clientUrl}/api/data/v9.2/RetrieveApplicationRibbon()`;

        const response = await fetch(url, {
          method: "GET",
          headers: {
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
          }
        });

        const text = await response.text();

        if (!response.ok) {
          throw new Error(text);
        }

        const json = JSON.parse(text);

        const raw =
          json.CompressedApplicationRibbonXml ||
          json.ApplicationRibbonXml ||
          json.RibbonXml ||
          json.value ||
          "";

        if (!raw) {
          throw new Error("Application Ribbon XML is empty.");
        }

        const xml = await decodeRibbonXml(raw);

        window.__rvCachedApplicationRibbonXml = xml;

        return xml;
      }

      async function decodeRibbonXml(raw) {
        if (raw.trim().startsWith("<")) {
          return raw;
        }

        const bytes = base64ToBytes(raw);

        if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
          return await extractXmlFromZip(bytes);
        }

        const utf8 = new TextDecoder("utf-8").decode(bytes);

        if (utf8.trim().startsWith("<")) {
          return utf8;
        }

        const utf16 = new TextDecoder("utf-16le").decode(bytes);

        if (utf16.trim().startsWith("<")) {
          return utf16;
        }

        throw new Error("Could not decode Ribbon XML.");
      }

      function base64ToBytes(base64) {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);

        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }

        return bytes;
      }

      async function extractXmlFromZip(bytes) {
        let offset = 0;

        while (offset < bytes.length - 30) {
          const signature =
            bytes[offset] |
            (bytes[offset + 1] << 8) |
            (bytes[offset + 2] << 16) |
            (bytes[offset + 3] << 24);

          if (signature !== 0x04034b50) {
            offset++;
            continue;
          }

          const compressionMethod =
            bytes[offset + 8] |
            (bytes[offset + 9] << 8);

          const compressedSize =
            bytes[offset + 18] |
            (bytes[offset + 19] << 8) |
            (bytes[offset + 20] << 16) |
            (bytes[offset + 21] << 24);

          const fileNameLength =
            bytes[offset + 26] |
            (bytes[offset + 27] << 8);

          const extraLength =
            bytes[offset + 28] |
            (bytes[offset + 29] << 8);

          const fileNameStart = offset + 30;
          const fileNameEnd = fileNameStart + fileNameLength;

          const fileName = new TextDecoder("utf-8").decode(
            bytes.slice(fileNameStart, fileNameEnd)
          );

          const dataStart = fileNameEnd + extraLength;
          const dataEnd = dataStart + compressedSize;
          const fileBytes = bytes.slice(dataStart, dataEnd);

          if (fileName.toLowerCase().endsWith(".xml")) {
            if (compressionMethod === 0) {
              return new TextDecoder("utf-8").decode(fileBytes);
            }

            if (compressionMethod === 8) {
              return await inflateRaw(fileBytes);
            }

            throw new Error("Unsupported ZIP compression method: " + compressionMethod);
          }

          offset = dataEnd;
        }

        throw new Error("No XML file found inside Ribbon ZIP.");
      }

      async function inflateRaw(bytes) {
        for (const format of ["deflate-raw", "deflate"]) {
          try {
            const stream = new Blob([bytes])
              .stream()
              .pipeThrough(new DecompressionStream(format));

            const text = await new Response(stream).text();

            if (text.trim().startsWith("<")) {
              return text;
            }
          } catch {}
        }

        throw new Error("Failed to inflate ZIP XML.");
      }

      function parseXml(xmlText) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, "text/xml");

        const parserError = xmlDoc.getElementsByTagName("parsererror")[0];

        if (parserError) {
          throw new Error(
            "Failed to parse Ribbon XML.\n\nFirst chars:\n" +
            xmlText.slice(0, 700)
          );
        }

        return xmlDoc;
      }

      function findExactById(xmlDoc, tagName, id) {
        if (!xmlDoc || !id) return null;

        return [...xmlDoc.getElementsByTagName(tagName)]
          .find(x => x.getAttribute("Id") === id);
      }

      function findContainsById(xmlDoc, tagName, id) {
        if (!xmlDoc || !id) return null;

        return [...xmlDoc.getElementsByTagName(tagName)]
          .find(x => {
            const xmlId = x.getAttribute("Id") || "";
            return xmlId === id || xmlId.includes(id) || id.includes(xmlId);
          });
      }

      function getNodeXml(node) {
        if (!node) return "";
        return new XMLSerializer().serializeToString(node);
      }

      function getRuleIds(commandDef, containerName) {
        if (!commandDef) return [];

        const container = commandDef.getElementsByTagName(containerName)?.[0];

        if (!container) return [];

        const childTag =
          containerName === "EnableRules"
            ? "EnableRule"
            : "DisplayRule";

        return [...container.getElementsByTagName(childTag)]
          .map(x => x.getAttribute("Id"))
          .filter(Boolean);
      }

      function getRuleDetails(xmlDocs, ruleIds, ruleTagName) {
        const docs = xmlDocs.filter(Boolean);

        return ruleIds.map(id => {
          let rule = null;
          let source = "";

          for (const docInfo of docs) {
            const allRules = [...docInfo.doc.getElementsByTagName(ruleTagName)];

            rule = allRules.find(x => {
              const ruleId = x.getAttribute("Id") || "";
              const hasChildren = x.children && x.children.length > 0;

              return ruleId === id && hasChildren;
            });

            if (rule) {
              source = docInfo.name;
              break;
            }
          }

          return {
            id,
            found: !!rule,
            source,
            xml: rule ? getNodeXml(rule) : "Rule definition not found"
          };
        });
      }

      function getJavaScriptActions(commandDef) {
        if (!commandDef) return [];

        return [...commandDef.getElementsByTagName("JavaScriptFunction")]
          .map(x => ({
            library:
              x.getAttribute("Library") ||
              x.getAttribute("library") ||
              "",
            functionName:
              x.getAttribute("FunctionName") ||
              x.getAttribute("Function") ||
              "",
            parameters: [...x.children].map(p => ({
              type: p.tagName,
              value:
                p.getAttribute("Value") ||
                p.getAttribute("value") ||
                p.getAttribute("Name") ||
                p.getAttribute("name") ||
                p.getAttribute("Parameter") ||
                ""
            }))
          }));
      }

      async function resolveRibbon(entityXmlText, clickedInfo) {
        const entityXmlDoc = entityXmlText ? parseXml(entityXmlText) : null;

        let appXmlDoc = null;
        let appXmlError = "";

        try {
          const appXmlText = await retrieveApplicationRibbonXml();
          appXmlDoc = parseXml(appXmlText);
        } catch (e) {
          appXmlError = e.message || String(e);
        }

        const buttonId = clickedInfo.buttonId;

        const button =
          findExactById(entityXmlDoc, "Button", buttonId) ||
          findContainsById(entityXmlDoc, "Button", buttonId) ||
          findExactById(entityXmlDoc, "FlyoutAnchor", buttonId) ||
          findContainsById(entityXmlDoc, "FlyoutAnchor", buttonId) ||
          findExactById(entityXmlDoc, "SplitButton", buttonId) ||
          findContainsById(entityXmlDoc, "SplitButton", buttonId) ||
          findExactById(entityXmlDoc, "MenuItem", buttonId) ||
          findContainsById(entityXmlDoc, "MenuItem", buttonId) ||
          findExactById(appXmlDoc, "Button", buttonId) ||
          findContainsById(appXmlDoc, "Button", buttonId) ||
          findExactById(appXmlDoc, "FlyoutAnchor", buttonId) ||
          findContainsById(appXmlDoc, "FlyoutAnchor", buttonId) ||
          findExactById(appXmlDoc, "SplitButton", buttonId) ||
          findContainsById(appXmlDoc, "SplitButton", buttonId) ||
          findExactById(appXmlDoc, "MenuItem", buttonId) ||
          findContainsById(appXmlDoc, "MenuItem", buttonId);

        const commandId =
          button?.getAttribute("Command") ||
          clickedInfo.buttonId ||
          "";

        const commandDef =
          findExactById(entityXmlDoc, "CommandDefinition", commandId) ||
          findContainsById(entityXmlDoc, "CommandDefinition", commandId) ||
          findExactById(appXmlDoc, "CommandDefinition", commandId) ||
          findContainsById(appXmlDoc, "CommandDefinition", commandId);

        const enableRuleIds  = getRuleIds(commandDef, "EnableRules");
        const displayRuleIds = getRuleIds(commandDef, "DisplayRules");

        const searchDocs = [
          entityXmlDoc ? { name: "Entity Ribbon",      doc: entityXmlDoc } : null,
          appXmlDoc    ? { name: "Application Ribbon",  doc: appXmlDoc    } : null
        ];

        const enableRules  = getRuleDetails(searchDocs, enableRuleIds,  "EnableRule");
        const displayRules = getRuleDetails(searchDocs, displayRuleIds, "DisplayRule");

        return {
          clicked: clickedInfo,
          buttonId,
          commandId,
          buttonFound:      !!button,
          commandFound:     !!commandDef,
          appRibbonLoaded:  !!appXmlDoc,
          appRibbonError:   appXmlError,
          buttonXml:        getNodeXml(button),
          commandXml:       getNodeXml(commandDef),
          jsActions:        getJavaScriptActions(commandDef),
          enableRules,
          displayRules
        };
      }

      function findRealRibbonButton(startElement) {
        const first = startElement.closest(
          "button,[role='button'],[role='menuitem'],[data-id]"
        );

        if (!first) return null;

        const candidates = [
          first,
          first.closest("[data-id]"),
          first.parentElement?.closest("[data-id]"),
          first.parentElement,
          first.parentElement?.parentElement?.closest("[data-id]")
        ].filter(Boolean);

        return (
          candidates.find(x => {
            const id = x.getAttribute?.("data-id") || x.id || "";
            return id.includes("|");
          }) ||
          candidates.find(x => {
            const id = x.getAttribute?.("data-id") || x.id || "";
            return id.length > 0;
          }) ||
          first
        );
      }

      function getClickedInfo(btn) {
        const dataId = btn.getAttribute("data-id") || btn.id || "";
        const parts  = dataId.split("|");

        const urlEntity = new URLSearchParams(location.search).get("etn") || "";
        const hasPipes  = parts.length >= 4;

        return {
          text:
            btn.innerText?.trim() ||
            btn.getAttribute("aria-label") ||
            btn.getAttribute("title") ||
            "",
          dataId,
          elementId:    btn.id || "",
          entity:       hasPipes ? parts[0] : urlEntity,
          relationship: hasPipes ? parts[1] : "",
          location:     hasPipes ? parts[2] : "",
          buttonId:     hasPipes ? parts.slice(3).join("|") : dataId,
          isDashboard:  !hasPipes
        };
      }

      function highlightRibbonButton(btn) {
        removeRibbonHighlight();

        if (!btn) return;

        btn.style.outline       = "3px solid #6aa9ff";
        btn.style.outlineOffset = "2px";

        currentHoveredRibbonButton = btn;
      }

      function removeRibbonHighlight() {
        if (currentHoveredRibbonButton) {
          currentHoveredRibbonButton.style.outline       = "";
          currentHoveredRibbonButton.style.outlineOffset = "";
        }
      }

      // Named functions so removeEventListener can target them exactly
      function onMouseOver(e) {
        const btn    = findRealRibbonButton(e.target);
        if (!btn) return;

        const dataId = btn.getAttribute("data-id") || btn.id || "";
        if (!dataId) return;

        highlightRibbonButton(btn);
      }

      async function onKeyDown(e) {
        if (!(e.altKey && e.shiftKey)) return;

        if (!currentHoveredRibbonButton) {
          alert("Hover a ribbon button first, then press ALT + SHIFT.");
          return;
        }

        e.preventDefault();
        e.stopPropagation();

        const clickedInfo = getClickedInfo(currentHoveredRibbonButton);

        if (!clickedInfo.buttonId) {
          alert("Could not resolve ribbon button ID.");
          return;
        }

        showLoadingPopup(clickedInfo);

        try {
          let entityXml = "";
          if (clickedInfo.entity) {
            entityXml = await retrieveEntityRibbonXml(clickedInfo.entity);
          }

          const resolved = await resolveRibbon(entityXml, clickedInfo);
          showPopup(resolved);
        } catch (err) {
          removePopup();

          alert(
            "Failed to resolve ribbon.\n\n" +
            (err.message || String(err))
          );
        }
      }

      document.addEventListener("mouseover", onMouseOver, true);
      document.addEventListener("keydown",   onKeyDown,   true);

      function escapeHtml(value) {
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
      }

      function removePopup() {
        document.getElementById("rv-element-inspector-popup")?.remove();
      }

      function injectStyle() {
        if (document.getElementById("rv-element-inspector-style")) return;

        const style = document.createElement("style");
        style.id = "rv-element-inspector-style";

        style.textContent = `
          #rv-element-inspector-popup {
            position: fixed;
            top: 16px;
            left: 16px;
            width: 780px;
            max-width: calc(100vw - 32px);
            max-height: calc(100vh - 32px);
            overflow-y: auto;
            background: #1f1f1f;
            color: white;
            z-index: 999999999;
            border: 1px solid #444;
            border-radius: 14px;
            box-shadow: 0 12px 40px rgba(0,0,0,.55);
            font-family: Arial, sans-serif;
            direction: rtl;
          }

          .rv-head {
            position: sticky;
            top: 0;
            background: #1f1f1f;
            padding: 14px 16px;
            border-bottom: 1px solid #333;
            display: flex;
            justify-content: space-between;
            align-items: center;
            z-index: 2;
          }

          .rv-head h2 {
            margin: 0;
            font-size: 18px;
          }

          .rv-close {
            background: transparent;
            border: 0;
            color: white;
            font-size: 26px;
            cursor: pointer;
          }

          .rv-body {
            padding: 14px 16px;
          }

          .rv-section {
            color: #6aa9ff;
            font-weight: bold;
            margin: 18px 0 8px;
            font-size: 15px;
          }

          .rv-label {
            color: #8cc8ff;
            direction: rtl;
            text-align: right;
            font-size: 12px;
            margin: 8px 0 4px;
          }

          .rv-box {
            background: #2a2a2a;
            border: 1px solid #333;
            border-radius: 8px;
            padding: 10px;
            margin-bottom: 10px;
            direction: ltr;
            text-align: left;
            white-space: pre-wrap;
            word-break: break-word;
            font-size: 12px;
          }

          .rv-empty {
            color: #aaa;
          }

          .rv-ok {
            color: #4caf50;
            font-weight: bold;
          }

          .rv-bad {
            color: #ff6b6b;
            font-weight: bold;
          }

          .rv-warn {
            color: #ffd36a;
            font-weight: bold;
          }

          .rv-source {
            color: #ffd36a;
            font-size: 11px;
            margin-inline-start: 8px;
          }

          .rv-actions {
            position: sticky;
            bottom: 0;
            background: #1f1f1f;
            border-top: 1px solid #333;
            display: flex;
            gap: 8px;
            padding: 12px 16px;
          }

          .rv-btn {
            flex: 1;
            background: #6aa9ff;
            color: #000;
            border: 0;
            border-radius: 8px;
            padding: 10px;
            font-weight: bold;
            cursor: pointer;
          }

          .rv-btn-secondary {
            background: #444;
            color: white;
          }

          .rv-btn-danger {
            background: #ff6b6b;
            color: #000;
          }
        `;

        document.head.appendChild(style);
      }

      function showLoadingPopup(clickedInfo) {
        removePopup();
        injectStyle();

        const popup = document.createElement("div");
        popup.id = "rv-element-inspector-popup";

        const entityLine = clickedInfo.entity
          ? `Reading Ribbon XML for entity: ${escapeHtml(clickedInfo.entity)}`
          : `No entity context (dashboard) — searching Application Ribbon only`;

        popup.innerHTML = `
          <div class="rv-head">
            <h2>🔍 Element Inspector</h2>
            <button id="rv-element-inspector-close" class="rv-close">×</button>
          </div>

          <div class="rv-body">
            <div class="rv-section">Loading...</div>
            <div class="rv-box">${entityLine}</div>
            <div class="rv-box">Also loading Application Ribbon rules...</div>
          </div>
        `;

        document.body.appendChild(popup);

        document.getElementById("rv-element-inspector-close").onclick = removePopup;
      }

      function renderJsActions(actions) {
        if (!actions.length) {
          return `<div class="rv-box rv-empty">No JavaScriptFunction found.</div>`;
        }

        return actions.map(a => `
          <div class="rv-box">
Library: ${escapeHtml(a.library)}
Function: ${escapeHtml(a.functionName)}
Parameters:
${escapeHtml(JSON.stringify(a.parameters, null, 2))}
          </div>
        `).join("");
      }

      function renderRules(rules) {
        if (!rules.length) {
          return `<div class="rv-box rv-empty">No rules found.</div>`;
        }

        return rules.map(r => `
          <div class="rv-label">
            ${escapeHtml(r.id)}
            ${r.found ? "<span class='rv-ok'> found</span>" : "<span class='rv-bad'> not found</span>"}
            ${r.source ? `<span class="rv-source">${escapeHtml(r.source)}</span>` : ""}
          </div>
          <div class="rv-box">${escapeHtml(r.xml || "Rule XML not found")}</div>
        `).join("");
      }

      function showPopup(data) {
        removePopup();
        injectStyle();

        const popup = document.createElement("div");
        popup.id = "rv-element-inspector-popup";

        const dashboardNote = data.clicked.isDashboard
          ? `<div class="rv-label">Context</div>
             <div class="rv-box"><span class="rv-warn">Dashboard</span> — no entity XML loaded. Results are from Application Ribbon only.</div>`
          : "";

        popup.innerHTML = `
          <div class="rv-head">
            <h2>🔍 Element Inspector</h2>
            <button id="rv-element-inspector-close" class="rv-close">×</button>
          </div>

          <div class="rv-body">
            <div class="rv-section">Clicked Button</div>

            ${dashboardNote}

            <div class="rv-label">Text</div>
            <div class="rv-box">${escapeHtml(data.clicked.text)}</div>

            <div class="rv-label">Entity</div>
            <div class="rv-box">${escapeHtml(data.clicked.entity) || '<span class="rv-empty">none (dashboard)</span>'}</div>

            <div class="rv-label">Button Id</div>
            <div class="rv-box">${escapeHtml(data.buttonId)}</div>

            <div class="rv-label">Command Id</div>
            <div class="rv-box">${escapeHtml(data.commandId)}</div>

            <div class="rv-label">Button Found</div>
            <div class="rv-box">${data.buttonFound ? "<span class='rv-ok'>true</span>" : "<span class='rv-bad'>false</span>"}</div>

            <div class="rv-label">Command Found</div>
            <div class="rv-box">${data.commandFound ? "<span class='rv-ok'>true</span>" : "<span class='rv-bad'>false</span>"}</div>

            <div class="rv-label">Application Ribbon Loaded</div>
            <div class="rv-box">${data.appRibbonLoaded ? "<span class='rv-ok'>true</span>" : "<span class='rv-bad'>false</span>"}</div>

            ${
              data.appRibbonError
                ? `
                  <div class="rv-label">Application Ribbon Error</div>
                  <div class="rv-box">${escapeHtml(data.appRibbonError)}</div>
                `
                : ""
            }

            <div class="rv-section">JavaScript Functions</div>
            ${renderJsActions(data.jsActions)}

            <div class="rv-section">Enable Rules</div>
            ${renderRules(data.enableRules)}

            <div class="rv-section">Display Rules</div>
            ${renderRules(data.displayRules)}

            <div class="rv-section">Command XML</div>
            <div class="rv-box">${escapeHtml(data.commandXml || "Not found")}</div>

            <div class="rv-section">Button XML</div>
            <div class="rv-box">${escapeHtml(data.buttonXml || "Not found")}</div>
          </div>

          <div class="rv-actions">
            <button id="rv-element-copy-json"    class="rv-btn">Copy JSON</button>
            <button id="rv-element-copy-command" class="rv-btn">Copy Command</button>
            <button id="rv-element-close2"       class="rv-btn rv-btn-secondary">Close</button>
          </div>
        `;

        document.body.appendChild(popup);

        document.getElementById("rv-element-inspector-close").onclick = removePopup;
        document.getElementById("rv-element-close2").onclick          = removePopup;

        document.getElementById("rv-element-copy-json").onclick = async () => {
          await navigator.clipboard.writeText(JSON.stringify(data, null, 2));
          alert("Copied JSON");
        };

        document.getElementById("rv-element-copy-command").onclick = async () => {
          await navigator.clipboard.writeText(data.commandId || data.buttonId || "");
          alert("Copied Command");
        };
      }

      function showElementInspectorHelper() {
        document.getElementById("rv-element-helper")?.remove();
        document.getElementById("rv-element-helper-style")?.remove();

        const style = document.createElement("style");
        style.id = "rv-element-helper-style";

        style.textContent = `
          #rv-element-helper {
            position: fixed;
            bottom: 20px;
            left: 20px;
            z-index: 999999999;
            background: #1f1f1f;
            color: white;
            border: 1px solid #444;
            border-radius: 12px;
            padding: 12px 14px;
            font-family: Arial, sans-serif;
            box-shadow: 0 10px 30px rgba(0,0,0,.45);
            direction: ltr;
            user-select: none;
            transition: all .2s ease;
          }

          #rv-element-helper:hover {
            opacity: 1 !important;
            transform: scale(1.03);
          }

          .rv-element-helper-title {
            font-weight: bold;
            margin-bottom: 6px;
            color: #6aa9ff;
            font-size: 14px;
          }

          .rv-element-helper-text {
            font-size: 12px;
            color: #ccc;
            line-height: 1.5;
          }

          .rv-element-helper-hotkey {
            margin-top: 10px;
            background: #6aa9ff;
            color: black;
            font-weight: bold;
            text-align: center;
            padding: 8px;
            border-radius: 8px;
            font-size: 14px;
          }

          #rv-deactivate-btn {
            margin-top: 10px;
            width: 100%;
            background: #ff6b6b;
            color: #000;
            border: 0;
            border-radius: 8px;
            padding: 7px 0;
            font-weight: bold;
            font-size: 13px;
            cursor: pointer;
            font-family: Arial, sans-serif;
          }

          #rv-deactivate-btn:hover {
            background: #ff4444;
          }
        `;

        document.head.appendChild(style);

        const helper = document.createElement("div");
        helper.id = "rv-element-helper";

        helper.innerHTML = `
          <div class="rv-element-helper-title">🔍 Element Inspector</div>
          <div class="rv-element-helper-text">
            Hover ribbon button<br>
            then press
          </div>
          <div class="rv-element-helper-hotkey">ALT + SHIFT</div>
          <button id="rv-deactivate-btn">⏹ Deactivate</button>
        `;

        document.body.appendChild(helper);

        document.getElementById("rv-deactivate-btn").onclick = deactivate;

        setTimeout(() => {
          helper.style.opacity = "0.15";
        }, 10000);

        helper.addEventListener("mouseenter", () => {
          helper.style.opacity = "1";
        });

        helper.addEventListener("mouseleave", () => {
          helper.style.opacity = "0.15";
        });
      }
    }
  });
});




document.getElementById("openRecordByGuidUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      document.getElementById("__d365_guid_resolver_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365_guid_resolver_modal";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,.35);
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
        direction: ltr;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(760px, 96vw);
        background: #fff;
        border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `
        padding: 14px 16px;
        font-weight: 800;
        border-bottom: 1px solid #e5e7eb;
        font-size: 15px;
      `;
      header.textContent = "Open Record By Any GUID";

      const body = document.createElement("div");
      body.style.cssText = `padding: 16px;`;

      const input = document.createElement("input");
      input.placeholder = "Paste GUID here...";
      input.style.cssText = `
        width: 100%;
        box-sizing: border-box;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 12px;
        font-size: 14px;
        direction: ltr;
        text-align: left;
        font-family: Consolas, Monaco, 'Courier New', monospace;
      `;

      const result = document.createElement("div");
      result.style.cssText = `
        margin-top: 12px;
        min-height: 120px;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        background: #f8fafc;
        padding: 12px;
        white-space: pre-wrap;
        font-family: Consolas, Monaco, 'Courier New', monospace;
        font-size: 12px;
        color: #0f172a;
      `;
      result.textContent = "Paste GUID and click Resolve.";

      const footer = document.createElement("div");
      footer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding: 12px 16px;
        border-top: 1px solid #e5e7eb;
      `;

      const makeBtn = (text, bg, color = "#fff") => {
        const btn = document.createElement("button");
        btn.textContent = text;
        btn.style.cssText = `
          border: ${bg === "#fff" ? "1px solid #cbd5e1" : "none"};
          padding: 10px 14px;
          border-radius: 10px;
          cursor: pointer;
          background: ${bg};
          color: ${color};
          font-weight: 800;
        `;
        return btn;
      };

      const btnClose = makeBtn("Close", "#fff", "#111827");
      const btnResolve = makeBtn("Resolve", "#2563eb");
      const btnOpen = makeBtn("Open Record", "#16a34a");
      btnOpen.style.display = "none";

      let found = null;

      const cleanGuid = (value) => {
        const m = String(value || "").match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
        return m ? m[0].toLowerCase() : null;
      };

      async function getJson(url) {
        const res = await fetch(url, {
          method: "GET",
          headers: {
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": 'odata.include-annotations="*"'
          }
        });

        if (!res.ok) return null;
        return await res.json();
      }

      async function resolveGuid(guid) {
        const Xrm = window.Xrm;
        const clientUrl = Xrm?.Utility?.getGlobalContext?.().getClientUrl?.();

        if (!clientUrl) {
          throw new Error("Xrm context not found. Open this inside D365.");
        }

        result.textContent = "Loading entity metadata...";

        const metaUrl =
          `${clientUrl}/api/data/v9.2/EntityDefinitions` +
          `?$select=LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute` +
          `&$filter=IsPrivate eq false`;

        const meta = await getJson(metaUrl);
        const entities = (meta?.value || [])
          .filter(e => e.EntitySetName && e.PrimaryIdAttribute)
          .sort((a, b) => a.LogicalName.localeCompare(b.LogicalName));

        result.textContent = `Scanning ${entities.length} entities...\nGUID: ${guid}`;

        const concurrency = 12;
        let index = 0;
        let checked = 0;

        async function worker() {
          while (index < entities.length && !found) {
            const e = entities[index++];
            checked++;

            if (checked % 25 === 0) {
              result.textContent = `Scanning entities...\nChecked: ${checked}/${entities.length}\nGUID: ${guid}`;
            }

            const select = e.PrimaryNameAttribute
              ? `${e.PrimaryIdAttribute},${e.PrimaryNameAttribute}`
              : `${e.PrimaryIdAttribute}`;

            const url = `${clientUrl}/api/data/v9.2/${e.EntitySetName}(${guid})?$select=${select}`;
            const data = await getJson(url);

            if (data && data[e.PrimaryIdAttribute]) {
              found = {
                guid,
                entityName: e.LogicalName,
                entitySetName: e.EntitySetName,
                primaryId: e.PrimaryIdAttribute,
                primaryNameAttr: e.PrimaryNameAttribute,
                primaryName: e.PrimaryNameAttribute ? data[e.PrimaryNameAttribute] : ""
              };
              return;
            }
          }
        }

        await Promise.all(Array.from({ length: concurrency }, worker));

        return found;
      }

      btnResolve.onclick = async () => {
        try {
          found = null;
          btnOpen.style.display = "none";

          const guid = cleanGuid(input.value);

          if (!guid) {
            result.textContent = "Invalid GUID.";
            return;
          }

          btnResolve.disabled = true;
          btnResolve.textContent = "Resolving...";

          const r = await resolveGuid(guid);

          if (!r) {
            result.textContent = `Not found.\nGUID: ${guid}`;
            return;
          }

          result.textContent =
            `FOUND ✅\n\n` +
            `Entity: ${r.entityName}\n` +
            `EntitySet: ${r.entitySetName}\n` +
            `Primary ID: ${r.primaryId}\n` +
            `Primary Name Field: ${r.primaryNameAttr || "-"}\n` +
            `Name: ${r.primaryName || "-"}\n` +
            `GUID: ${r.guid}`;

          btnOpen.style.display = "";
        } catch (e) {
          result.textContent = "Error:\n" + (e?.message || String(e));
        } finally {
          btnResolve.disabled = false;
          btnResolve.textContent = "Resolve";
        }
      };

      btnOpen.onclick = async () => {
        if (!found) return;

        await window.Xrm.Navigation.openForm({
          entityName: found.entityName,
          entityId: found.guid
        });
      };

      btnClose.onclick = () => overlay.remove();

      footer.appendChild(btnClose);
      footer.appendChild(btnResolve);
      footer.appendChild(btnOpen);

      body.appendChild(input);
      body.appendChild(result);

      box.appendChild(header);
      box.appendChild(body);
      box.appendChild(footer);
      overlay.appendChild(box);
      document.body.appendChild(overlay);

      input.focus();
    }
  });
});

document.getElementById("whyDidThisHappenUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: true },
    world: "MAIN",
    func: () => {
      const store = window.top.__whyRecorderStore ||= {
        logs: [],
        recording: false
      };

      function safe(v) {
        try {
          if (v === undefined) return "undefined";
          if (v === null) return "null";
          if (typeof v === "string") return v.slice(0, 1500);
          if (typeof v === "number" || typeof v === "boolean") return String(v);
          if (v instanceof Date) return v.toISOString();
          if (Array.isArray(v)) return `[Array ${v.length}]`;
          if (typeof v === "function") return `[Function ${v.name || "anonymous"}]`;
          if (typeof v === "object") return JSON.stringify(v).slice(0, 2000);
          return String(v).slice(0, 1500);
        } catch {
          return "[unserializable]";
        }
      }

      function stack() {
        return (new Error().stack || "")
          .split("\n")
          .slice(2, 22)
          .join("\n");
      }

      function shouldIgnore(type, data = {}) {
        const ignoredPatterns = [
          "events.data.microsoft.com",
          "aria.microsoft.com",
          "OneCollector",
          "Collector/3.0",
          "browser.pipe",
          "measure.office.com",
          "upload.fp.measure.office.com",
          "setImmediate$",
          "trans.gif",
          "__whyRecorder"
        ];

        const combined = `${type} ${JSON.stringify(data || {})}`;
        return ignoredPatterns.some((x) => combined.includes(x));
      }

      function addLog(type, data = {}) {
        if (!store.recording) return;
        if (shouldIgnore(type, data)) return;

        store.logs.push({
          time: new Date().toLocaleTimeString(),
          frameUrl: location.href,
          type,
          ...data,
          stack: stack()
        });

        if (store.logs.length > 2000) store.logs.shift();

        try {
          window.top.__whyRecorderRender?.();
        } catch {}
      }

      function wrap(obj, method, label, parser) {
        if (!obj || typeof obj[method] !== "function") return;
        if (obj[method].__whyWrapped) return;

        const original = obj[method];

        obj[method] = function (...args) {
          try {
            addLog(label, parser ? parser(args, this) : {});
          } catch {}

          return original.apply(this, args);
        };

        obj[method].__whyWrapped = true;
        obj[method].__whyOriginal = original;
      }

      function installXrmHooks() {
        const Xrm = window.Xrm || window.parent?.Xrm || window.top?.Xrm;
        const page = window.Xrm?.Page || window.parent?.Xrm?.Page || window.top?.Xrm?.Page;

        try {
          if (page?.data?.entity?.attributes?.forEach) {
            page.data.entity.attributes.forEach((attr) => {
              wrap(attr, "setValue", "setValue", (args, ctx) => ({
                attribute: ctx.getName?.(),
                newValue: safe(args[0])
              }));

              wrap(attr, "fireOnChange", "fireOnChange", (args, ctx) => ({
                attribute: ctx.getName?.()
              }));
            });
          }
        } catch {}

        try {
          if (page?.ui?.controls?.forEach) {
            page.ui.controls.forEach((ctrl) => {
              wrap(ctrl, "setVisible", "setVisible", (args, ctx) => ({
                control: ctx.getName?.(),
                visible: args[0]
              }));

              wrap(ctrl, "setDisabled", "setDisabled", (args, ctx) => ({
                control: ctx.getName?.(),
                disabled: args[0]
              }));

              wrap(ctrl, "setNotification", "setNotification", (args, ctx) => ({
                control: ctx.getName?.(),
                message: args[0],
                uniqueId: args[1]
              }));

              wrap(ctrl, "clearNotification", "clearNotification", (args, ctx) => ({
                control: ctx.getName?.(),
                uniqueId: args[0]
              }));
            });
          }
        } catch {}

        try {
          if (Xrm?.WebApi) {
            wrap(Xrm.WebApi, "retrieveRecord", "retrieveRecord", (args) => ({
              entity: args[0],
              id: args[1],
              options: safe(args[2])
            }));

            wrap(Xrm.WebApi, "retrieveMultipleRecords", "retrieveMultipleRecords", (args) => ({
              entity: args[0],
              options: safe(args[1])
            }));

            wrap(Xrm.WebApi, "createRecord", "createRecord", (args) => ({
              entity: args[0],
              data: safe(args[1])
            }));

            wrap(Xrm.WebApi, "updateRecord", "updateRecord", (args) => ({
              entity: args[0],
              id: args[1],
              data: safe(args[2])
            }));

            wrap(Xrm.WebApi, "deleteRecord", "deleteRecord", (args) => ({
              entity: args[0],
              id: args[1]
            }));

            wrap(Xrm.WebApi.online || Xrm.WebApi, "execute", "execute", (args) => ({
              request: safe(args[0])
            }));
          }
        } catch {}

        try {
          if (Xrm?.Navigation) {
            wrap(Xrm.Navigation, "openForm", "openForm", (args) => ({
              options: safe(args[0]),
              params: safe(args[1])
            }));

            wrap(Xrm.Navigation, "navigateTo", "navigateTo", (args) => ({
              pageInput: safe(args[0]),
              navigationOptions: safe(args[1])
            }));

            wrap(Xrm.Navigation, "openAlertDialog", "openAlertDialog", (args) => ({
              alertStrings: safe(args[0]),
              alertOptions: safe(args[1])
            }));

            wrap(Xrm.Navigation, "openConfirmDialog", "openConfirmDialog", (args) => ({
              confirmStrings: safe(args[0]),
              confirmOptions: safe(args[1])
            }));
          }
        } catch {}

        try {
          if (page?.data) {
            wrap(page.data, "save", "formSave", (args) => ({
              saveOptions: safe(args[0])
            }));

            wrap(page.data, "refresh", "formRefresh", (args) => ({
              save: safe(args[0])
            }));
          }
        } catch {}

        try {
          if (page?.data?.entity?.addOnSave && !window.__whyOnSaveHookInstalled) {
            window.__whyOnSaveHookInstalled = true;

            page.data.entity.addOnSave((executionContext) => {
              const args = executionContext.getEventArgs?.();

              addLog("OnSave", {
                saveMode: args?.getSaveMode?.(),
                prevented: args?.isDefaultPrevented?.()
              });
            });
          }
        } catch {}
      }

      function installHtmlHooks() {
        if (window.__whyHtmlHooksInstalled) return;
        window.__whyHtmlHooksInstalled = true;

        document.addEventListener(
          "click",
          (e) => {
            const el = e.target?.closest?.(
              "button, a, input, select, textarea, div, span, [onclick], [role='button'], [data-id], [aria-label]"
            );

            if (!el) return;

            if (
              el.id?.startsWith("__whyRecorder") ||
              el.closest?.("#__whyRecorderModal")
            ) {
              return;
            }

            addLog("HTML CLICK", {
              tag: el.tagName,
              id: el.id || "",
              name: el.getAttribute?.("name") || "",
              typeAttr: el.getAttribute?.("type") || "",
              dataId: el.getAttribute?.("data-id") || "",
              aria: el.getAttribute?.("aria-label") || "",
              text: String(el.innerText || el.value || "").trim().slice(0, 500),
              onclick: String(el.getAttribute?.("onclick") || "").slice(0, 800),
              inlineCode: String(el.getAttribute?.("onclick") || "").slice(0, 800)
            });
          },
          true
        );

        try {
          const originalAddEventListener = EventTarget.prototype.addEventListener;

          if (!originalAddEventListener.__whyWrapped) {
            EventTarget.prototype.addEventListener = function (type, handler, options) {
              if (
                typeof handler === "function" &&
                ["click", "change", "input", "submit", "keydown"].includes(type)
              ) {
                const wrappedHandler = function (...args) {
                  addLog("HTML EVENT HANDLER", {
                    eventType: type,
                    handlerName: handler.name || "anonymous",
                    targetTag: this?.tagName || "",
                    targetId: this?.id || "",
                    targetName: this?.getAttribute?.("name") || "",
                    targetClass: String(this?.className || "").slice(0, 300),
                    targetText: String(this?.innerText || this?.value || "").slice(0, 500)
                  });

                  return handler.apply(this, args);
                };

                return originalAddEventListener.call(this, type, wrappedHandler, options);
              }

              return originalAddEventListener.call(this, type, handler, options);
            };

            EventTarget.prototype.addEventListener.__whyWrapped = true;
          }
        } catch {}

        try {
          const elements = document.querySelectorAll("[onclick], [onchange], button, input, select, textarea, a");

          elements.forEach((el) => {
            if (el.__whyInlineWrapped) return;
            el.__whyInlineWrapped = true;

            if (typeof el.onclick === "function") {
              const originalOnClick = el.onclick;

              el.onclick = function (...args) {
                addLog("HTML INLINE onclick", {
                  tag: el.tagName,
                  id: el.id || "",
                  text: String(el.innerText || el.value || "").trim().slice(0, 500),
                  handlerName: originalOnClick.name || "anonymous",
                  inlineCode: String(el.getAttribute?.("onclick") || "").slice(0, 800)
                });

                return originalOnClick.apply(this, args);
              };
            }

            if (typeof el.onchange === "function") {
              const originalOnChange = el.onchange;

              el.onchange = function (...args) {
                addLog("HTML INLINE onchange", {
                  tag: el.tagName,
                  id: el.id || "",
                  value: safe(el.value),
                  handlerName: originalOnChange.name || "anonymous",
                  inlineCode: String(el.getAttribute?.("onchange") || "").slice(0, 800)
                });

                return originalOnChange.apply(this, args);
              };
            }
          });
        } catch {}

        try {
          if (typeof window.fetch === "function" && !window.fetch.__whyWrapped) {
            const originalFetch = window.fetch;

            window.fetch = function (...args) {
              addLog("HTML fetch", {
                url: String(args[0]).slice(0, 1500),
                options: safe(args[1])
              });

              return originalFetch.apply(this, args);
            };

            window.fetch.__whyWrapped = true;
          }
        } catch {}

        try {
          if (
            XMLHttpRequest?.prototype?.open &&
            !XMLHttpRequest.prototype.open.__whyWrapped
          ) {
            const originalOpen = XMLHttpRequest.prototype.open;
            const originalSend = XMLHttpRequest.prototype.send;

            XMLHttpRequest.prototype.open = function (method, url, ...rest) {
              this.__whyXhrInfo = {
                method,
                url: String(url).slice(0, 1500)
              };

              return originalOpen.call(this, method, url, ...rest);
            };

            XMLHttpRequest.prototype.open.__whyWrapped = true;

            XMLHttpRequest.prototype.send = function (body) {
              addLog("HTML XMLHttpRequest", {
                method: this.__whyXhrInfo?.method,
                url: this.__whyXhrInfo?.url,
                body: safe(body)
              });

              return originalSend.call(this, body);
            };

            XMLHttpRequest.prototype.send.__whyWrapped = true;
          }
        } catch {}
      }

      installXrmHooks();
      installHtmlHooks();
    }
  });

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      const store = window.__whyRecorderStore ||= {
        logs: [],
        recording: false
      };

      function safe(v) {
        try {
          if (v === undefined) return "undefined";
          if (v === null) return "null";
          if (typeof v === "string") return v.slice(0, 1500);
          if (typeof v === "number" || typeof v === "boolean") return String(v);
          if (typeof v === "object") return JSON.stringify(v).slice(0, 2000);
          return String(v).slice(0, 1500);
        } catch {
          return "[unserializable]";
        }
      }

      function extractCaller(stackText) {
        if (!stackText) return "";

        const skipFiles = [
          "content.powerapps.com",
          "ey_Elad.Global.Loader.js",
          "Main_system_library.js",
          "<anonymous>"
        ];

        const lines = stackText.split("\n");
        const crmLines = lines.filter((l) => l.includes("/webresources/"));

        const preferredLine =
          crmLines.find((l) => !skipFiles.some((skip) => l.includes(skip))) ||
          crmLines[0];

        if (!preferredLine) return "";

        const fnMatch = preferredLine.match(/at\s+(.+?)\s+\(/);
        const fileMatch = preferredLine.match(/\/webresources\/([^:]+):(\d+):(\d+)/);

        const fn = fnMatch?.[1]?.trim() || "anonymous";
        const file = fileMatch?.[1] || "";
        const line = fileMatch?.[2] || "";

        return `${fn} (${file}:${line})`;
      }

      function buildRawText() {
        return store.logs
          .map((x, i) => {
            const caller = extractCaller(x.stack);

            return [
              `#${i + 1}`,
              `[${x.time}] ${x.type}`,
              x.attribute ? `attribute: ${x.attribute}` : "",
              x.control ? `control: ${x.control}` : "",
              x.entity ? `entity: ${x.entity}` : "",
              x.id ? `id: ${x.id}` : "",
              x.visible !== undefined ? `visible: ${x.visible}` : "",
              x.disabled !== undefined ? `disabled: ${x.disabled}` : "",
              x.newValue !== undefined ? `newValue: ${x.newValue}` : "",
              x.message ? `message: ${x.message}` : "",
              x.options ? `options: ${x.options}` : "",
              x.text ? `text: ${x.text}` : "",
              x.onclick ? `onclick: ${x.onclick}` : "",
              x.inlineCode ? `inlineCode: ${x.inlineCode}` : "",
              x.eventType ? `eventType: ${x.eventType}` : "",
              x.handlerName ? `handlerName: ${x.handlerName}` : "",
              x.url ? `url: ${x.url}` : "",
              x.method ? `method: ${x.method}` : "",
              x.dataId ? `dataId: ${x.dataId}` : "",
              x.aria ? `aria: ${x.aria}` : "",
              caller ? `caller: ${caller}` : "",
              x.frameUrl ? `frame: ${x.frameUrl}` : ""
            ]
              .filter(Boolean)
              .join("\n");
          })
          .join("\n\n");
      }

      function buildWhy(field) {
        const lower = field.toLowerCase();

        const relevant = store.logs
          .map((x, index) => ({ ...x, index }))
          .filter((x) => {
            const attr = String(x.attribute || "").toLowerCase();
            const control = String(x.control || "").toLowerCase();
            const text = String(x.text || "").toLowerCase();
            const dataId = String(x.dataId || "").toLowerCase();
            const id = String(x.id || "").toLowerCase();
            const inlineCode = String(x.inlineCode || "").toLowerCase();

            return (
              attr.includes(lower) ||
              control.includes(lower) ||
              text.includes(lower) ||
              dataId.includes(lower) ||
              id.includes(lower) ||
              inlineCode.includes(lower)
            );
          });

        if (!relevant.length) {
          return `No events found for: ${field}\n\nTip: click Start Recording, do the action, click Stop Recording, then Analyze Field.`;
        }

        let text = `WHY DID THIS HAPPEN?\n`;
        text += `====================\n\n`;
        text += `Target: ${field}\n`;
        text += `Events found: ${relevant.length}\n\n`;

        relevant.forEach((x, i) => {
          const caller = extractCaller(x.stack);

          text += `#${i + 1} ${x.type}\n`;

          if (x.attribute) text += `Attribute: ${x.attribute}\n`;
          if (x.control) text += `Control: ${x.control}\n`;
          if (x.visible !== undefined) text += `Visible: ${x.visible}\n`;
          if (x.disabled !== undefined) text += `Disabled: ${x.disabled}\n`;
          if (x.newValue !== undefined) text += `New Value: ${x.newValue}\n`;
          if (x.text) text += `Text: ${x.text}\n`;
          if (x.onclick) text += `OnClick: ${x.onclick}\n`;
          if (x.inlineCode) text += `Inline Code: ${x.inlineCode}\n`;
          if (x.eventType) text += `Event: ${x.eventType}\n`;
          if (x.handlerName) text += `Handler: ${x.handlerName}\n`;
          if (x.entity) text += `Entity: ${x.entity}\n`;
          if (x.options) text += `Options: ${x.options}\n`;
          if (caller) text += `Changed/Called By: ${caller}\n`;

          text += `\n`;
        });

        return text;
      }

      function render() {
        const ta = document.getElementById("__whyRecorderText");
        const status = document.getElementById("__whyRecorderStatus");
        if (!ta || !status) return;

        status.textContent = store.recording
          ? `🔴 Recording (${store.logs.length})`
          : `⚪ Idle (${store.logs.length})`;

        ta.value = buildRawText();
        ta.scrollTop = ta.scrollHeight;
      }

      function showModal() {
        document.getElementById("__whyRecorderModal")?.remove();

        const overlay = document.createElement("div");
        overlay.id = "__whyRecorderModal";
        overlay.style.cssText = `
          position: fixed;
          inset: 0;
          background: rgba(0,0,0,.35);
          z-index: 2147483647;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 16px;
          direction: ltr;
        `;

        const box = document.createElement("div");
        box.style.cssText = `
          width: min(1150px, 96vw);
          height: min(760px, 90vh);
          background: white;
          border-radius: 14px;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        `;

        const header = document.createElement("div");
        header.style.cssText = `
          padding: 12px 14px;
          border-bottom: 1px solid #e5e7eb;
          font-weight: 800;
          display: flex;
          justify-content: space-between;
          align-items: center;
        `;
        header.innerHTML = `
          <span>❓ Why Did This Happen</span>
          <span id="__whyRecorderStatus" style="font-size:13px;"></span>
        `;

        const toolbar = document.createElement("div");
        toolbar.style.cssText = `
          display:flex;
          gap:8px;
          padding:10px 14px;
          border-bottom:1px solid #e5e7eb;
          flex-wrap:wrap;
        `;

        function btn(text, color) {
          const b = document.createElement("button");
          b.textContent = text;
          b.style.cssText = `
            border:none;
            background:${color};
            color:white;
            padding:10px 14px;
            border-radius:10px;
            cursor:pointer;
            font-weight:800;
          `;
          return b;
        }

        const startBtn = btn("🎥 Start Recording", "#16a34a");
        const stopBtn = btn("⏹ Stop Recording", "#dc2626");
        const whyBtn = btn("❓ Analyze Field / Button", "#2563eb");
        const rawBtn = btn("Raw Logs", "#334155");
        const clearBtn = btn("🧹 Clear", "#f97316");
        const closeBtn = btn("Close", "#64748b");

        const ta = document.createElement("textarea");
        ta.id = "__whyRecorderText";
        ta.readOnly = true;
        ta.style.cssText = `
          flex:1;
          width:100%;
          border:none;
          resize:none;
          outline:none;
          padding:12px 14px;
          box-sizing:border-box;
          font-family:Consolas, Monaco, 'Courier New', monospace;
          font-size:12px;
          line-height:1.45;
          white-space:pre;
          direction:ltr;
          text-align:left;
        `;

        startBtn.onclick = () => {
          store.recording = true;
          render();
        };

        stopBtn.onclick = () => {
          store.recording = false;
          render();
        };

        clearBtn.onclick = () => {
          store.logs.length = 0;
          render();
        };

        rawBtn.onclick = () => render();

        whyBtn.onclick = () => {
          const field = prompt("Enter field/control/button/id/text\nExample: ey_contactid / showAll / הצג הכל");
          if (!field) return;

          ta.value = buildWhy(field.trim());
          ta.scrollTop = 0;
        };

        closeBtn.onclick = () => overlay.remove();

        toolbar.appendChild(startBtn);
        toolbar.appendChild(stopBtn);
        toolbar.appendChild(whyBtn);
        toolbar.appendChild(rawBtn);
        toolbar.appendChild(clearBtn);
        toolbar.appendChild(closeBtn);

        box.appendChild(header);
        box.appendChild(toolbar);
        box.appendChild(ta);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        render();
      }

      window.__whyRecorderRender = render;
      showModal();
    }
  });
});



document.getElementById("pluginPipelineVisualizer")?.addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });

  if (!tab?.id) {
    alert("No active tab found.");
    return;
  }

  const frameResults = await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      allFrames: true
    },
    world: "MAIN",
    func: () => ({
      hasXrm: !!window.Xrm,
      href: location.href
    })
  });

  const targetFrame = frameResults.find(r => r.result?.hasXrm);

  if (!targetFrame) {
    alert("Could not find Dynamics frame.");
    return;
  }

  await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      frameIds: [targetFrame.frameId]
    },
    world: "MAIN",
    func: () => {
      document.getElementById("rv-plugin-pipeline-popup")?.remove();

      showPluginPipelinePopup();

      function getClientUrl() {
        return Xrm.Utility.getGlobalContext().getClientUrl();
      }

      function escapeHtml(value) {
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
      }

      function removePopup() {
        document.getElementById("rv-plugin-pipeline-popup")?.remove();
      }

      function injectStyle() {
        if (document.getElementById("rv-plugin-pipeline-style")) return;

        const style = document.createElement("style");
        style.id = "rv-plugin-pipeline-style";

        style.textContent = `
          #rv-plugin-pipeline-popup {
            position: fixed;
            top: 18px;
            left: 18px;
            width: 900px;
            max-width: calc(100vw - 36px);
            max-height: calc(100vh - 36px);
            background: #1f1f1f;
            color: white;
            border: 1px solid #444;
            border-radius: 14px;
            z-index: 999999999;
            font-family: Arial, sans-serif;
            box-shadow: 0 12px 40px rgba(0,0,0,.55);
            direction: rtl;
            overflow: hidden;
          }

          .rv-pipe-head {
            padding: 14px 16px;
            border-bottom: 1px solid #333;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #1f1f1f;
          }

          .rv-pipe-head h2 {
            margin: 0;
            font-size: 18px;
          }

          .rv-pipe-close {
            background: transparent;
            border: 0;
            color: white;
            font-size: 26px;
            cursor: pointer;
          }

          .rv-pipe-body {
            padding: 16px;
            overflow-y: auto;
            max-height: calc(100vh - 120px);
          }

          .rv-pipe-row {
            display: grid;
            grid-template-columns: 1fr 1fr auto;
            gap: 10px;
            margin-bottom: 14px;
          }

          .rv-pipe-input,
          .rv-pipe-select {
            background: #2a2a2a;
            color: white;
            border: 1px solid #444;
            border-radius: 8px;
            padding: 10px;
            font-size: 13px;
          }

          .rv-pipe-btn {
            background: #6aa9ff;
            color: black;
            border: 0;
            border-radius: 8px;
            padding: 10px 16px;
            font-weight: bold;
            cursor: pointer;
          }

          .rv-pipe-stage {
            margin-top: 16px;
            border: 1px solid #333;
            border-radius: 12px;
            overflow: hidden;
          }

          .rv-pipe-stage-title {
            background: #2b2b2b;
            color: #6aa9ff;
            padding: 10px 12px;
            font-weight: bold;
          }

          .rv-pipe-step {
            padding: 12px;
            border-top: 1px solid #333;
            direction: ltr;
            text-align: left;
            background: #242424;
          }

          .rv-pipe-step-title {
            font-weight: bold;
            color: #fff;
            margin-bottom: 6px;
          }

          .rv-pipe-meta {
            color: #ccc;
            font-size: 12px;
            white-space: pre-wrap;
            line-height: 1.5;
          }

          .rv-pipe-empty {
            padding: 12px;
            color: #aaa;
            background: #242424;
            direction: ltr;
            text-align: left;
            border-radius: 8px;
          }

          .rv-pipe-badge {
            display: inline-block;
            padding: 3px 7px;
            border-radius: 999px;
            font-size: 11px;
            font-weight: bold;
            margin-right: 6px;
            background: #444;
            color: white;
          }

          .rv-pipe-on {
            background: #236b35;
          }

          .rv-pipe-off {
            background: #7a2a2a;
          }
        `;

        document.head.appendChild(style);
      }

      function showPluginPipelinePopup() {
        injectStyle();
        removePopup();

        const popup = document.createElement("div");
        popup.id = "rv-plugin-pipeline-popup";

        popup.innerHTML = `
          <div class="rv-pipe-head">
            <button class="rv-pipe-close" id="rv-pipe-close">×</button>
            <h2>Plugin Pipeline Visualizer 🧩</h2>
          </div>

          <div class="rv-pipe-body">
            <div class="rv-pipe-row">
              <button id="rv-pipe-run" class="rv-pipe-btn">Load</button>

              <select id="rv-pipe-message" class="rv-pipe-select">
                <option value="Create">Create</option>
                <option value="Update">Update</option>
                <option value="Delete">Delete</option>
                <option value="Retrieve">Retrieve</option>
                <option value="RetrieveMultiple">RetrieveMultiple</option>
                <option value="Assign">Assign</option>
                <option value="SetState">SetState</option>
              </select>

              <input id="rv-pipe-entity" class="rv-pipe-input" placeholder="Entity logical name, example: ey_case" />
            </div>

            <div id="rv-pipe-results"></div>
          </div>
        `;

        document.body.appendChild(popup);

        document.getElementById("rv-pipe-close").onclick = removePopup;
        document.getElementById("rv-pipe-run").onclick = runPipelineSearch;
      }

      async function runPipelineSearch() {
        const entityName = document.getElementById("rv-pipe-entity").value.trim();
        const messageName = document.getElementById("rv-pipe-message").value;
        const results = document.getElementById("rv-pipe-results");

        if (!entityName) {
          alert("Insert entity logical name.");
          return;
        }

        results.innerHTML = `<div class="rv-pipe-empty">Loading...</div>`;

        try {
          const steps = await retrievePluginSteps(entityName, messageName);
          renderPipeline(steps, entityName, messageName);
        } catch (e) {
          results.innerHTML = `<div class="rv-pipe-empty">${escapeHtml(e.message || String(e))}</div>`;
        }
      }

      async function getEntityObjectTypeCode(entityName) {
        const url =
          `${getClientUrl()}/api/data/v9.2/EntityDefinitions(LogicalName='${entityName}')?$select=ObjectTypeCode`;

        const response = await fetch(url, {
          method: "GET",
          headers: {
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
          }
        });

        const text = await response.text();

        if (!response.ok) {
          throw new Error(text);
        }

        const json = JSON.parse(text);

        if (json.ObjectTypeCode === null || json.ObjectTypeCode === undefined) {
          throw new Error("Could not resolve ObjectTypeCode for entity: " + entityName);
        }

        return json.ObjectTypeCode;
      }

      async function retrievePluginSteps(entityName, messageName) {
        const objectTypeCode = await getEntityObjectTypeCode(entityName);

        const fetchXml = `
<fetch>
  <entity name="sdkmessageprocessingstep">
    <attribute name="name" />
    <attribute name="stage" />
    <attribute name="mode" />
    <attribute name="rank" />
    <attribute name="statecode" />
    <attribute name="filteringattributes" />
    <attribute name="configuration" />
    <attribute name="asyncautodelete" />

    <order attribute="stage" />
    <order attribute="rank" />

    <link-entity name="sdkmessagefilter" from="sdkmessagefilterid" to="sdkmessagefilterid" link-type="inner" alias="filter">
      <attribute name="primaryobjecttypecode" />
      <filter>
        <condition attribute="primaryobjecttypecode" operator="eq" value="${objectTypeCode}" />
      </filter>
    </link-entity>

    <link-entity name="sdkmessage" from="sdkmessageid" to="sdkmessageid" link-type="inner" alias="message">
      <attribute name="name" />
      <filter>
        <condition attribute="name" operator="eq" value="${messageName}" />
      </filter>
    </link-entity>

    <link-entity name="plugintype" from="plugintypeid" to="plugintypeid" link-type="outer" alias="ptype">
      <attribute name="typename" />
      <attribute name="friendlyname" />

      <link-entity name="pluginassembly" from="pluginassemblyid" to="pluginassemblyid" link-type="outer" alias="assembly">
        <attribute name="name" />
        <attribute name="version" />
      </link-entity>
    </link-entity>
  </entity>
</fetch>`.trim();

        const url =
          `${getClientUrl()}/api/data/v9.2/sdkmessageprocessingsteps` +
          `?fetchXml=${encodeURIComponent(fetchXml)}`;

        const response = await fetch(url, {
          method: "GET",
          headers: {
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Prefer: 'odata.include-annotations="*"'
          }
        });

        const text = await response.text();

        if (!response.ok) {
          throw new Error(text);
        }

        return JSON.parse(text).value || [];
      }

      function getAliased(row, key) {
        return row[key] ?? row[`${key}@OData.Community.Display.V1.FormattedValue`] ?? "";
      }

      function stageName(stage) {
        const value = Number(stage);

        if (value === 10) return "PreValidation";
        if (value === 20) return "PreOperation";
        if (value === 30) return "MainOperation";
        if (value === 40) return "PostOperation";

        return `Stage ${stage}`;
      }

      function modeName(mode) {
        return Number(mode) === 1 ? "Async" : "Sync";
      }

      function renderPipeline(steps, entityName, messageName) {
        const results = document.getElementById("rv-pipe-results");

        if (!steps.length) {
          results.innerHTML = `
            <div class="rv-pipe-empty">
              No plugin steps found for ${escapeHtml(entityName)} / ${escapeHtml(messageName)}.
            </div>
          `;
          return;
        }

        const stages = [10, 20, 30, 40];

        results.innerHTML = `
          <div class="rv-pipe-empty">
            Found ${steps.length} step(s) for ${escapeHtml(entityName)} / ${escapeHtml(messageName)}
          </div>
          ${stages.map(stage =>
            renderStage(stage, steps.filter(s => Number(s.stage) === stage))
          ).join("")}
        `;
      }

      function renderStage(stage, steps) {
        return `
          <div class="rv-pipe-stage">
            <div class="rv-pipe-stage-title">${stageName(stage)}</div>
            ${
              steps.length
                ? steps.map(renderStep).join("")
                : `<div class="rv-pipe-empty">No steps</div>`
            }
          </div>
        `;
      }

      function renderStep(step) {
        const enabled = Number(step.statecode) === 0;

        const pluginType =
          getAliased(step, "ptype.typename") ||
          getAliased(step, "ptype.friendlyname") ||
          "";

        const assembly =
          getAliased(step, "assembly.name") ||
          "";

        const assemblyVersion =
          getAliased(step, "assembly.version") ||
          "";

        return `
          <div class="rv-pipe-step">
            <div class="rv-pipe-step-title">
              ${escapeHtml(step.name)}
              <span class="rv-pipe-badge ${enabled ? "rv-pipe-on" : "rv-pipe-off"}">
                ${enabled ? "Enabled" : "Disabled"}
              </span>
              <span class="rv-pipe-badge">${escapeHtml(modeName(step.mode))}</span>
              <span class="rv-pipe-badge">Rank ${escapeHtml(step.rank)}</span>
            </div>

            <div class="rv-pipe-meta">
Stage: ${escapeHtml(stageName(step.stage))}
Plugin Type: ${escapeHtml(pluginType)}
Assembly: ${escapeHtml(assembly)} ${escapeHtml(assemblyVersion)}
Filtering Attributes: ${escapeHtml(step.filteringattributes || "-")}
Configuration: ${escapeHtml(step.configuration || "-")}
Async Auto Delete: ${escapeHtml(step.asyncautodelete ?? "-")}
            </div>
          </div>
        `;
      }
    }
  });
});

document.getElementById("pluginTraceExplorer")?.addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });

  if (!tab?.id) {
    alert("No active tab found.");
    return;
  }

  const frameResults = await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      allFrames: true
    },
    world: "MAIN",
    func: () => ({
      hasXrm: !!window.Xrm,
      href: location.href
    })
  });

  const targetFrame = frameResults.find(r => r.result?.hasXrm);

  if (!targetFrame) {
    alert("Could not find Dynamics frame.");
    return;
  }

  await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      frameIds: [targetFrame.frameId]
    },
    world: "MAIN",
    func: () => {
      document.getElementById("rv-plugin-trace-popup")?.remove();
      showPluginTracePopup();

      function getClientUrl() {
        return Xrm.Utility.getGlobalContext().getClientUrl();
      }

      function escapeHtml(value) {
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
      }

      function removePopup() {
        document.getElementById("rv-plugin-trace-popup")?.remove();
      }

      function injectStyle() {
        if (document.getElementById("rv-plugin-trace-style")) return;

        const style = document.createElement("style");
        style.id = "rv-plugin-trace-style";

        style.textContent = `
          #rv-plugin-trace-popup {
            position: fixed;
            top: 18px;
            left: 18px;
            width: 980px;
            max-width: calc(100vw - 36px);
            max-height: calc(100vh - 36px);
            background: #1f1f1f;
            color: white;
            border: 1px solid #444;
            border-radius: 14px;
            z-index: 999999999;
            font-family: Arial, sans-serif;
            box-shadow: 0 12px 40px rgba(0,0,0,.55);
            direction: rtl;
            overflow: hidden;
          }

          .rv-trace-head {
            padding: 14px 16px;
            border-bottom: 1px solid #333;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #1f1f1f;
          }

          .rv-trace-head h2 {
            margin: 0;
            font-size: 18px;
          }

          .rv-trace-close {
            background: transparent;
            border: 0;
            color: white;
            font-size: 26px;
            cursor: pointer;
          }

          .rv-trace-body {
            padding: 16px;
            overflow-y: auto;
            max-height: calc(100vh - 120px);
          }

          .rv-trace-grid {
            display: grid;
            grid-template-columns: 1fr 1fr 1.4fr auto;
            gap: 10px;
            margin-bottom: 10px;
          }

          .rv-trace-grid-2 {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr auto;
            gap: 10px;
            margin-bottom: 14px;
          }

          .rv-trace-input,
          .rv-trace-select {
            background: #2a2a2a;
            color: white;
            border: 1px solid #444;
            border-radius: 8px;
            padding: 10px;
            font-size: 13px;
          }

          .rv-trace-btn {
            background: #6aa9ff;
            color: black;
            border: 0;
            border-radius: 8px;
            padding: 10px 16px;
            font-weight: bold;
            cursor: pointer;
          }

          .rv-trace-btn-secondary {
            background: #444;
            color: white;
          }

          .rv-trace-card {
            background: #242424;
            border: 1px solid #333;
            border-radius: 12px;
            margin-bottom: 14px;
            overflow: hidden;
            direction: ltr;
            text-align: left;
          }

          .rv-trace-card-head {
            background: #2b2b2b;
            padding: 10px 12px;
            display: flex;
            justify-content: space-between;
            gap: 10px;
            align-items: center;
          }

          .rv-trace-title {
            font-weight: bold;
            color: #fff;
            word-break: break-word;
          }

          .rv-trace-meta {
            padding: 10px 12px;
            color: #ccc;
            font-size: 12px;
            line-height: 1.55;
            white-space: pre-wrap;
          }

          .rv-trace-section-title {
            color: #6aa9ff;
            font-weight: bold;
            margin: 10px 12px 6px;
            font-size: 13px;
          }

          .rv-trace-box {
            background: #1b1b1b;
            color: #ddd;
            border: 1px solid #333;
            border-radius: 8px;
            margin: 0 12px 12px;
            padding: 10px;
            font-size: 12px;
            white-space: pre-wrap;
            word-break: break-word;
            max-height: 220px;
            overflow: auto;
          }

          .rv-trace-empty {
            padding: 12px;
            color: #aaa;
            background: #242424;
            direction: ltr;
            text-align: left;
            border-radius: 8px;
          }

          .rv-trace-badge {
            display: inline-block;
            padding: 3px 7px;
            border-radius: 999px;
            font-size: 11px;
            font-weight: bold;
            margin-right: 6px;
            background: #444;
            color: white;
          }

          .rv-trace-error {
            background: #7a2a2a;
          }

          .rv-trace-ok {
            background: #236b35;
          }

          .rv-trace-actions {
            display: flex;
            gap: 8px;
            padding: 0 12px 12px;
          }

          .rv-trace-small-btn {
            background: #444;
            color: white;
            border: 0;
            border-radius: 8px;
            padding: 8px 10px;
            cursor: pointer;
            font-size: 12px;
          }
        `;

        document.head.appendChild(style);
      }

      function showPluginTracePopup() {
        injectStyle();
        removePopup();

        const popup = document.createElement("div");
        popup.id = "rv-plugin-trace-popup";

        popup.innerHTML = `
          <div class="rv-trace-head">
            <button class="rv-trace-close" id="rv-trace-close">×</button>
            <h2>Plugin Trace Explorer 🧾</h2>
          </div>

          <div class="rv-trace-body">
            <div class="rv-trace-grid">
              <input id="rv-trace-entity" class="rv-trace-input" placeholder="Entity, example: ey_case" />
              <select id="rv-trace-message" class="rv-trace-select">
                <option value="">Any Message</option>
                <option value="Create">Create</option>
                <option value="Update">Update</option>
                <option value="Delete">Delete</option>
                <option value="Retrieve">Retrieve</option>
                <option value="RetrieveMultiple">RetrieveMultiple</option>
                <option value="Assign">Assign</option>
                <option value="SetState">SetState</option>
              </select>
              <input id="rv-trace-text" class="rv-trace-input" placeholder="Text contains / exception / correlation id" />
              <button id="rv-trace-run" class="rv-trace-btn">Load</button>
            </div>

            <div class="rv-trace-grid-2">
              <select id="rv-trace-errors" class="rv-trace-select">
                <option value="all">All traces</option>
                <option value="errors">Only errors</option>
              </select>

              <select id="rv-trace-limit" class="rv-trace-select">
                <option value="20">Top 20</option>
                <option value="50">Top 50</option>
                <option value="100">Top 100</option>
              </select>

              <select id="rv-trace-days" class="rv-trace-select">
                <option value="1">Last 1 day</option>
                <option value="3">Last 3 days</option>
                <option value="7">Last 7 days</option>
                <option value="30">Last 30 days</option>
              </select>

              <button id="rv-trace-clear" class="rv-trace-btn rv-trace-btn-secondary">Clear</button>
            </div>

            <div id="rv-trace-results"></div>
          </div>
        `;

        document.body.appendChild(popup);

        document.getElementById("rv-trace-close").onclick = removePopup;
        document.getElementById("rv-trace-run").onclick = runTraceSearch;
        document.getElementById("rv-trace-clear").onclick = () => {
          document.getElementById("rv-trace-results").innerHTML = "";
        };
      }

      async function runTraceSearch() {
        const entity = document.getElementById("rv-trace-entity").value.trim();
        const message = document.getElementById("rv-trace-message").value;
        const text = document.getElementById("rv-trace-text").value.trim();
        const errorsMode = document.getElementById("rv-trace-errors").value;
        const limit = Number(document.getElementById("rv-trace-limit").value || 20);
        const days = Number(document.getElementById("rv-trace-days").value || 1);
        const results = document.getElementById("rv-trace-results");

        results.innerHTML = `<div class="rv-trace-empty">Loading...</div>`;

        try {
          const traces = await retrievePluginTraceLogs({
            entity,
            message,
            text,
            onlyErrors: errorsMode === "errors",
            limit,
            days
          });

          renderTraces(traces);
        } catch (e) {
          results.innerHTML = `<div class="rv-trace-empty">${escapeHtml(e.message || String(e))}</div>`;
        }
      }

      async function retrievePluginTraceLogs({ entity, message, text, onlyErrors, limit, days }) {
        const fromDate = new Date();
        fromDate.setDate(fromDate.getDate() - days);

        const conditions = [];

        conditions.push(
          `<condition attribute="createdon" operator="on-or-after" value="${fromDate.toISOString()}" />`
        );

        if (entity) {
          conditions.push(
            `<condition attribute="primaryentity" operator="eq" value="${escapeXml(entity)}" />`
          );
        }

        if (message) {
          conditions.push(
            `<condition attribute="messagename" operator="eq" value="${escapeXml(message)}" />`
          );
        }

        if (onlyErrors) {
          conditions.push(
            `<condition attribute="exceptiondetails" operator="not-null" />`
          );
        }

        const textFilter = text
          ? `
            <filter type="or">
              <condition attribute="exceptiondetails" operator="like" value="%${escapeXml(text)}%" />
              <condition attribute="messageblock" operator="like" value="%${escapeXml(text)}%" />
              <condition attribute="correlationid" operator="eq" value="${escapeXml(text)}" />
              <condition attribute="typename" operator="like" value="%${escapeXml(text)}%" />
            </filter>
          `
          : "";

        const fetchXml = `
<fetch count="${limit}">
  <entity name="plugintracelog">
    <attribute name="plugintracelogid" />
    <attribute name="typename" />
    <attribute name="messagename" />
    <attribute name="primaryentity" />
    <attribute name="correlationid" />
    <attribute name="depth" />
    <attribute name="mode" />
    <attribute name="operationtype" />
    <attribute name="performanceexecutionstarttime" />
    <attribute name="performanceexecutionduration" />
    <attribute name="exceptiondetails" />
    <attribute name="messageblock" />
    <attribute name="createdon" />

    <order attribute="createdon" descending="true" />

    <filter type="and">
      ${conditions.join("\n")}
      ${textFilter}
    </filter>
  </entity>
</fetch>`.trim();

        const url =
          `${getClientUrl()}/api/data/v9.2/plugintracelogs` +
          `?fetchXml=${encodeURIComponent(fetchXml)}`;

        const response = await fetch(url, {
          method: "GET",
          headers: {
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Prefer: 'odata.include-annotations="*"'
          }
        });

        const responseText = await response.text();

        if (!response.ok) {
          throw new Error(responseText);
        }

        return JSON.parse(responseText).value || [];
      }

      function escapeXml(value) {
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&apos;");
      }

      function modeName(mode) {
        const n = Number(mode);
        if (n === 0) return "Sync";
        if (n === 1) return "Async";
        return mode ?? "-";
      }

      function renderTraces(traces) {
        const results = document.getElementById("rv-trace-results");

        if (!traces.length) {
          results.innerHTML = `<div class="rv-trace-empty">No plugin trace logs found.</div>`;
          return;
        }

        results.innerHTML = traces.map(renderTrace).join("");

        traces.forEach((trace, index) => {
          document.getElementById(`rv-trace-copy-error-${index}`)?.addEventListener("click", async () => {
            await navigator.clipboard.writeText(trace.exceptiondetails || "");
            alert("Copied error");
          });

          document.getElementById(`rv-trace-copy-full-${index}`)?.addEventListener("click", async () => {
            await navigator.clipboard.writeText(JSON.stringify(trace, null, 2));
            alert("Copied full trace JSON");
          });
        });
      }

      function renderTrace(trace, index) {
        const hasError = !!trace.exceptiondetails;

        return `
          <div class="rv-trace-card">
            <div class="rv-trace-card-head">
              <div class="rv-trace-title">
                ${escapeHtml(trace.typename || "-")}
              </div>
              <div>
                <span class="rv-trace-badge ${hasError ? "rv-trace-error" : "rv-trace-ok"}">
                  ${hasError ? "Error" : "OK"}
                </span>
                <span class="rv-trace-badge">${escapeHtml(trace.messagename || "-")}</span>
                <span class="rv-trace-badge">${escapeHtml(trace.primaryentity || "-")}</span>
              </div>
            </div>

            <div class="rv-trace-meta">
Created On: ${escapeHtml(trace["createdon@OData.Community.Display.V1.FormattedValue"] || trace.createdon || "-")}
Correlation Id: ${escapeHtml(trace.correlationid || "-")}
Depth: ${escapeHtml(trace.depth ?? "-")}
Mode: ${escapeHtml(modeName(trace.mode))}
Operation Type: ${escapeHtml(trace.operationtype ?? "-")}
Duration: ${escapeHtml(trace.performanceexecutionduration ?? "-")} ms
Start Time: ${escapeHtml(trace["performanceexecutionstarttime@OData.Community.Display.V1.FormattedValue"] || trace.performanceexecutionstarttime || "-")}
            </div>

            <div class="rv-trace-section-title">Exception Details</div>
            <div class="rv-trace-box">${escapeHtml(trace.exceptiondetails || "-")}</div>

            <div class="rv-trace-section-title">Message Block</div>
            <div class="rv-trace-box">${escapeHtml(trace.messageblock || "-")}</div>

            <div class="rv-trace-actions">
              <button class="rv-trace-small-btn" id="rv-trace-copy-error-${index}">Copy Error</button>
              <button class="rv-trace-small-btn" id="rv-trace-copy-full-${index}">Copy Full JSON</button>
            </div>
          </div>
        `;
      }
    }
  });
});






document.getElementById("fetchBuilderUi").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: async () => {
      document.getElementById("__d365_fetch_builder")?.remove();

      const API_VERSION = "v9.2";
      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const metadataCache = new Map();
      const relationshipCache = new Map();
      const optionCache = new Map();

      const escapeXml = (value) =>
        String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&apos;");

      const escapeHtml = (value) =>
        String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");

      const getLabel = (displayName, fallback) =>
        displayName?.UserLocalizedLabel?.Label ||
        displayName?.LocalizedLabels?.[0]?.Label ||
        fallback ||
        "";

      const createId = (prefix = "id") =>
        `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;

      const normalizeGuid = (value) => String(value || "").replace(/[{}]/g, "").trim();

      const requestJson = async (url, options = {}) => {
        const response = await fetch(url, {
          ...options,
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            ...(options.headers || {})
          }
        });

        const text = await response.text();
        let data = null;

        if (text) {
          try {
            data = JSON.parse(text);
          } catch {
            data = text;
          }
        }

        if (!response.ok) {
          throw new Error(data?.error?.message || data?.Message || text || `HTTP ${response.status}`);
        }

        return data;
      };

      const fetchAll = async (url) => {
        const rows = [];
        let nextUrl = url;

        while (nextUrl) {
          const data = await requestJson(nextUrl);
          rows.push(...(data?.value || []));
          nextUrl = data?.["@odata.nextLink"] || null;
        }

        return rows;
      };

      const defaultFilter = () => ({
        id: createId("filter"),
        type: "and",
        items: []
      });

      const defaultCondition = (field = "") => ({
        id: createId("condition"),
        kind: "condition",
        field,
        operator: "eq",
        value: "",
        values: []
      });

      const state = {
        entities: [],
        entity: null,
        fields: [],
        selectedColumns: [],
        rootFilter: defaultFilter(),
        orders: [],
        links: [],
        options: {
          top: 50,
          distinct: false,
          noLock: false,
          returnTotalRecordCount: false
        },
        activeTab: "builder",
        lastResults: [],
        lastDuration: 0,
        importedXml: ""
      };

      const overlay = document.createElement("div");
      overlay.id = "__d365_fetch_builder";
      overlay.innerHTML = `
        <style>
          #__d365_fetch_builder {
            position: fixed;
            inset: 0;
            z-index: 2147483647;
            background: rgba(15, 23, 42, .48);
            backdrop-filter: blur(3px);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 14px;
            direction: ltr;
            color: #0f172a;
            font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
          }
          #__d365_fetch_builder * { box-sizing: border-box; }
          #__d365_fetch_builder .fb-shell {
            width: min(1500px, 98vw);
            height: min(960px, 96vh);
            background: #f8fafc;
            border: 1px solid rgba(148, 163, 184, .35);
            border-radius: 18px;
            box-shadow: 0 24px 80px rgba(15, 23, 42, .38);
            display: flex;
            flex-direction: column;
            overflow: hidden;
          }
          #__d365_fetch_builder .fb-header {
            height: 66px;
            padding: 0 18px;
            background: #fff;
            border-bottom: 1px solid #e2e8f0;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex: 0 0 auto;
          }
          #__d365_fetch_builder .fb-title { display:flex; align-items:center; gap:11px; }
          #__d365_fetch_builder .fb-logo {
            width: 34px; height: 34px; border-radius: 10px;
            background: linear-gradient(135deg,#f97316,#ea580c);
            color: white; display:grid; place-items:center; font-weight:900;
            box-shadow: 0 8px 18px rgba(234,88,12,.25);
          }
          #__d365_fetch_builder .fb-title-main { font-size:16px; font-weight:900; }
          #__d365_fetch_builder .fb-title-sub { font-size:11px; color:#64748b; margin-top:2px; }
          #__d365_fetch_builder .fb-header-actions { display:flex; gap:8px; align-items:center; }
          #__d365_fetch_builder button,
          #__d365_fetch_builder input,
          #__d365_fetch_builder select,
          #__d365_fetch_builder textarea { font: inherit; }
          #__d365_fetch_builder .fb-btn {
            border: 1px solid #cbd5e1;
            background: #fff;
            color: #0f172a;
            border-radius: 9px;
            padding: 8px 11px;
            font-weight: 750;
            cursor: pointer;
            transition: .15s ease;
            white-space: nowrap;
          }
          #__d365_fetch_builder .fb-btn:hover { background:#f1f5f9; transform:translateY(-1px); }
          #__d365_fetch_builder .fb-btn:disabled { opacity:.45; cursor:not-allowed; transform:none; }
          #__d365_fetch_builder .fb-btn-primary { background:#2563eb; color:#fff; border-color:#2563eb; }
          #__d365_fetch_builder .fb-btn-primary:hover { background:#1d4ed8; }
          #__d365_fetch_builder .fb-btn-success { background:#16a34a; color:#fff; border-color:#16a34a; }
          #__d365_fetch_builder .fb-btn-danger { color:#b91c1c; border-color:#fecaca; background:#fff; }
          #__d365_fetch_builder .fb-btn-ghost { border-color:transparent; background:transparent; }
          #__d365_fetch_builder .fb-btn-small { padding:5px 8px; font-size:12px; border-radius:7px; }
          #__d365_fetch_builder .fb-tabs {
            height: 43px; padding:0 18px; background:#fff; border-bottom:1px solid #e2e8f0;
            display:flex; align-items:flex-end; gap:4px; flex:0 0 auto;
          }
          #__d365_fetch_builder .fb-tab {
            border:none; background:transparent; cursor:pointer; padding:11px 14px 10px;
            font-weight:800; color:#64748b; border-bottom:3px solid transparent;
          }
          #__d365_fetch_builder .fb-tab.active { color:#2563eb; border-bottom-color:#2563eb; }
          #__d365_fetch_builder .fb-main { flex:1; min-height:0; display:flex; }
          #__d365_fetch_builder .fb-page { display:none; width:100%; height:100%; min-height:0; }
          #__d365_fetch_builder .fb-page.active { display:flex; }
          #__d365_fetch_builder .fb-builder-layout { display:grid; grid-template-columns:310px minmax(560px,1fr) 340px; width:100%; min-height:0; }
          #__d365_fetch_builder .fb-panel { min-width:0; min-height:0; background:#fff; border-right:1px solid #e2e8f0; display:flex; flex-direction:column; }
          #__d365_fetch_builder .fb-panel:last-child { border-right:none; }
          #__d365_fetch_builder .fb-panel-head { padding:13px 14px 10px; border-bottom:1px solid #e2e8f0; flex:0 0 auto; }
          #__d365_fetch_builder .fb-panel-title { font-weight:900; font-size:13px; }
          #__d365_fetch_builder .fb-panel-sub { color:#64748b; font-size:11px; margin-top:3px; }
          #__d365_fetch_builder .fb-panel-body { padding:12px; overflow:auto; flex:1; min-height:0; }
          #__d365_fetch_builder .fb-input,
          #__d365_fetch_builder .fb-select,
          #__d365_fetch_builder .fb-textarea {
            width:100%; border:1px solid #cbd5e1; border-radius:8px; padding:8px 10px;
            background:#fff; color:#0f172a; outline:none;
          }
          #__d365_fetch_builder .fb-input:focus,
          #__d365_fetch_builder .fb-select:focus,
          #__d365_fetch_builder .fb-textarea:focus { border-color:#60a5fa; box-shadow:0 0 0 3px rgba(59,130,246,.12); }
          #__d365_fetch_builder .fb-search { margin-bottom:10px; }
          #__d365_fetch_builder .fb-list { border:1px solid #e2e8f0; border-radius:9px; overflow:auto; background:#fff; }
          #__d365_fetch_builder .fb-entity-list { height:240px; }
          #__d365_fetch_builder .fb-field-list { height:310px; }
          #__d365_fetch_builder .fb-list-row {
            display:flex; align-items:flex-start; gap:8px; padding:8px 9px; cursor:pointer;
            border-bottom:1px solid #f1f5f9; font-size:12px;
          }
          #__d365_fetch_builder .fb-list-row:last-child { border-bottom:none; }
          #__d365_fetch_builder .fb-list-row:hover { background:#f8fafc; }
          #__d365_fetch_builder .fb-list-row.active { background:#eff6ff; color:#1d4ed8; }
          #__d365_fetch_builder .fb-row-main { font-weight:750; line-height:1.25; }
          #__d365_fetch_builder .fb-row-sub { color:#64748b; font-size:10px; margin-top:2px; word-break:break-all; }
          #__d365_fetch_builder .fb-section { margin-bottom:16px; }
          #__d365_fetch_builder .fb-section:last-child { margin-bottom:0; }
          #__d365_fetch_builder .fb-section-head { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:7px; }
          #__d365_fetch_builder .fb-section-title { font-size:12px; font-weight:900; }
          #__d365_fetch_builder .fb-muted { color:#64748b; font-size:11px; }
          #__d365_fetch_builder .fb-chip-wrap { display:flex; flex-wrap:wrap; gap:6px; }
          #__d365_fetch_builder .fb-chip {
            display:inline-flex; align-items:center; gap:5px; border:1px solid #bfdbfe; background:#eff6ff;
            color:#1e40af; border-radius:999px; padding:5px 7px; font-size:11px; max-width:100%;
          }
          #__d365_fetch_builder .fb-chip span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:200px; }
          #__d365_fetch_builder .fb-chip button { border:none; background:transparent; cursor:pointer; color:inherit; padding:0 2px; font-weight:900; }
          #__d365_fetch_builder .fb-workspace { background:#f8fafc; overflow:auto; padding:14px; }
          #__d365_fetch_builder .fb-card { background:#fff; border:1px solid #dbe3ee; border-radius:12px; box-shadow:0 2px 8px rgba(15,23,42,.04); margin-bottom:12px; overflow:hidden; }
          #__d365_fetch_builder .fb-card-head { padding:10px 12px; background:#f8fafc; border-bottom:1px solid #e2e8f0; display:flex; justify-content:space-between; align-items:center; gap:8px; }
          #__d365_fetch_builder .fb-card-title { font-weight:900; font-size:12px; }
          #__d365_fetch_builder .fb-card-body { padding:11px 12px; }
          #__d365_fetch_builder .fb-inline { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
          #__d365_fetch_builder .fb-grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
          #__d365_fetch_builder .fb-grid-3 { display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; }
          #__d365_fetch_builder .fb-label { display:block; font-size:10px; color:#64748b; font-weight:800; margin-bottom:4px; }
          #__d365_fetch_builder .fb-filter-group { border:1px solid #cbd5e1; border-radius:10px; background:#fff; margin-top:8px; overflow:hidden; }
          #__d365_fetch_builder .fb-filter-group-head { padding:7px 8px; background:#f8fafc; border-bottom:1px solid #e2e8f0; display:flex; justify-content:space-between; align-items:center; gap:8px; }
          #__d365_fetch_builder .fb-filter-items { padding:8px; display:grid; gap:7px; }
          #__d365_fetch_builder .fb-condition { display:grid; grid-template-columns:minmax(160px,1.35fr) minmax(130px,.85fr) minmax(170px,1fr) auto; gap:7px; align-items:center; }
          #__d365_fetch_builder .fb-link { border-left:4px solid #8b5cf6; }
          #__d365_fetch_builder .fb-link-children { padding-left:18px; }
          #__d365_fetch_builder .fb-empty { border:1px dashed #cbd5e1; border-radius:10px; padding:18px; text-align:center; color:#64748b; font-size:12px; background:#fff; }
          #__d365_fetch_builder .fb-options { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
          #__d365_fetch_builder .fb-check { display:flex; align-items:center; gap:7px; font-size:12px; cursor:pointer; }
          #__d365_fetch_builder .fb-status { padding:8px 10px; border-radius:8px; font-size:11px; background:#f1f5f9; color:#475569; }
          #__d365_fetch_builder .fb-status.ok { background:#ecfdf5; color:#047857; }
          #__d365_fetch_builder .fb-status.error { background:#fef2f2; color:#b91c1c; }
          #__d365_fetch_builder .fb-xml-page.active { display:grid; grid-template-columns:1fr 1fr; gap:12px; padding:14px; overflow:hidden; }
          #__d365_fetch_builder .fb-xml-column { min-width:0; min-height:0; display:flex; flex-direction:column; }
          #__d365_fetch_builder .fb-xml-column .fb-textarea { flex:1; resize:none; font-family:Consolas,Monaco,'Courier New',monospace; font-size:12px; direction:ltr; text-align:left; }
          #__d365_fetch_builder .fb-preview-page { flex-direction:column; padding:14px; gap:10px; overflow:hidden; }
          #__d365_fetch_builder .fb-preview-toolbar { display:flex; justify-content:space-between; align-items:center; gap:10px; flex:0 0 auto; }
          #__d365_fetch_builder .fb-table-wrap { flex:1; min-height:0; overflow:auto; background:#fff; border:1px solid #e2e8f0; border-radius:10px; }
          #__d365_fetch_builder table { border-collapse:collapse; width:max-content; min-width:100%; font-size:11px; }
          #__d365_fetch_builder th { position:sticky; top:0; background:#f1f5f9; z-index:1; text-align:left; font-weight:900; padding:8px; border-bottom:1px solid #cbd5e1; white-space:nowrap; }
          #__d365_fetch_builder td { padding:7px 8px; border-bottom:1px solid #f1f5f9; white-space:nowrap; max-width:360px; overflow:hidden; text-overflow:ellipsis; }
          #__d365_fetch_builder tr:hover td { background:#f8fafc; }
          #__d365_fetch_builder .fb-toast { position:absolute; right:22px; bottom:22px; background:#0f172a; color:#fff; border-radius:9px; padding:10px 13px; font-size:12px; box-shadow:0 12px 30px rgba(15,23,42,.3); opacity:0; transform:translateY(8px); pointer-events:none; transition:.18s ease; }
          #__d365_fetch_builder .fb-toast.show { opacity:1; transform:translateY(0); }
          @media (max-width: 1120px) {
            #__d365_fetch_builder .fb-builder-layout { grid-template-columns:280px 1fr; }
            #__d365_fetch_builder .fb-properties { display:none; }
          }
        </style>
        <div class="fb-shell">
          <div class="fb-header">
            <div class="fb-title">
              <div class="fb-logo">F</div>
              <div>
                <div class="fb-title-main">Advanced FetchXML Builder</div>
                <div class="fb-title-sub">Visual query designer for Dataverse</div>
              </div>
            </div>
            <div class="fb-header-actions">
              <button class="fb-btn" id="fbImportButton">Import XML</button>
              <button class="fb-btn" id="fbCopyButton">Copy XML</button>
              <button class="fb-btn fb-btn-success" id="fbRunButton">Run</button>
              <button class="fb-btn" id="fbCloseButton">Close</button>
            </div>
          </div>
          <div class="fb-tabs">
            <button class="fb-tab active" data-tab="builder">Builder</button>
            <button class="fb-tab" data-tab="preview">Results</button>
            <button class="fb-tab" data-tab="xml">XML / Import</button>
          </div>
          <div class="fb-main">
            <div class="fb-page active" data-page="builder">
              <div class="fb-builder-layout">
                <aside class="fb-panel">
                  <div class="fb-panel-head">
                    <div class="fb-panel-title">Source</div>
                    <div class="fb-panel-sub">Choose a table and columns</div>
                  </div>
                  <div class="fb-panel-body">
                    <div class="fb-section">
                      <div class="fb-section-head"><div class="fb-section-title">Table</div></div>
                      <input class="fb-input fb-search" id="fbEntitySearch" placeholder="Search display or logical name">
                      <div class="fb-list fb-entity-list" id="fbEntityList"></div>
                    </div>
                    <div class="fb-section">
                      <div class="fb-section-head">
                        <div class="fb-section-title">Columns</div>
                        <div class="fb-muted" id="fbColumnCount">0 selected</div>
                      </div>
                      <input class="fb-input fb-search" id="fbFieldSearch" placeholder="Search columns">
                      <div class="fb-list fb-field-list" id="fbFieldList"></div>
                    </div>
                  </div>
                </aside>

                <main class="fb-workspace" id="fbWorkspace"></main>

                <aside class="fb-panel fb-properties">
                  <div class="fb-panel-head">
                    <div class="fb-panel-title">Query settings</div>
                    <div class="fb-panel-sub">Fetch options and output</div>
                  </div>
                  <div class="fb-panel-body">
                    <div class="fb-section">
                      <div class="fb-section-title" style="margin-bottom:8px">Options</div>
                      <div class="fb-options">
                        <div>
                          <label class="fb-label">Top</label>
                          <input class="fb-input" id="fbTop" type="number" min="1" max="5000" value="50">
                        </div>
                        <div>
                          <label class="fb-label">Primary alias</label>
                          <input class="fb-input" id="fbPrimaryAlias" placeholder="optional">
                        </div>
                      </div>
                      <div style="display:grid;gap:7px;margin-top:10px">
                        <label class="fb-check"><input type="checkbox" id="fbDistinct"> Distinct</label>
                        <label class="fb-check"><input type="checkbox" id="fbNoLock"> No lock</label>
                        <label class="fb-check"><input type="checkbox" id="fbCount"> Return total record count</label>
                      </div>
                    </div>
                    <div class="fb-section">
                      <div class="fb-section-head">
                        <div class="fb-section-title">Selected columns</div>
                        <button class="fb-btn fb-btn-small" id="fbClearColumns">Clear</button>
                      </div>
                      <div class="fb-chip-wrap" id="fbSelectedColumnChips"></div>
                    </div>
                    <div class="fb-section">
                      <div class="fb-section-title" style="margin-bottom:7px">Generated XML</div>
                      <textarea class="fb-textarea" id="fbMiniOutput" readonly style="height:250px;font-family:Consolas,monospace;font-size:10px;resize:vertical"></textarea>
                    </div>
                    <div class="fb-status" id="fbStatus">Loading metadata...</div>
                  </div>
                </aside>
              </div>
            </div>

            <div class="fb-page fb-preview-page" data-page="preview">
              <div class="fb-preview-toolbar">
                <div>
                  <div style="font-weight:900">Query results</div>
                  <div class="fb-muted" id="fbResultSummary">Run the query to preview data.</div>
                </div>
                <div class="fb-inline">
                  <button class="fb-btn" id="fbExportCsv">Export CSV</button>
                  <button class="fb-btn fb-btn-success" id="fbRunAgain">Run query</button>
                </div>
              </div>
              <div class="fb-table-wrap" id="fbResultTable"><div class="fb-empty" style="margin:16px">No results yet.</div></div>
            </div>

            <div class="fb-page fb-xml-page" data-page="xml">
              <div class="fb-xml-column">
                <div class="fb-section-head" style="padding-left:0;padding-right:0">
                  <div>
                    <div class="fb-panel-title">Generated FetchXML</div>
                    <div class="fb-panel-sub">Updates automatically</div>
                  </div>
                  <button class="fb-btn fb-btn-small" id="fbCopyXmlTab">Copy</button>
                </div>
                <textarea class="fb-textarea" id="fbOutput" readonly></textarea>
              </div>
              <div class="fb-xml-column">
                <div class="fb-section-head" style="padding-left:0;padding-right:0">
                  <div>
                    <div class="fb-panel-title">Import FetchXML</div>
                    <div class="fb-panel-sub">Paste an existing query and rebuild it visually</div>
                  </div>
                  <button class="fb-btn fb-btn-primary fb-btn-small" id="fbParseImport">Parse and load</button>
                </div>
                <textarea class="fb-textarea" id="fbImportInput" placeholder="Paste FetchXML here..."></textarea>
              </div>
            </div>
          </div>
          <div class="fb-toast" id="fbToast"></div>
        </div>
      `;

      document.body.appendChild(overlay);
      const shell = overlay.querySelector(".fb-shell");
      const $ = (selector) => shell.querySelector(selector);
      const $$ = (selector) => [...shell.querySelectorAll(selector)];

      const entitySearch = $("#fbEntitySearch");
      const entityList = $("#fbEntityList");
      const fieldSearch = $("#fbFieldSearch");
      const fieldList = $("#fbFieldList");
      const workspace = $("#fbWorkspace");
      const output = $("#fbOutput");
      const miniOutput = $("#fbMiniOutput");
      const importInput = $("#fbImportInput");
      const status = $("#fbStatus");
      const resultTable = $("#fbResultTable");
      const resultSummary = $("#fbResultSummary");
      const selectedColumnChips = $("#fbSelectedColumnChips");
      const columnCount = $("#fbColumnCount");

      let toastTimer = null;
      const toast = (message) => {
        const el = $("#fbToast");
        el.textContent = message;
        el.classList.add("show");
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => el.classList.remove("show"), 1800);
      };

      const setStatus = (message, type = "") => {
        status.textContent = message;
        status.className = `fb-status${type ? ` ${type}` : ""}`;
      };

      const setTab = (tab) => {
        state.activeTab = tab;
        $$(".fb-tab").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
        $$(".fb-page").forEach((page) => page.classList.toggle("active", page.dataset.page === tab));
      };

      const getEntity = (logicalName) => state.entities.find((item) => item.logicalName === logicalName);
      const getField = (logicalName, entityLogicalName = state.entity?.logicalName) =>
        metadataCache.get(entityLogicalName)?.fields?.find((field) => field.logicalName === logicalName);

      const getOperatorsByType = (type) => {
        const t = String(type || "").toLowerCase();
        const commonEmpty = [["null", "Is empty"], ["not-null", "Is not empty"]];

        if (t.includes("string") || t.includes("memo")) {
          return [
            ["eq", "Equals"], ["ne", "Not equals"], ["like", "Contains"],
            ["not-like", "Does not contain"], ["begins-with", "Begins with"],
            ["not-begin-with", "Does not begin with"], ["ends-with", "Ends with"],
            ["not-end-with", "Does not end with"], ["in", "In"], ["not-in", "Not in"],
            ...commonEmpty
          ];
        }

        if (t.includes("datetime")) {
          return [
            ["on", "On"], ["on-or-after", "On or after"], ["on-or-before", "On or before"],
            ["today", "Today"], ["yesterday", "Yesterday"], ["tomorrow", "Tomorrow"],
            ["last-seven-days", "Last 7 days"], ["next-seven-days", "Next 7 days"],
            ["last-x-days", "Last X days"], ["next-x-days", "Next X days"],
            ["last-x-months", "Last X months"], ["next-x-months", "Next X months"],
            ["olderthan-x-days", "Older than X days"], ...commonEmpty
          ];
        }

        if (t.includes("lookup") || t.includes("customer") || t.includes("owner")) {
          return [
            ["eq", "Equals GUID"], ["ne", "Not equals GUID"], ["eq-userid", "Current user"],
            ["ne-userid", "Not current user"], ["eq-userteams", "Current user's teams"],
            ["in", "In"], ["not-in", "Not in"], ...commonEmpty
          ];
        }

        if (t.includes("integer") || t.includes("decimal") || t.includes("double") || t.includes("money") || t.includes("bigint")) {
          return [
            ["eq", "Equals"], ["ne", "Not equals"], ["gt", "Greater than"],
            ["ge", "Greater or equal"], ["lt", "Less than"], ["le", "Less or equal"],
            ["between", "Between"], ["not-between", "Not between"], ["in", "In"], ["not-in", "Not in"],
            ...commonEmpty
          ];
        }

        return [["eq", "Equals"], ["ne", "Not equals"], ["in", "In"], ["not-in", "Not in"], ...commonEmpty];
      };

      const operatorNeedsNoValue = (operator) => [
        "null", "not-null", "eq-userid", "ne-userid", "eq-userteams",
        "today", "yesterday", "tomorrow", "last-seven-days", "next-seven-days"
      ].includes(operator);

      const operatorNeedsMultipleValues = (operator) => ["in", "not-in", "between", "not-between"].includes(operator);

      const loadEntities = async () => {
        const url = `${clientUrl}/api/data/${API_VERSION}/EntityDefinitions?$select=LogicalName,SchemaName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsActivity,IsCustomEntity&$filter=IsPrivate eq false`;
        const rows = await fetchAll(url);

        state.entities = rows
          .filter((row) => row.LogicalName && row.EntitySetName)
          .map((row) => ({
            logicalName: row.LogicalName,
            schemaName: row.SchemaName,
            label: getLabel(row.DisplayName, row.LogicalName),
            entitySetName: row.EntitySetName,
            primaryIdAttribute: row.PrimaryIdAttribute,
            primaryNameAttribute: row.PrimaryNameAttribute,
            isActivity: row.IsActivity,
            isCustomEntity: row.IsCustomEntity
          }))
          .sort((a, b) => (a.label || a.logicalName).localeCompare(b.label || b.logicalName));

        renderEntities();
        setStatus(`${state.entities.length} tables loaded`, "ok");
      };

      const loadOptionSet = async (entityLogicalName, field) => {
        const key = `${entityLogicalName}:${field.logicalName}`;
        if (optionCache.has(key)) return optionCache.get(key);

        const type = String(field.type || "").toLowerCase();
        let cast = null;
        if (type.includes("picklist") || type.includes("multiselectpicklist")) cast = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata";
        else if (type.includes("status")) cast = "Microsoft.Dynamics.CRM.StatusAttributeMetadata";
        else if (type.includes("state")) cast = "Microsoft.Dynamics.CRM.StateAttributeMetadata";
        else if (type.includes("boolean")) cast = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata";
        if (!cast) return [];

        try {
          const url = `${clientUrl}/api/data/${API_VERSION}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${field.logicalName}')/${cast}?$expand=OptionSet`;
          const data = await requestJson(url);
          let options = [];

          if (type.includes("boolean")) {
            const trueOption = data?.OptionSet?.TrueOption;
            const falseOption = data?.OptionSet?.FalseOption;
            options = [
              { value: falseOption?.Value ?? 0, label: getLabel(falseOption?.Label, "No") },
              { value: trueOption?.Value ?? 1, label: getLabel(trueOption?.Label, "Yes") }
            ];
          } else {
            options = (data?.OptionSet?.Options || []).map((option) => ({
              value: option.Value,
              label: getLabel(option.Label, String(option.Value))
            }));
          }

          optionCache.set(key, options);
          return options;
        } catch {
          optionCache.set(key, []);
          return [];
        }
      };

      const loadFields = async (entityLogicalName) => {
        if (metadataCache.has(entityLogicalName)) return metadataCache.get(entityLogicalName);

        const url = `${clientUrl}/api/data/${API_VERSION}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,SchemaName,DisplayName,AttributeType,IsValidForRead,IsValidForAdvancedFind,IsPrimaryId,IsPrimaryName`;
        const attrs = await fetchAll(url);
        const fields = attrs
          .filter((attr) => attr.LogicalName && attr.IsValidForRead !== false && !attr.LogicalName.startsWith("_"))
          .map((attr) => ({
            logicalName: attr.LogicalName,
            schemaName: attr.SchemaName,
            label: getLabel(attr.DisplayName, attr.LogicalName),
            type: attr.AttributeType,
            isPrimaryId: attr.IsPrimaryId,
            isPrimaryName: attr.IsPrimaryName,
            isAdvancedFind: attr.IsValidForAdvancedFind,
            options: null
          }))
          .sort((a, b) => (a.label || a.logicalName).localeCompare(b.label || b.logicalName));

        const metadata = { fields };
        metadataCache.set(entityLogicalName, metadata);
        return metadata;
      };

      const loadRelationships = async (entityLogicalName) => {
        if (relationshipCache.has(entityLogicalName)) return relationshipCache.get(entityLogicalName);

        const base = `${clientUrl}/api/data/${API_VERSION}/EntityDefinitions(LogicalName='${entityLogicalName}')`;
        const [manyToOne, oneToMany] = await Promise.all([
          fetchAll(`${base}/ManyToOneRelationships?$select=SchemaName,ReferencedEntity,ReferencingEntity,ReferencingAttribute,ReferencedAttribute`),
          fetchAll(`${base}/OneToManyRelationships?$select=SchemaName,ReferencedEntity,ReferencingEntity,ReferencingAttribute,ReferencedAttribute`)
        ]);

        const map = new Map();
        const add = (relationship, direction) => {
          let relatedEntity;
          let from;
          let to;

          if (direction === "many-to-one") {
            relatedEntity = relationship.ReferencedEntity;
            from = relationship.ReferencedAttribute;
            to = relationship.ReferencingAttribute;
          } else {
            relatedEntity = relationship.ReferencingEntity;
            from = relationship.ReferencingAttribute;
            to = relationship.ReferencedAttribute;
          }

          if (!relatedEntity || !from || !to || relatedEntity === entityLogicalName) return;
          const key = `${relationship.SchemaName}:${direction}`;
          map.set(key, {
            id: key,
            schemaName: relationship.SchemaName,
            direction,
            sourceEntity: entityLogicalName,
            relatedEntity,
            from,
            to
          });
        };

        manyToOne.forEach((item) => add(item, "many-to-one"));
        oneToMany.forEach((item) => add(item, "one-to-many"));

        const relationships = [...map.values()].sort((a, b) => {
          const ea = getEntity(a.relatedEntity);
          const eb = getEntity(b.relatedEntity);
          return (ea?.label || a.relatedEntity).localeCompare(eb?.label || b.relatedEntity);
        });

        relationshipCache.set(entityLogicalName, relationships);
        return relationships;
      };

      const selectEntity = async (logicalName, preserve = false) => {
        const entity = getEntity(logicalName);
        if (!entity) return;

        setStatus(`Loading ${entity.label || entity.logicalName}...`);
        state.entity = entity;
        const metadata = await loadFields(logicalName);
        state.fields = metadata.fields;

        if (!preserve) {
          state.selectedColumns = [];
          state.rootFilter = defaultFilter();
          state.orders = [];
          state.links = [];
        }

        renderAll();
        setStatus(`${state.fields.length} columns loaded`, "ok");
      };

      const renderEntities = () => {
        const q = entitySearch.value.trim().toLowerCase();
        const entities = state.entities.filter((entity) => {
          const text = `${entity.label} ${entity.logicalName} ${entity.schemaName}`.toLowerCase();
          return !q || text.includes(q);
        });

        entityList.innerHTML = entities.slice(0, 500).map((entity) => `
          <div class="fb-list-row ${state.entity?.logicalName === entity.logicalName ? "active" : ""}" data-entity="${escapeHtml(entity.logicalName)}">
            <div>
              <div class="fb-row-main">${escapeHtml(entity.label || entity.logicalName)}</div>
              <div class="fb-row-sub">${escapeHtml(entity.logicalName)}</div>
            </div>
          </div>
        `).join("") || `<div class="fb-empty" style="margin:8px">No tables found.</div>`;

        entityList.querySelectorAll("[data-entity]").forEach((row) => {
          row.onclick = async () => {
            try {
              await selectEntity(row.dataset.entity);
            } catch (error) {
              setStatus(error.message || String(error), "error");
            }
          };
        });
      };

      const renderFields = () => {
        const q = fieldSearch.value.trim().toLowerCase();
        const fields = state.fields.filter((field) => {
          const text = `${field.label} ${field.logicalName} ${field.schemaName} ${field.type}`.toLowerCase();
          return !q || text.includes(q);
        });

        fieldList.innerHTML = state.entity
          ? fields.map((field) => {
              const checked = state.selectedColumns.includes(field.logicalName);
              return `
                <label class="fb-list-row">
                  <input type="checkbox" data-field="${escapeHtml(field.logicalName)}" ${checked ? "checked" : ""}>
                  <div>
                    <div class="fb-row-main">${escapeHtml(field.label || field.logicalName)}</div>
                    <div class="fb-row-sub">${escapeHtml(field.logicalName)} · ${escapeHtml(field.type || "Unknown")}</div>
                  </div>
                </label>
              `;
            }).join("")
          : `<div class="fb-empty" style="margin:8px">Select a table first.</div>`;

        fieldList.querySelectorAll("[data-field]").forEach((checkbox) => {
          checkbox.onchange = () => {
            const name = checkbox.dataset.field;
            if (checkbox.checked) {
              if (!state.selectedColumns.includes(name)) state.selectedColumns.push(name);
            } else {
              state.selectedColumns = state.selectedColumns.filter((item) => item !== name);
            }
            renderSelectedColumns();
            renderWorkspace();
            updateXml();
          };
        });
      };

      const renderSelectedColumns = () => {
        columnCount.textContent = `${state.selectedColumns.length} selected`;
        selectedColumnChips.innerHTML = state.selectedColumns.map((name, index) => {
          const field = getField(name);
          return `
            <div class="fb-chip">
              <button title="Move left" data-move-column="${index}:up">‹</button>
              <span title="${escapeHtml(name)}">${escapeHtml(field?.label || name)}</span>
              <button title="Move right" data-move-column="${index}:down">›</button>
              <button title="Remove" data-remove-column="${escapeHtml(name)}">×</button>
            </div>
          `;
        }).join("") || `<span class="fb-muted">No columns selected.</span>`;

        selectedColumnChips.querySelectorAll("[data-remove-column]").forEach((button) => {
          button.onclick = () => {
            state.selectedColumns = state.selectedColumns.filter((name) => name !== button.dataset.removeColumn);
            renderFields();
            renderSelectedColumns();
            renderWorkspace();
            updateXml();
          };
        });

        selectedColumnChips.querySelectorAll("[data-move-column]").forEach((button) => {
          button.onclick = () => {
            const [indexText, direction] = button.dataset.moveColumn.split(":");
            const index = Number(indexText);
            const target = direction === "up" ? index - 1 : index + 1;
            if (target < 0 || target >= state.selectedColumns.length) return;
            [state.selectedColumns[index], state.selectedColumns[target]] = [state.selectedColumns[target], state.selectedColumns[index]];
            renderSelectedColumns();
            renderWorkspace();
            updateXml();
          };
        });
      };

      const findFilterGroup = (group, id) => {
        if (group.id === id) return group;
        for (const item of group.items) {
          if (item.kind === "group") {
            const found = findFilterGroup(item, id);
            if (found) return found;
          }
        }
        return null;
      };

      const removeFilterItem = (group, itemId) => {
        const index = group.items.findIndex((item) => item.id === itemId);
        if (index >= 0) {
          group.items.splice(index, 1);
          return true;
        }
        return group.items.some((item) => item.kind === "group" && removeFilterItem(item, itemId));
      };

      const renderValueEditor = async (container, condition, entityLogicalName, onChange) => {
        const field = getField(condition.field, entityLogicalName);
        const operator = condition.operator;
        container.innerHTML = "";

        if (operatorNeedsNoValue(operator)) {
          container.innerHTML = `<div class="fb-muted">No value required</div>`;
          return;
        }

        if (operatorNeedsMultipleValues(operator)) {
          const input = document.createElement("input");
          input.className = "fb-input";
          input.placeholder = operator.includes("between") ? "value 1, value 2" : "comma-separated values";
          input.value = condition.values?.length ? condition.values.join(", ") : condition.value || "";
          input.oninput = () => {
            condition.values = input.value.split(",").map((item) => item.trim()).filter(Boolean);
            condition.value = "";
            onChange();
          };
          container.appendChild(input);
          return;
        }

        const type = String(field?.type || "").toLowerCase();
        if (["picklist", "status", "state", "boolean", "multiselectpicklist"].some((token) => type.includes(token))) {
          if (!field.options) field.options = await loadOptionSet(entityLogicalName, field);
          if (field.options?.length) {
            const select = document.createElement("select");
            select.className = "fb-select";
            select.innerHTML = `<option value="">Select value</option>` + field.options.map((option) =>
              `<option value="${escapeHtml(option.value)}" ${String(condition.value) === String(option.value) ? "selected" : ""}>${escapeHtml(option.label)} (${escapeHtml(option.value)})</option>`
            ).join("");
            select.onchange = () => {
              condition.value = select.value;
              onChange();
            };
            container.appendChild(select);
            return;
          }
        }

        const input = document.createElement("input");
        input.className = "fb-input";
        input.value = condition.value ?? "";
        input.placeholder = type.includes("lookup") || type.includes("owner") || type.includes("customer") ? "GUID" : "Value";
        if (type.includes("datetime")) input.type = "date";
        else if (["integer", "decimal", "double", "money", "bigint"].some((token) => type.includes(token))) input.type = "number";
        input.oninput = () => {
          condition.value = input.value;
          onChange();
        };
        container.appendChild(input);
      };

      const renderFilterGroup = (group, entityLogicalName, isRoot = false) => {
        const wrapper = document.createElement("div");
        wrapper.className = "fb-filter-group";
        wrapper.dataset.filterId = group.id;
        wrapper.innerHTML = `
          <div class="fb-filter-group-head">
            <div class="fb-inline">
              <strong style="font-size:11px">Filter group</strong>
              <select class="fb-select" data-filter-type style="width:auto;padding:5px 8px">
                <option value="and" ${group.type === "and" ? "selected" : ""}>AND</option>
                <option value="or" ${group.type === "or" ? "selected" : ""}>OR</option>
              </select>
            </div>
            <div class="fb-inline">
              <button class="fb-btn fb-btn-small" data-add-condition>+ Condition</button>
              <button class="fb-btn fb-btn-small" data-add-group>+ Group</button>
              ${isRoot ? "" : `<button class="fb-btn fb-btn-small fb-btn-danger" data-remove-group>Remove</button>`}
            </div>
          </div>
          <div class="fb-filter-items"></div>
        `;

        const itemsWrap = wrapper.querySelector(".fb-filter-items");
        if (!group.items.length) itemsWrap.innerHTML = `<div class="fb-muted">No conditions.</div>`;

        group.items.forEach((item) => {
          if (item.kind === "group") {
            itemsWrap.appendChild(renderFilterGroup(item, entityLogicalName, false));
            return;
          }

          const conditionRow = document.createElement("div");
          conditionRow.className = "fb-condition";
          const fields = metadataCache.get(entityLogicalName)?.fields || [];
          const selectedField = getField(item.field, entityLogicalName);
          const operators = getOperatorsByType(selectedField?.type);
          if (!operators.some(([value]) => value === item.operator)) item.operator = operators[0]?.[0] || "eq";

          conditionRow.innerHTML = `
            <select class="fb-select" data-condition-field>
              ${fields.map((field) => `<option value="${escapeHtml(field.logicalName)}" ${field.logicalName === item.field ? "selected" : ""}>${escapeHtml(field.label || field.logicalName)} (${escapeHtml(field.logicalName)})</option>`).join("")}
            </select>
            <select class="fb-select" data-condition-operator>
              ${operators.map(([value, label]) => `<option value="${escapeHtml(value)}" ${value === item.operator ? "selected" : ""}>${escapeHtml(label)}</option>`).join("")}
            </select>
            <div data-value-wrap></div>
            <button class="fb-btn fb-btn-small fb-btn-danger" data-remove-condition>×</button>
          `;

          const rerender = () => {
            renderWorkspace();
            updateXml();
          };

          conditionRow.querySelector("[data-condition-field]").onchange = (event) => {
            item.field = event.target.value;
            item.operator = "eq";
            item.value = "";
            item.values = [];
            rerender();
          };

          conditionRow.querySelector("[data-condition-operator]").onchange = (event) => {
            item.operator = event.target.value;
            item.value = "";
            item.values = [];
            rerender();
          };

          conditionRow.querySelector("[data-remove-condition]").onclick = () => {
            removeFilterItem(group, item.id);
            rerender();
          };

          itemsWrap.appendChild(conditionRow);
          renderValueEditor(conditionRow.querySelector("[data-value-wrap]"), item, entityLogicalName, updateXml);
        });

        wrapper.querySelector("[data-filter-type]").onchange = (event) => {
          group.type = event.target.value;
          updateXml();
        };

        wrapper.querySelector("[data-add-condition]").onclick = () => {
          const fields = metadataCache.get(entityLogicalName)?.fields || [];
          group.items.push(defaultCondition(fields[0]?.logicalName || ""));
          renderWorkspace();
          updateXml();
        };

        wrapper.querySelector("[data-add-group]").onclick = () => {
          const child = defaultFilter();
          child.kind = "group";
          group.items.push(child);
          renderWorkspace();
          updateXml();
        };

        wrapper.querySelector("[data-remove-group]")?.addEventListener("click", () => {
          removeFilterItem(state.rootFilter, group.id);
          state.links.forEach((link) => removeFilterItem(link.filter, group.id));
          renderWorkspace();
          updateXml();
        });

        return wrapper;
      };

      const renderOrders = () => {
        const card = document.createElement("div");
        card.className = "fb-card";
        card.innerHTML = `
          <div class="fb-card-head">
            <div class="fb-card-title">Sort order</div>
            <button class="fb-btn fb-btn-small" data-add-order>+ Add sort</button>
          </div>
          <div class="fb-card-body" data-order-list></div>
        `;

        const list = card.querySelector("[data-order-list]");
        if (!state.orders.length) list.innerHTML = `<div class="fb-muted">No sort columns.</div>`;

        state.orders.forEach((order, index) => {
          const row = document.createElement("div");
          row.className = "fb-grid-3";
          row.style.marginBottom = "7px";
          row.innerHTML = `
            <select class="fb-select" data-order-field>
              ${state.fields.map((field) => `<option value="${escapeHtml(field.logicalName)}" ${field.logicalName === order.field ? "selected" : ""}>${escapeHtml(field.label || field.logicalName)}</option>`).join("")}
            </select>
            <select class="fb-select" data-order-direction>
              <option value="false" ${!order.descending ? "selected" : ""}>Ascending</option>
              <option value="true" ${order.descending ? "selected" : ""}>Descending</option>
            </select>
            <div class="fb-inline">
              <button class="fb-btn fb-btn-small" data-order-up>↑</button>
              <button class="fb-btn fb-btn-small" data-order-down>↓</button>
              <button class="fb-btn fb-btn-small fb-btn-danger" data-order-remove>Remove</button>
            </div>
          `;

          row.querySelector("[data-order-field]").onchange = (event) => { order.field = event.target.value; updateXml(); };
          row.querySelector("[data-order-direction]").onchange = (event) => { order.descending = event.target.value === "true"; updateXml(); };
          row.querySelector("[data-order-remove]").onclick = () => { state.orders.splice(index, 1); renderWorkspace(); updateXml(); };
          row.querySelector("[data-order-up]").onclick = () => {
            if (index === 0) return;
            [state.orders[index - 1], state.orders[index]] = [state.orders[index], state.orders[index - 1]];
            renderWorkspace(); updateXml();
          };
          row.querySelector("[data-order-down]").onclick = () => {
            if (index >= state.orders.length - 1) return;
            [state.orders[index + 1], state.orders[index]] = [state.orders[index], state.orders[index + 1]];
            renderWorkspace(); updateXml();
          };
          list.appendChild(row);
        });

        card.querySelector("[data-add-order]").onclick = () => {
          if (!state.fields.length) return;
          state.orders.push({ id: createId("order"), field: state.fields[0].logicalName, descending: false });
          renderWorkspace();
          updateXml();
        };

        return card;
      };

      const getLinksByParent = (parentId = null) => state.links.filter((link) => (link.parentId || null) === parentId);

      const addLink = async (parentEntity, parentId = null) => {
        const relationships = await loadRelationships(parentEntity);
        if (!relationships.length) {
          alert("No supported 1:N or N:1 relationships were found.");
          return;
        }

        const dialog = document.createElement("div");
        dialog.style.cssText = "position:absolute;inset:0;background:rgba(15,23,42,.35);display:grid;place-items:center;z-index:10;padding:20px";
        dialog.innerHTML = `
          <div style="width:min(720px,90vw);max-height:80vh;background:#fff;border-radius:14px;box-shadow:0 20px 60px rgba(15,23,42,.35);display:flex;flex-direction:column;overflow:hidden">
            <div class="fb-card-head"><div class="fb-card-title">Add related table</div><button class="fb-btn fb-btn-small" data-close>Close</button></div>
            <div style="padding:12px"><input class="fb-input" data-search placeholder="Search relationship or table"></div>
            <div data-list style="overflow:auto;padding:0 12px 12px"></div>
          </div>
        `;
        overlay.appendChild(dialog);

        const list = dialog.querySelector("[data-list]");
        const search = dialog.querySelector("[data-search]");
        const render = () => {
          const q = search.value.trim().toLowerCase();
          const filtered = relationships.filter((relationship) => {
            const entity = getEntity(relationship.relatedEntity);
            return `${entity?.label} ${relationship.relatedEntity} ${relationship.schemaName}`.toLowerCase().includes(q);
          });
          list.innerHTML = filtered.map((relationship) => {
            const entity = getEntity(relationship.relatedEntity);
            return `
              <div class="fb-list-row" data-relationship="${escapeHtml(relationship.id)}">
                <div>
                  <div class="fb-row-main">${escapeHtml(entity?.label || relationship.relatedEntity)} (${escapeHtml(relationship.relatedEntity)})</div>
                  <div class="fb-row-sub">${escapeHtml(relationship.schemaName)} · ${escapeHtml(relationship.direction)} · from ${escapeHtml(relationship.from)} to ${escapeHtml(relationship.to)}</div>
                </div>
              </div>
            `;
          }).join("") || `<div class="fb-empty">No relationships found.</div>`;

          list.querySelectorAll("[data-relationship]").forEach((row) => {
            row.onclick = async () => {
              const relationship = relationships.find((item) => item.id === row.dataset.relationship);
              const metadata = await loadFields(relationship.relatedEntity);
              const aliasBase = relationship.relatedEntity.replace(/[^a-z0-9_]/gi, "");
              let alias = aliasBase;
              let suffix = 2;
              while (state.links.some((link) => link.alias === alias)) alias = `${aliasBase}${suffix++}`;

              state.links.push({
                id: createId("link"),
                parentId,
                entity: relationship.relatedEntity,
                relationshipId: relationship.id,
                relationshipSchema: relationship.schemaName,
                from: relationship.from,
                to: relationship.to,
                alias,
                linkType: "inner",
                selectedColumns: [],
                filter: defaultFilter(),
                fields: metadata.fields
              });
              dialog.remove();
              renderWorkspace();
              updateXml();
            };
          });
        };

        search.oninput = render;
        dialog.querySelector("[data-close]").onclick = () => dialog.remove();
        dialog.onclick = (event) => { if (event.target === dialog) dialog.remove(); };
        render();
      };

      const renderLinkCard = (link) => {
        const entity = getEntity(link.entity);
        const card = document.createElement("div");
        card.className = "fb-card fb-link";
        card.innerHTML = `
          <div class="fb-card-head">
            <div>
              <div class="fb-card-title">Related: ${escapeHtml(entity?.label || link.entity)} <span class="fb-muted">(${escapeHtml(link.entity)})</span></div>
              <div class="fb-row-sub">${escapeHtml(link.from)} → ${escapeHtml(link.to)}</div>
            </div>
            <div class="fb-inline">
              <button class="fb-btn fb-btn-small" data-add-child>+ Child link</button>
              <button class="fb-btn fb-btn-small fb-btn-danger" data-remove-link>Remove</button>
            </div>
          </div>
          <div class="fb-card-body">
            <div class="fb-grid-3">
              <div><label class="fb-label">Alias</label><input class="fb-input" data-link-alias value="${escapeHtml(link.alias)}"></div>
              <div><label class="fb-label">Join type</label><select class="fb-select" data-link-type><option value="inner" ${link.linkType === "inner" ? "selected" : ""}>Inner</option><option value="outer" ${link.linkType === "outer" ? "selected" : ""}>Outer</option></select></div>
              <div><label class="fb-label">Relationship</label><input class="fb-input" value="${escapeHtml(link.relationshipSchema || "Custom")}" readonly></div>
            </div>
            <div style="margin-top:10px">
              <label class="fb-label">Columns</label>
              <select class="fb-select" data-link-column>
                <option value="">Add a column...</option>
                ${(link.fields || []).map((field) => `<option value="${escapeHtml(field.logicalName)}">${escapeHtml(field.label || field.logicalName)} (${escapeHtml(field.logicalName)})</option>`).join("")}
              </select>
              <div class="fb-chip-wrap" data-link-chips style="margin-top:7px"></div>
            </div>
            <div style="margin-top:10px" data-link-filter></div>
            <div class="fb-link-children" data-link-children></div>
          </div>
        `;

        const chips = card.querySelector("[data-link-chips]");
        const renderChips = () => {
          chips.innerHTML = link.selectedColumns.map((name) => {
            const field = link.fields.find((item) => item.logicalName === name);
            return `<div class="fb-chip"><span>${escapeHtml(field?.label || name)}</span><button data-remove-link-column="${escapeHtml(name)}">×</button></div>`;
          }).join("") || `<span class="fb-muted">No columns selected.</span>`;
          chips.querySelectorAll("[data-remove-link-column]").forEach((button) => {
            button.onclick = () => {
              link.selectedColumns = link.selectedColumns.filter((name) => name !== button.dataset.removeLinkColumn);
              renderChips();
              updateXml();
            };
          });
        };

        card.querySelector("[data-link-column]").onchange = (event) => {
          const value = event.target.value;
          if (value && !link.selectedColumns.includes(value)) link.selectedColumns.push(value);
          event.target.value = "";
          renderChips();
          updateXml();
        };
        card.querySelector("[data-link-alias]").oninput = (event) => { link.alias = event.target.value.trim(); updateXml(); };
        card.querySelector("[data-link-type]").onchange = (event) => { link.linkType = event.target.value; updateXml(); };
        card.querySelector("[data-remove-link]").onclick = () => {
          const ids = new Set([link.id]);
          let changed = true;
          while (changed) {
            changed = false;
            state.links.forEach((item) => {
              if (item.parentId && ids.has(item.parentId) && !ids.has(item.id)) { ids.add(item.id); changed = true; }
            });
          }
          state.links = state.links.filter((item) => !ids.has(item.id));
          renderWorkspace();
          updateXml();
        };
        card.querySelector("[data-add-child]").onclick = () => addLink(link.entity, link.id);
        card.querySelector("[data-link-filter]").appendChild(renderFilterGroup(link.filter, link.entity, true));
        const childrenWrap = card.querySelector("[data-link-children]");
        getLinksByParent(link.id).forEach((child) => childrenWrap.appendChild(renderLinkCard(child)));
        renderChips();
        return card;
      };

      const renderWorkspace = () => {
        workspace.innerHTML = "";
        if (!state.entity) {
          workspace.innerHTML = `<div class="fb-empty">Choose a table on the left to start building a query.</div>`;
          return;
        }

        const overview = document.createElement("div");
        overview.className = "fb-card";
        overview.innerHTML = `
          <div class="fb-card-head">
            <div>
              <div class="fb-card-title">${escapeHtml(state.entity.label || state.entity.logicalName)}</div>
              <div class="fb-row-sub">${escapeHtml(state.entity.logicalName)} · ${state.selectedColumns.length} columns · ${state.links.length} links</div>
            </div>
            <button class="fb-btn fb-btn-primary fb-btn-small" data-add-root-link>+ Related table</button>
          </div>
          <div class="fb-card-body">
            <div class="fb-muted">Select columns from the left. Add filters, sort rules, and related tables below.</div>
          </div>
        `;
        overview.querySelector("[data-add-root-link]").onclick = () => addLink(state.entity.logicalName, null);
        workspace.appendChild(overview);

        const filterCard = document.createElement("div");
        filterCard.className = "fb-card";
        filterCard.innerHTML = `<div class="fb-card-head"><div class="fb-card-title">Main table filters</div></div><div class="fb-card-body"></div>`;
        filterCard.querySelector(".fb-card-body").appendChild(renderFilterGroup(state.rootFilter, state.entity.logicalName, true));
        workspace.appendChild(filterCard);
        workspace.appendChild(renderOrders());

        const linksCard = document.createElement("div");
        linksCard.className = "fb-card";
        linksCard.innerHTML = `<div class="fb-card-head"><div class="fb-card-title">Related tables</div><button class="fb-btn fb-btn-small" data-add-link>+ Add link</button></div><div class="fb-card-body" data-links></div>`;
        linksCard.querySelector("[data-add-link]").onclick = () => addLink(state.entity.logicalName, null);
        const linksWrap = linksCard.querySelector("[data-links]");
        const rootLinks = getLinksByParent(null);
        if (!rootLinks.length) linksWrap.innerHTML = `<div class="fb-muted">No related tables. Add a link-entity visually.</div>`;
        rootLinks.forEach((link) => linksWrap.appendChild(renderLinkCard(link)));
        workspace.appendChild(linksCard);
      };

      const buildConditionXml = (condition, indent) => {
        if (!condition.field || !condition.operator) return "";
        const attrs = [`attribute="${escapeXml(condition.field)}"`, `operator="${escapeXml(condition.operator)}"`];

        if (operatorNeedsNoValue(condition.operator)) return `${indent}<condition ${attrs.join(" ")} />`;

        if (operatorNeedsMultipleValues(condition.operator)) {
          const values = (condition.values || []).filter((value) => value !== "");
          if (!values.length) return `${indent}<condition ${attrs.join(" ")} />`;
          return `${indent}<condition ${attrs.join(" ")}>
${values.map((value) => `${indent}  <value>${escapeXml(value)}</value>`).join("\n")}
${indent}</condition>`;
        }

        let value = condition.value ?? "";
        if (["like", "not-like"].includes(condition.operator) && !String(value).includes("%")) value = `%${value}%`;
        else if (["begins-with", "not-begin-with"].includes(condition.operator) && !String(value).includes("%")) value = `${value}%`;
        else if (["ends-with", "not-end-with"].includes(condition.operator) && !String(value).includes("%")) value = `%${value}`;
        if (["eq", "ne"].includes(condition.operator)) value = normalizeGuid(value) || value;
        return `${indent}<condition ${attrs.join(" ")} value="${escapeXml(value)}" />`;
      };

      const buildFilterXml = (group, indent) => {
        const parts = group.items.map((item) =>
          item.kind === "group" ? buildFilterXml(item, `${indent}  `) : buildConditionXml(item, `${indent}  `)
        ).filter(Boolean);
        if (!parts.length) return "";
        return `${indent}<filter type="${escapeXml(group.type || "and")}">
${parts.join("\n")}
${indent}</filter>`;
      };

      const buildLinkXml = (link, indent) => {
        const attrs = [
          `name="${escapeXml(link.entity)}"`,
          `from="${escapeXml(link.from)}"`,
          `to="${escapeXml(link.to)}"`,
          `link-type="${escapeXml(link.linkType || "inner")}"`
        ];
        if (link.alias) attrs.push(`alias="${escapeXml(link.alias)}"`);

        const body = [];
        link.selectedColumns.forEach((name) => body.push(`${indent}  <attribute name="${escapeXml(name)}" />`));
        const filterXml = buildFilterXml(link.filter, `${indent}  `);
        if (filterXml) body.push(filterXml);
        getLinksByParent(link.id).forEach((child) => body.push(buildLinkXml(child, `${indent}  `)));

        if (!body.length) return `${indent}<link-entity ${attrs.join(" ")} />`;
        return `${indent}<link-entity ${attrs.join(" ")}>
${body.join("\n")}
${indent}</link-entity>`;
      };

      const buildFetchXml = () => {
        if (!state.entity) return "";
        const fetchAttrs = [];
        const top = Number(state.options.top);
        if (Number.isFinite(top) && top > 0) fetchAttrs.push(`top="${Math.min(top, 5000)}"`);
        if (state.options.distinct) fetchAttrs.push(`distinct="true"`);
        if (state.options.noLock) fetchAttrs.push(`no-lock="true"`);
        if (state.options.returnTotalRecordCount) fetchAttrs.push(`returntotalrecordcount="true"`);

        const entityAttrs = [`name="${escapeXml(state.entity.logicalName)}"`];
        const primaryAlias = $("#fbPrimaryAlias").value.trim();
        if (primaryAlias) entityAttrs.push(`alias="${escapeXml(primaryAlias)}"`);

        const body = [];
        const columns = state.selectedColumns.length
          ? state.selectedColumns
          : [state.entity.primaryNameAttribute || state.entity.primaryIdAttribute].filter(Boolean);
        columns.forEach((name) => body.push(`    <attribute name="${escapeXml(name)}" />`));
        state.orders.forEach((order) => body.push(`    <order attribute="${escapeXml(order.field)}" descending="${order.descending ? "true" : "false"}" />`));
        const rootFilterXml = buildFilterXml(state.rootFilter, "    ");
        if (rootFilterXml) body.push(rootFilterXml);
        getLinksByParent(null).forEach((link) => body.push(buildLinkXml(link, "    ")));

        return `<fetch${fetchAttrs.length ? ` ${fetchAttrs.join(" ")}` : ""}>
  <entity ${entityAttrs.join(" ")}>
${body.join("\n")}
  </entity>
</fetch>`;
      };

      const updateXml = () => {
        const xml = buildFetchXml();
        output.value = xml;
        miniOutput.value = xml;
      };

      const renderAll = () => {
        renderEntities();
        renderFields();
        renderSelectedColumns();
        renderWorkspace();
        updateXml();
      };

      const parseFilterNode = (filterNode) => {
        const group = defaultFilter();
        group.type = filterNode.getAttribute("type") || "and";
        group.items = [];

        [...filterNode.children].forEach((child) => {
          if (child.tagName === "filter") {
            const nested = parseFilterNode(child);
            nested.kind = "group";
            group.items.push(nested);
          } else if (child.tagName === "condition") {
            const condition = defaultCondition(child.getAttribute("attribute") || "");
            condition.operator = child.getAttribute("operator") || "eq";
            condition.value = child.getAttribute("value") || "";
            condition.values = [...child.querySelectorAll(":scope > value")].map((node) => node.textContent || "");
            group.items.push(condition);
          }
        });

        return group;
      };

      const parseLinkNode = async (node, parentId = null) => {
        const entityName = node.getAttribute("name");
        if (!entityName) return;
        const metadata = await loadFields(entityName);
        const link = {
          id: createId("link"),
          parentId,
          entity: entityName,
          relationshipId: "",
          relationshipSchema: "Imported",
          from: node.getAttribute("from") || "",
          to: node.getAttribute("to") || "",
          alias: node.getAttribute("alias") || entityName,
          linkType: node.getAttribute("link-type") || "inner",
          selectedColumns: [...node.querySelectorAll(":scope > attribute")].map((attr) => attr.getAttribute("name")).filter(Boolean),
          filter: node.querySelector(":scope > filter") ? parseFilterNode(node.querySelector(":scope > filter")) : defaultFilter(),
          fields: metadata.fields
        };
        state.links.push(link);

        for (const child of [...node.querySelectorAll(":scope > link-entity")]) {
          await parseLinkNode(child, link.id);
        }
      };

      const importFetchXml = async () => {
        const xml = importInput.value.trim();
        if (!xml) return;
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const parserError = doc.querySelector("parsererror");
        if (parserError) throw new Error("Invalid XML.");

        const fetchNode = doc.querySelector("fetch");
        const entityNode = fetchNode?.querySelector(":scope > entity");
        const entityName = entityNode?.getAttribute("name");
        if (!fetchNode || !entityNode || !entityName) throw new Error("FetchXML must contain fetch > entity.");

        if (!getEntity(entityName)) throw new Error(`Table '${entityName}' was not found in this environment.`);
        await selectEntity(entityName);

        state.options.top = Number(fetchNode.getAttribute("top")) || 50;
        state.options.distinct = fetchNode.getAttribute("distinct") === "true";
        state.options.noLock = fetchNode.getAttribute("no-lock") === "true";
        state.options.returnTotalRecordCount = fetchNode.getAttribute("returntotalrecordcount") === "true";
        $("#fbTop").value = state.options.top;
        $("#fbDistinct").checked = state.options.distinct;
        $("#fbNoLock").checked = state.options.noLock;
        $("#fbCount").checked = state.options.returnTotalRecordCount;
        $("#fbPrimaryAlias").value = entityNode.getAttribute("alias") || "";

        state.selectedColumns = [...entityNode.querySelectorAll(":scope > attribute")].map((attr) => attr.getAttribute("name")).filter(Boolean);
        state.orders = [...entityNode.querySelectorAll(":scope > order")].map((order) => ({
          id: createId("order"),
          field: order.getAttribute("attribute") || "",
          descending: order.getAttribute("descending") === "true"
        }));
        state.rootFilter = entityNode.querySelector(":scope > filter") ? parseFilterNode(entityNode.querySelector(":scope > filter")) : defaultFilter();
        state.links = [];
        for (const linkNode of [...entityNode.querySelectorAll(":scope > link-entity")]) await parseLinkNode(linkNode, null);

        renderAll();
        setTab("builder");
        setStatus("FetchXML imported successfully", "ok");
        toast("FetchXML imported");
      };

      const runQuery = async () => {
        if (!state.entity) {
          alert("Select a table first.");
          return;
        }

        const xml = buildFetchXml();
        const button = $("#fbRunButton");
        button.disabled = true;
        button.textContent = "Running...";
        setStatus("Executing FetchXML...");

        try {
          const started = performance.now();
          const url = `${clientUrl}/api/data/${API_VERSION}/${state.entity.entitySetName}?fetchXml=${encodeURIComponent(xml)}`;
          const data = await requestJson(url, { headers: { Prefer: 'odata.include-annotations="*"' } });
          state.lastDuration = performance.now() - started;
          state.lastResults = data?.value || [];
          renderResults(data);
          setTab("preview");
          setStatus(`Query completed in ${Math.round(state.lastDuration)} ms`, "ok");
        } catch (error) {
          setStatus(error.message || String(error), "error");
          alert(`FetchXML failed:\n${error.message || error}`);
        } finally {
          button.disabled = false;
          button.textContent = "Run";
        }
      };

      const displayValue = (row, key) => {
        const formatted = row[`${key}@OData.Community.Display.V1.FormattedValue`];
        const value = formatted ?? row[key];
        if (value === null || value === undefined) return "";
        if (typeof value === "object") return JSON.stringify(value);
        return String(value);
      };

      const renderResults = (data) => {
        const rows = state.lastResults;
        const total = data?.["@Microsoft.Dynamics.CRM.totalrecordcount"];
        resultSummary.textContent = `${rows.length} rows${typeof total === "number" && total >= 0 ? ` · total ${total}` : ""} · ${Math.round(state.lastDuration)} ms`;

        if (!rows.length) {
          resultTable.innerHTML = `<div class="fb-empty" style="margin:16px">The query returned no rows.</div>`;
          return;
        }

        const keys = [...new Set(rows.flatMap((row) => Object.keys(row).filter((key) => !key.includes("@") && !key.startsWith("_transactioncurrencyid"))))];
        resultTable.innerHTML = `
          <table>
            <thead><tr>${keys.map((key) => `<th>${escapeHtml(key)}</th>`).join("")}</tr></thead>
            <tbody>${rows.map((row) => `<tr>${keys.map((key) => `<td title="${escapeHtml(displayValue(row, key))}">${escapeHtml(displayValue(row, key))}</td>`).join("")}</tr>`).join("")}</tbody>
          </table>
        `;
      };

      const exportCsv = () => {
        if (!state.lastResults.length) return;
        const keys = [...new Set(state.lastResults.flatMap((row) => Object.keys(row).filter((key) => !key.includes("@"))))];
        const csvEscape = (value) => `"${String(value ?? "").replaceAll('"', '""')}"`;
        const csv = [
          keys.map(csvEscape).join(","),
          ...state.lastResults.map((row) => keys.map((key) => csvEscape(displayValue(row, key))).join(","))
        ].join("\r\n");
        const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${state.entity?.logicalName || "fetch-results"}.csv`;
        anchor.click();
        URL.revokeObjectURL(url);
      };

      const copyXml = async () => {
        const xml = buildFetchXml();
        if (!xml) return;
        await navigator.clipboard.writeText(xml);
        toast("FetchXML copied");
      };

      $$(".fb-tab").forEach((button) => button.onclick = () => setTab(button.dataset.tab));
      $("#fbCloseButton").onclick = () => overlay.remove();
      $("#fbCopyButton").onclick = copyXml;
      $("#fbCopyXmlTab").onclick = copyXml;
      $("#fbRunButton").onclick = runQuery;
      $("#fbRunAgain").onclick = runQuery;
      $("#fbExportCsv").onclick = exportCsv;
      $("#fbParseImport").onclick = async () => {
        try { await importFetchXml(); }
        catch (error) { setStatus(error.message || String(error), "error"); alert(error.message || error); }
      };
      $("#fbImportButton").onclick = () => {
        importInput.value = output.value;
        setTab("xml");
        importInput.focus();
        importInput.select();
      };
      $("#fbClearColumns").onclick = () => {
        state.selectedColumns = [];
        renderFields();
        renderSelectedColumns();
        renderWorkspace();
        updateXml();
      };

      entitySearch.oninput = renderEntities;
      fieldSearch.oninput = renderFields;
      $("#fbTop").oninput = (event) => { state.options.top = Number(event.target.value) || 0; updateXml(); };
      $("#fbDistinct").onchange = (event) => { state.options.distinct = event.target.checked; updateXml(); };
      $("#fbNoLock").onchange = (event) => { state.options.noLock = event.target.checked; updateXml(); };
      $("#fbCount").onchange = (event) => { state.options.returnTotalRecordCount = event.target.checked; updateXml(); };
      $("#fbPrimaryAlias").oninput = updateXml;

      overlay.onclick = (event) => {
        if (event.target === overlay) return;
      };
      document.addEventListener("keydown", function closeOnEscape(event) {
        if (event.key === "Escape" && document.getElementById("__d365_fetch_builder")) {
          overlay.remove();
          document.removeEventListener("keydown", closeOnEscape);
        }
      });

      renderAll();
      try {
        await loadEntities();
      } catch (error) {
        setStatus(`Failed loading metadata: ${error.message || error}`, "error");
      }
    }
  });
});








document.getElementById("teamManagementUi")?.addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: {
      tabId: tab.id,
      allFrames: false
    },
    world: "MAIN",
    func: async () => {
      document.getElementById("__d365_team_management")?.remove();

      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

      const state = {
        mode: "manage-team",
        selectedTeam: null,
        selectedUser: null,
        teamSearchResults: [],
        userSearchResults: [],
        teamMembers: [],
        userTeams: []
      };

      const normalizeGuid = (value) =>
        String(value || "")
          .replace(/[{}]/g, "")
          .trim()
          .toLowerCase();

      const escapeHtml = (value) =>
        String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");

      const escapeOData = (value) =>
        String(value ?? "").replaceAll("'", "''");

      const getErrorMessage = (data, fallback) =>
        data?.error?.message ||
        data?.Message ||
        fallback ||
        "Unknown Dataverse error";

      const fetchJson = async (url, options = {}) => {
        const response = await fetch(url, {
          ...options,
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            ...(options.headers || {})
          }
        });

        if (response.status === 204) {
          return null;
        }

        const text = await response.text();
        let data = null;

        if (text) {
          try {
            data = JSON.parse(text);
          } catch {
            data = {
              raw: text
            };
          }
        }

        if (!response.ok) {
          throw new Error(
            getErrorMessage(
              data,
              `Request failed: ${response.status} ${response.statusText}`
            )
          );
        }

        return data;
      };

      const fetchAll = async (url) => {
        const rows = [];

        while (url) {
          const data = await fetchJson(url);

          rows.push(...(data?.value || []));
          url = data?.["@odata.nextLink"] || null;
        }

        return rows;
      };

      const getTeamTypeText = (teamType) => {
        const value = Number(teamType);

        switch (value) {
          case 0:
            return "Owner Team";
          case 1:
            return "Access Team";
          case 2:
            return "Microsoft Entra Security Group";
          case 3:
            return "Microsoft Entra Office Group";
          default:
            return `Team Type ${teamType}`;
        }
      };

      const isManagedExternally = (team) =>
        Number(team?.teamtype) === 2 || Number(team?.teamtype) === 3;

      const overlay = document.createElement("div");
      overlay.id = "__d365_team_management";

      overlay.style.cssText = `
        position:fixed;
        inset:0;
        z-index:2147483647;
        display:flex;
        align-items:center;
        justify-content:center;
        padding:18px;
        background:rgba(15,23,42,.45);
        direction:ltr;
      `;

      const box = document.createElement("div");

      box.style.cssText = `
        width:min(1180px,96vw);
        height:min(850px,94vh);
        display:flex;
        flex-direction:column;
        overflow:hidden;
        border-radius:18px;
        background:#ffffff;
        box-shadow:0 24px 70px rgba(0,0,0,.35);
        font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
        color:#0f172a;
      `;

      box.innerHTML = `
        <div style="
          display:flex;
          align-items:center;
          justify-content:space-between;
          gap:12px;
          padding:16px 18px;
          border-bottom:1px solid #e2e8f0;
          background:#f8fafc;
        ">
          <div>
            <div style="font-size:18px;font-weight:900;">
              👥 Team Management
            </div>

            <div style="margin-top:3px;font-size:12px;color:#64748b;">
              Add users to teams, remove users and inspect team memberships
            </div>
          </div>

          <button
            id="tmClose"
            type="button"
            style="
              padding:9px 14px;
              border:1px solid #cbd5e1;
              border-radius:10px;
              background:#ffffff;
              font-weight:800;
              cursor:pointer;
            "
          >
            Close
          </button>
        </div>

        <div style="
          display:flex;
          gap:8px;
          padding:12px 18px;
          border-bottom:1px solid #e2e8f0;
          background:#ffffff;
        ">
          <button
            id="tmModeManageTeam"
            type="button"
            style="
              padding:10px 14px;
              border:none;
              border-radius:10px;
              background:#2563eb;
              color:#ffffff;
              font-weight:800;
              cursor:pointer;
            "
          >
            Manage Team Members
          </button>

          <button
            id="tmModeUserTeams"
            type="button"
            style="
              padding:10px 14px;
              border:1px solid #cbd5e1;
              border-radius:10px;
              background:#ffffff;
              color:#0f172a;
              font-weight:800;
              cursor:pointer;
            "
          >
            View User Teams
          </button>
        </div>

        <div
          id="tmStatus"
          style="
            display:none;
            margin:12px 18px 0;
            padding:10px 12px;
            border-radius:10px;
            font-size:13px;
            font-weight:700;
          "
        ></div>

        <div
          id="tmManageTeamPage"
          style="
            flex:1;
            min-height:0;
            padding:16px 18px 18px;
            display:grid;
            grid-template-columns:minmax(280px,.85fr) minmax(420px,1.4fr);
            gap:16px;
          "
        >
          <div style="
            min-height:0;
            display:flex;
            flex-direction:column;
            border:1px solid #e2e8f0;
            border-radius:14px;
            overflow:hidden;
          ">
            <div style="padding:14px;border-bottom:1px solid #e2e8f0;">
              <div style="font-weight:900;margin-bottom:8px;">
                1. Search Team
              </div>

              <div style="display:flex;gap:8px;">
                <input
                  id="tmTeamSearch"
                  type="text"
                  placeholder="Team name..."
                  style="
                    flex:1;
                    min-width:0;
                    padding:10px 11px;
                    border:1px solid #cbd5e1;
                    border-radius:10px;
                    outline:none;
                  "
                />

                <button
                  id="tmSearchTeamButton"
                  type="button"
                  style="
                    padding:10px 13px;
                    border:none;
                    border-radius:10px;
                    background:#2563eb;
                    color:#ffffff;
                    font-weight:800;
                    cursor:pointer;
                  "
                >
                  Search
                </button>
              </div>
            </div>

            <div
              id="tmTeamResults"
              style="
                flex:1;
                min-height:0;
                overflow:auto;
                padding:8px;
              "
            >
              <div style="padding:20px;text-align:center;color:#64748b;">
                Search for a team
              </div>
            </div>
          </div>

          <div style="
            min-height:0;
            display:flex;
            flex-direction:column;
            border:1px solid #e2e8f0;
            border-radius:14px;
            overflow:hidden;
          ">
            <div
              id="tmSelectedTeamHeader"
              style="
                padding:14px;
                border-bottom:1px solid #e2e8f0;
                background:#f8fafc;
              "
            >
              <div style="font-weight:900;">
                2. Select a team
              </div>

              <div style="margin-top:4px;font-size:12px;color:#64748b;">
                Team members will appear here
              </div>
            </div>

            <div
              id="tmAddUserArea"
              style="
                display:none;
                padding:12px 14px;
                border-bottom:1px solid #e2e8f0;
              "
            >
              <div style="font-weight:800;margin-bottom:8px;">
                Add user to selected team
              </div>

              <div style="display:flex;gap:8px;">
                <input
                  id="tmUserSearch"
                  type="text"
                  placeholder="User name or email..."
                  style="
                    flex:1;
                    min-width:0;
                    padding:10px 11px;
                    border:1px solid #cbd5e1;
                    border-radius:10px;
                    outline:none;
                  "
                />

                <button
                  id="tmSearchUserButton"
                  type="button"
                  style="
                    padding:10px 13px;
                    border:none;
                    border-radius:10px;
                    background:#0f766e;
                    color:#ffffff;
                    font-weight:800;
                    cursor:pointer;
                  "
                >
                  Search User
                </button>
              </div>

              <div
                id="tmUserSearchResults"
                style="
                  display:none;
                  max-height:210px;
                  overflow:auto;
                  margin-top:8px;
                  border:1px solid #e2e8f0;
                  border-radius:10px;
                  padding:6px;
                  background:#ffffff;
                "
              ></div>
            </div>

            <div
              id="tmMembers"
              style="
                flex:1;
                min-height:0;
                overflow:auto;
                padding:8px;
              "
            >
              <div style="padding:30px;text-align:center;color:#64748b;">
                No team selected
              </div>
            </div>
          </div>
        </div>

        <div
          id="tmUserTeamsPage"
          style="
            flex:1;
            min-height:0;
            display:none;
            padding:16px 18px 18px;
            grid-template-columns:minmax(280px,.85fr) minmax(420px,1.4fr);
            gap:16px;
          "
        >
          <div style="
            min-height:0;
            display:flex;
            flex-direction:column;
            border:1px solid #e2e8f0;
            border-radius:14px;
            overflow:hidden;
          ">
            <div style="padding:14px;border-bottom:1px solid #e2e8f0;">
              <div style="font-weight:900;margin-bottom:8px;">
                1. Search User
              </div>

              <div style="display:flex;gap:8px;">
                <input
                  id="tmInspectUserSearch"
                  type="text"
                  placeholder="User name or email..."
                  style="
                    flex:1;
                    min-width:0;
                    padding:10px 11px;
                    border:1px solid #cbd5e1;
                    border-radius:10px;
                    outline:none;
                  "
                />

                <button
                  id="tmInspectUserSearchButton"
                  type="button"
                  style="
                    padding:10px 13px;
                    border:none;
                    border-radius:10px;
                    background:#2563eb;
                    color:#ffffff;
                    font-weight:800;
                    cursor:pointer;
                  "
                >
                  Search
                </button>
              </div>
            </div>

            <div
              id="tmInspectUserResults"
              style="
                flex:1;
                min-height:0;
                overflow:auto;
                padding:8px;
              "
            >
              <div style="padding:20px;text-align:center;color:#64748b;">
                Search for a user
              </div>
            </div>
          </div>

          <div style="
            min-height:0;
            display:flex;
            flex-direction:column;
            border:1px solid #e2e8f0;
            border-radius:14px;
            overflow:hidden;
          ">
            <div
              id="tmSelectedUserHeader"
              style="
                padding:14px;
                border-bottom:1px solid #e2e8f0;
                background:#f8fafc;
              "
            >
              <div style="font-weight:900;">
                2. Select a user
              </div>

              <div style="margin-top:4px;font-size:12px;color:#64748b;">
                The user's teams will appear here
              </div>
            </div>

            <div
              id="tmUserTeams"
              style="
                flex:1;
                min-height:0;
                overflow:auto;
                padding:8px;
              "
            >
              <div style="padding:30px;text-align:center;color:#64748b;">
                No user selected
              </div>
            </div>
          </div>
        </div>
      `;

      overlay.appendChild(box);
      document.body.appendChild(overlay);

      const $ = (selector) => box.querySelector(selector);

      const elements = {
        close: $("#tmClose"),
        status: $("#tmStatus"),

        modeManageTeam: $("#tmModeManageTeam"),
        modeUserTeams: $("#tmModeUserTeams"),

        manageTeamPage: $("#tmManageTeamPage"),
        userTeamsPage: $("#tmUserTeamsPage"),

        teamSearch: $("#tmTeamSearch"),
        searchTeamButton: $("#tmSearchTeamButton"),
        teamResults: $("#tmTeamResults"),

        selectedTeamHeader: $("#tmSelectedTeamHeader"),
        addUserArea: $("#tmAddUserArea"),
        userSearch: $("#tmUserSearch"),
        searchUserButton: $("#tmSearchUserButton"),
        userSearchResults: $("#tmUserSearchResults"),
        members: $("#tmMembers"),

        inspectUserSearch: $("#tmInspectUserSearch"),
        inspectUserSearchButton: $("#tmInspectUserSearchButton"),
        inspectUserResults: $("#tmInspectUserResults"),
        selectedUserHeader: $("#tmSelectedUserHeader"),
        userTeams: $("#tmUserTeams")
      };

      let statusTimeout = null;

      const showStatus = (message, type = "info") => {
        clearTimeout(statusTimeout);

        const styles = {
          info: {
            background: "#eff6ff",
            color: "#1d4ed8",
            border: "#bfdbfe"
          },
          success: {
            background: "#f0fdf4",
            color: "#15803d",
            border: "#bbf7d0"
          },
          error: {
            background: "#fef2f2",
            color: "#b91c1c",
            border: "#fecaca"
          },
          warning: {
            background: "#fffbeb",
            color: "#a16207",
            border: "#fde68a"
          }
        };

        const selectedStyle = styles[type] || styles.info;

        elements.status.style.display = "block";
        elements.status.style.background = selectedStyle.background;
        elements.status.style.color = selectedStyle.color;
        elements.status.style.border = `1px solid ${selectedStyle.border}`;
        elements.status.textContent = message;

        if (type === "success" || type === "info") {
          statusTimeout = setTimeout(() => {
            elements.status.style.display = "none";
          }, 4000);
        }
      };

      const setButtonLoading = (button, isLoading, loadingText) => {
        if (isLoading) {
          button.dataset.originalText = button.textContent;
          button.textContent = loadingText || "Loading...";
          button.disabled = true;
          button.style.opacity = ".65";
          button.style.cursor = "wait";
        } else {
          button.textContent =
            button.dataset.originalText || button.textContent;
          button.disabled = false;
          button.style.opacity = "1";
          button.style.cursor = "pointer";
        }
      };

      const createEmptyMessage = (text) => `
        <div style="
          padding:30px;
          text-align:center;
          color:#64748b;
        ">
          ${escapeHtml(text)}
        </div>
      `;

      const setMode = (mode) => {
        state.mode = mode;

        const manageSelected = mode === "manage-team";

        elements.manageTeamPage.style.display = manageSelected
          ? "grid"
          : "none";

        elements.userTeamsPage.style.display = manageSelected
          ? "none"
          : "grid";

        elements.modeManageTeam.style.background = manageSelected
          ? "#2563eb"
          : "#ffffff";

        elements.modeManageTeam.style.color = manageSelected
          ? "#ffffff"
          : "#0f172a";

        elements.modeManageTeam.style.border = manageSelected
          ? "none"
          : "1px solid #cbd5e1";

        elements.modeUserTeams.style.background = manageSelected
          ? "#ffffff"
          : "#2563eb";

        elements.modeUserTeams.style.color = manageSelected
          ? "#0f172a"
          : "#ffffff";

        elements.modeUserTeams.style.border = manageSelected
          ? "1px solid #cbd5e1"
          : "none";
      };

      const searchTeams = async () => {
        const searchText = elements.teamSearch.value.trim();

        if (searchText.length < 2) {
          showStatus("Enter at least 2 characters for the team search.", "warning");
          return;
        }

        setButtonLoading(
          elements.searchTeamButton,
          true,
          "Searching..."
        );

        elements.teamResults.innerHTML =
          createEmptyMessage("Searching teams...");

        try {
          const safeText = escapeOData(searchText);

          const url =
            `${clientUrl}/api/data/v9.2/teams` +
            `?$select=teamid,name,teamtype,isdefault,description` +
            `&$filter=contains(name,'${safeText}')` +
            `&$orderby=name asc` +
            `&$top=100`;

          const data = await fetchJson(url);

          state.teamSearchResults = data?.value || [];

          renderTeamResults();
        } catch (error) {
          console.error("Team search failed", error);

          elements.teamResults.innerHTML =
            createEmptyMessage("Failed loading teams");

          showStatus(
            `Team search failed: ${error.message || error}`,
            "error"
          );
        } finally {
          setButtonLoading(elements.searchTeamButton, false);
        }
      };

      const renderTeamResults = () => {
        elements.teamResults.innerHTML = "";

        if (!state.teamSearchResults.length) {
          elements.teamResults.innerHTML =
            createEmptyMessage("No teams found");
          return;
        }

        state.teamSearchResults.forEach((team) => {
          const selected =
            normalizeGuid(state.selectedTeam?.teamid) ===
            normalizeGuid(team.teamid);

          const row = document.createElement("button");
          row.type = "button";

          row.style.cssText = `
            width:100%;
            display:block;
            padding:11px 12px;
            margin-bottom:6px;
            border:1px solid ${selected ? "#93c5fd" : "#e2e8f0"};
            border-radius:10px;
            background:${selected ? "#eff6ff" : "#ffffff"};
            text-align:left;
            cursor:pointer;
          `;

          row.innerHTML = `
            <div style="font-weight:850;color:#0f172a;">
              ${escapeHtml(team.name || "Unnamed Team")}
            </div>

            <div style="
              display:flex;
              gap:6px;
              align-items:center;
              flex-wrap:wrap;
              margin-top:5px;
              font-size:11px;
              color:#64748b;
            ">
              <span>${escapeHtml(getTeamTypeText(team.teamtype))}</span>

              ${
                team.isdefault
                  ? `
                    <span style="
                      padding:2px 6px;
                      border-radius:999px;
                      background:#f1f5f9;
                      color:#475569;
                    ">
                      Default
                    </span>
                  `
                  : ""
              }

              ${
                isManagedExternally(team)
                  ? `
                    <span style="
                      padding:2px 6px;
                      border-radius:999px;
                      background:#fff7ed;
                      color:#c2410c;
                    ">
                      Externally managed
                    </span>
                  `
                  : ""
              }
            </div>
          `;

          row.onclick = () => selectTeam(team);

          elements.teamResults.appendChild(row);
        });
      };

      const selectTeam = async (team) => {
        state.selectedTeam = team;

        renderTeamResults();

        elements.selectedTeamHeader.innerHTML = `
          <div style="font-weight:900;font-size:15px;">
            ${escapeHtml(team.name || "Unnamed Team")}
          </div>

          <div style="margin-top:4px;font-size:12px;color:#64748b;">
            ${escapeHtml(getTeamTypeText(team.teamtype))}
            ·
            ${escapeHtml(team.teamid)}
          </div>

          ${
            isManagedExternally(team)
              ? `
                <div style="
                  margin-top:8px;
                  padding:8px 10px;
                  border:1px solid #fed7aa;
                  border-radius:8px;
                  background:#fff7ed;
                  color:#c2410c;
                  font-size:12px;
                  font-weight:700;
                ">
                  Membership of this Entra group team is managed externally.
                </div>
              `
              : ""
          }
        `;

        elements.addUserArea.style.display = isManagedExternally(team)
          ? "none"
          : "block";

        elements.userSearch.value = "";
        elements.userSearchResults.innerHTML = "";
        elements.userSearchResults.style.display = "none";

        await loadTeamMembers();
      };

      const loadTeamMembers = async () => {
        if (!state.selectedTeam?.teamid) return;

        elements.members.innerHTML =
          createEmptyMessage("Loading team members...");

        try {
          const teamId = normalizeGuid(state.selectedTeam.teamid);

          const url =
            `${clientUrl}/api/data/v9.2/teams(${teamId})` +
            `/teammembership_association` +
            `?$select=systemuserid,fullname,internalemailaddress,domainname,isdisabled` +
            `&$orderby=fullname asc`;

          state.teamMembers = await fetchAll(url);

          renderTeamMembers();
        } catch (error) {
          console.error("Loading team members failed", error);

          elements.members.innerHTML =
            createEmptyMessage("Failed loading team members");

          showStatus(
            `Failed loading team members: ${error.message || error}`,
            "error"
          );
        }
      };

      const renderTeamMembers = () => {
        elements.members.innerHTML = "";

        if (!state.teamMembers.length) {
          elements.members.innerHTML =
            createEmptyMessage("This team has no users");
          return;
        }

        const title = document.createElement("div");
        title.style.cssText = `
          padding:7px 6px 11px;
          font-size:12px;
          color:#64748b;
          font-weight:800;
        `;

        title.textContent = `${state.teamMembers.length} team member(s)`;
        elements.members.appendChild(title);

        state.teamMembers.forEach((user) => {
          const row = document.createElement("div");

          row.style.cssText = `
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            padding:11px 12px;
            margin-bottom:6px;
            border:1px solid #e2e8f0;
            border-radius:10px;
            background:#ffffff;
          `;

          const details = document.createElement("div");
          details.style.cssText = "min-width:0;flex:1;";

          details.innerHTML = `
            <div style="
              font-weight:850;
              overflow:hidden;
              text-overflow:ellipsis;
              white-space:nowrap;
            ">
              ${escapeHtml(user.fullname || user.domainname || "Unnamed User")}

              ${
                user.isdisabled
                  ? `
                    <span style="
                      margin-left:6px;
                      padding:2px 6px;
                      border-radius:999px;
                      background:#fee2e2;
                      color:#b91c1c;
                      font-size:10px;
                    ">
                      Disabled
                    </span>
                  `
                  : ""
              }
            </div>

            <div style="
              margin-top:4px;
              color:#64748b;
              font-size:12px;
              overflow:hidden;
              text-overflow:ellipsis;
              white-space:nowrap;
            ">
              ${escapeHtml(
                user.internalemailaddress ||
                  user.domainname ||
                  user.systemuserid
              )}
            </div>
          `;

          row.appendChild(details);

          if (!isManagedExternally(state.selectedTeam)) {
            const removeButton = document.createElement("button");
            removeButton.type = "button";
            removeButton.textContent = "Remove";

            removeButton.style.cssText = `
              flex:none;
              padding:8px 11px;
              border:1px solid #fecaca;
              border-radius:9px;
              background:#ffffff;
              color:#b91c1c;
              font-weight:800;
              cursor:pointer;
            `;

            removeButton.onclick = () =>
              removeUserFromTeam(user, removeButton);

            row.appendChild(removeButton);
          }

          elements.members.appendChild(row);
        });
      };

      const searchUsersForTeam = async () => {
        if (!state.selectedTeam?.teamid) {
          showStatus("Select a team first.", "warning");
          return;
        }

        if (isManagedExternally(state.selectedTeam)) {
          showStatus(
            "This team is managed by Microsoft Entra ID.",
            "warning"
          );
          return;
        }

        const searchText = elements.userSearch.value.trim();

        if (searchText.length < 2) {
          showStatus("Enter at least 2 characters for user search.", "warning");
          return;
        }

        setButtonLoading(
          elements.searchUserButton,
          true,
          "Searching..."
        );

        elements.userSearchResults.style.display = "block";
        elements.userSearchResults.innerHTML =
          createEmptyMessage("Searching users...");

        try {
          const safeText = escapeOData(searchText);

          const filter =
            `isdisabled eq false and (` +
            `contains(fullname,'${safeText}') or ` +
            `contains(internalemailaddress,'${safeText}') or ` +
            `contains(domainname,'${safeText}')` +
            `)`;

          const url =
            `${clientUrl}/api/data/v9.2/systemusers` +
            `?$select=systemuserid,fullname,internalemailaddress,domainname,isdisabled` +
            `&$filter=${filter}` +
            `&$orderby=fullname asc` +
            `&$top=50`;

          const data = await fetchJson(url);

          const existingMemberIds = new Set(
            state.teamMembers.map((member) =>
              normalizeGuid(member.systemuserid)
            )
          );

          state.userSearchResults = (data?.value || []).map((user) => ({
            ...user,
            isAlreadyMember: existingMemberIds.has(
              normalizeGuid(user.systemuserid)
            )
          }));

          renderUsersForTeam();
        } catch (error) {
          console.error("User search failed", error);

          elements.userSearchResults.innerHTML =
            createEmptyMessage("Failed searching users");

          showStatus(
            `User search failed: ${error.message || error}`,
            "error"
          );
        } finally {
          setButtonLoading(elements.searchUserButton, false);
        }
      };

      const renderUsersForTeam = () => {
        elements.userSearchResults.innerHTML = "";

        if (!state.userSearchResults.length) {
          elements.userSearchResults.innerHTML =
            createEmptyMessage("No users found");
          return;
        }

        state.userSearchResults.forEach((user) => {
          const row = document.createElement("div");

          row.style.cssText = `
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:10px;
            padding:9px 10px;
            margin-bottom:5px;
            border:1px solid #e2e8f0;
            border-radius:9px;
          `;

          const details = document.createElement("div");
          details.style.cssText = "min-width:0;flex:1;";

          details.innerHTML = `
            <div style="
              font-weight:800;
              overflow:hidden;
              text-overflow:ellipsis;
              white-space:nowrap;
            ">
              ${escapeHtml(user.fullname || user.domainname || "Unnamed User")}
            </div>

            <div style="
              margin-top:3px;
              color:#64748b;
              font-size:11px;
              overflow:hidden;
              text-overflow:ellipsis;
              white-space:nowrap;
            ">
              ${escapeHtml(
                user.internalemailaddress ||
                  user.domainname ||
                  user.systemuserid
              )}
            </div>
          `;

          const addButton = document.createElement("button");
          addButton.type = "button";

          if (user.isAlreadyMember) {
            addButton.textContent = "Already Member";
            addButton.disabled = true;

            addButton.style.cssText = `
              flex:none;
              padding:7px 10px;
              border:1px solid #cbd5e1;
              border-radius:8px;
              background:#f1f5f9;
              color:#64748b;
              font-weight:800;
              cursor:not-allowed;
            `;
          } else {
            addButton.textContent = "Add";

            addButton.style.cssText = `
              flex:none;
              padding:7px 12px;
              border:none;
              border-radius:8px;
              background:#16a34a;
              color:#ffffff;
              font-weight:800;
              cursor:pointer;
            `;

            addButton.onclick = () =>
              addUserToTeam(user, addButton);
          }

          row.appendChild(details);
          row.appendChild(addButton);

          elements.userSearchResults.appendChild(row);
        });
      };

      const addUserToTeam = async (user, button) => {
        if (!state.selectedTeam?.teamid || !user?.systemuserid) {
          return;
        }

        const teamName = state.selectedTeam.name || "the selected team";
        const userName =
          user.fullname ||
          user.internalemailaddress ||
          user.domainname ||
          "the selected user";

        const confirmed = window.confirm(
          `Add "${userName}" to team "${teamName}"?`
        );

        if (!confirmed) return;

        setButtonLoading(button, true, "Adding...");

        try {
          const teamId = normalizeGuid(state.selectedTeam.teamid);
          const userId = normalizeGuid(user.systemuserid);

          const url =
            `${clientUrl}/api/data/v9.2/teams(${teamId})` +
            `/teammembership_association/$ref`;

          await fetchJson(url, {
            method: "POST",
            body: JSON.stringify({
              "@odata.id":
                `${clientUrl}/api/data/v9.2/systemusers(${userId})`
            })
          });

          showStatus(
            `${userName} was added to ${teamName}.`,
            "success"
          );

          await loadTeamMembers();
          await searchUsersForTeam();
        } catch (error) {
          console.error("Adding user to team failed", error);

          showStatus(
            `Failed adding user to team: ${error.message || error}`,
            "error"
          );

          setButtonLoading(button, false);
        }
      };

      const removeUserFromTeam = async (user, button) => {
        if (!state.selectedTeam?.teamid || !user?.systemuserid) {
          return;
        }

        const teamName = state.selectedTeam.name || "the selected team";
        const userName =
          user.fullname ||
          user.internalemailaddress ||
          user.domainname ||
          "the selected user";

        const confirmed = window.confirm(
          `Remove "${userName}" from team "${teamName}"?`
        );

        if (!confirmed) return;

        setButtonLoading(button, true, "Removing...");

        try {
          const teamId = normalizeGuid(state.selectedTeam.teamid);
          const userId = normalizeGuid(user.systemuserid);

          const url =
            `${clientUrl}/api/data/v9.2/teams(${teamId})` +
            `/teammembership_association(${userId})/$ref`;

          await fetchJson(url, {
            method: "DELETE"
          });

          showStatus(
            `${userName} was removed from ${teamName}.`,
            "success"
          );

          await loadTeamMembers();
        } catch (error) {
          console.error("Removing user from team failed", error);

          showStatus(
            `Failed removing user from team: ${error.message || error}`,
            "error"
          );

          setButtonLoading(button, false);
        }
      };

      const searchUsersForInspection = async () => {
        const searchText = elements.inspectUserSearch.value.trim();

        if (searchText.length < 2) {
          showStatus("Enter at least 2 characters for user search.", "warning");
          return;
        }

        setButtonLoading(
          elements.inspectUserSearchButton,
          true,
          "Searching..."
        );

        elements.inspectUserResults.innerHTML =
          createEmptyMessage("Searching users...");

        try {
          const safeText = escapeOData(searchText);

          const filter =
            `contains(fullname,'${safeText}') or ` +
            `contains(internalemailaddress,'${safeText}') or ` +
            `contains(domainname,'${safeText}')`;

          const url =
            `${clientUrl}/api/data/v9.2/systemusers` +
            `?$select=systemuserid,fullname,internalemailaddress,domainname,isdisabled` +
            `&$filter=${filter}` +
            `&$orderby=fullname asc` +
            `&$top=100`;

          const data = await fetchJson(url);

          state.userSearchResults = data?.value || [];

          renderInspectionUsers();
        } catch (error) {
          console.error("User inspection search failed", error);

          elements.inspectUserResults.innerHTML =
            createEmptyMessage("Failed searching users");

          showStatus(
            `User search failed: ${error.message || error}`,
            "error"
          );
        } finally {
          setButtonLoading(
            elements.inspectUserSearchButton,
            false
          );
        }
      };

      const renderInspectionUsers = () => {
        elements.inspectUserResults.innerHTML = "";

        if (!state.userSearchResults.length) {
          elements.inspectUserResults.innerHTML =
            createEmptyMessage("No users found");
          return;
        }

        state.userSearchResults.forEach((user) => {
          const selected =
            normalizeGuid(state.selectedUser?.systemuserid) ===
            normalizeGuid(user.systemuserid);

          const row = document.createElement("button");
          row.type = "button";

          row.style.cssText = `
            width:100%;
            display:block;
            padding:11px 12px;
            margin-bottom:6px;
            border:1px solid ${selected ? "#93c5fd" : "#e2e8f0"};
            border-radius:10px;
            background:${selected ? "#eff6ff" : "#ffffff"};
            text-align:left;
            cursor:pointer;
          `;

          row.innerHTML = `
            <div style="font-weight:850;">
              ${escapeHtml(user.fullname || user.domainname || "Unnamed User")}

              ${
                user.isdisabled
                  ? `
                    <span style="
                      margin-left:6px;
                      padding:2px 6px;
                      border-radius:999px;
                      background:#fee2e2;
                      color:#b91c1c;
                      font-size:10px;
                    ">
                      Disabled
                    </span>
                  `
                  : ""
              }
            </div>

            <div style="
              margin-top:4px;
              font-size:12px;
              color:#64748b;
              overflow:hidden;
              text-overflow:ellipsis;
              white-space:nowrap;
            ">
              ${escapeHtml(
                user.internalemailaddress ||
                  user.domainname ||
                  user.systemuserid
              )}
            </div>
          `;

          row.onclick = () => selectInspectionUser(user);

          elements.inspectUserResults.appendChild(row);
        });
      };

      const selectInspectionUser = async (user) => {
        state.selectedUser = user;

        renderInspectionUsers();

        elements.selectedUserHeader.innerHTML = `
          <div style="font-weight:900;font-size:15px;">
            ${escapeHtml(user.fullname || user.domainname || "Unnamed User")}
          </div>

          <div style="margin-top:4px;font-size:12px;color:#64748b;">
            ${escapeHtml(
              user.internalemailaddress ||
                user.domainname ||
                user.systemuserid
            )}
          </div>
        `;

        await loadUserTeams();
      };

      const loadUserTeams = async () => {
        if (!state.selectedUser?.systemuserid) return;

        elements.userTeams.innerHTML =
          createEmptyMessage("Loading user teams...");

        try {
          const userId = normalizeGuid(
            state.selectedUser.systemuserid
          );

          const url =
            `${clientUrl}/api/data/v9.2/systemusers(${userId})` +
            `/teammembership_association` +
            `?$select=teamid,name,teamtype,isdefault,description` +
            `&$orderby=name asc`;

          state.userTeams = await fetchAll(url);

          renderUserTeams();
        } catch (error) {
          console.error("Loading user teams failed", error);

          elements.userTeams.innerHTML =
            createEmptyMessage("Failed loading user teams");

          showStatus(
            `Failed loading user teams: ${error.message || error}`,
            "error"
          );
        }
      };

      const renderUserTeams = () => {
        elements.userTeams.innerHTML = "";

        if (!state.userTeams.length) {
          elements.userTeams.innerHTML =
            createEmptyMessage("This user has no team memberships");
          return;
        }

        const title = document.createElement("div");
        title.style.cssText = `
          padding:7px 6px 11px;
          color:#64748b;
          font-size:12px;
          font-weight:800;
        `;

        title.textContent = `${state.userTeams.length} team(s)`;

        elements.userTeams.appendChild(title);

        state.userTeams.forEach((team) => {
          const row = document.createElement("div");

          row.style.cssText = `
            padding:12px;
            margin-bottom:7px;
            border:1px solid #e2e8f0;
            border-radius:10px;
            background:#ffffff;
          `;

          row.innerHTML = `
            <div style="font-weight:850;">
              ${escapeHtml(team.name || "Unnamed Team")}
            </div>

            <div style="
              display:flex;
              align-items:center;
              gap:6px;
              flex-wrap:wrap;
              margin-top:6px;
              font-size:11px;
              color:#64748b;
            ">
              <span>${escapeHtml(getTeamTypeText(team.teamtype))}</span>

              ${
                team.isdefault
                  ? `
                    <span style="
                      padding:2px 6px;
                      border-radius:999px;
                      background:#f1f5f9;
                      color:#475569;
                    ">
                      Default
                    </span>
                  `
                  : ""
              }

              ${
                isManagedExternally(team)
                  ? `
                    <span style="
                      padding:2px 6px;
                      border-radius:999px;
                      background:#fff7ed;
                      color:#c2410c;
                    ">
                      Externally managed
                    </span>
                  `
                  : ""
              }
            </div>

            <div style="
              margin-top:5px;
              font-family:Consolas,Monaco,monospace;
              font-size:10px;
              color:#94a3b8;
            ">
              ${escapeHtml(team.teamid)}
            </div>
          `;

          elements.userTeams.appendChild(row);
        });
      };

      elements.close.onclick = () => overlay.remove();

  

    
      elements.modeManageTeam.onclick = () =>
        setMode("manage-team");

      elements.modeUserTeams.onclick = () =>
        setMode("user-teams");

      elements.searchTeamButton.onclick = searchTeams;

      elements.teamSearch.onkeydown = (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          searchTeams();
        }
      };

      elements.searchUserButton.onclick = searchUsersForTeam;

      elements.userSearch.onkeydown = (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          searchUsersForTeam();
        }
      };

      elements.inspectUserSearchButton.onclick =
        searchUsersForInspection;

      elements.inspectUserSearch.onkeydown = (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          searchUsersForInspection();
        }
      };

      setMode("manage-team");
    }
  });
});


document
  .getElementById("solutionLayersInspector")
  .addEventListener("click", async () => {
    const tab = await __d365GetActiveTab();
    if (!tab?.id) return;

    await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: false },
      world: "MAIN",
      func: async () => {
        const MODAL_ID = "__d365_solution_layers_only";
        document.getElementById(MODAL_ID)?.remove();

        const clientUrl =
          window.Xrm?.Utility?.getGlobalContext?.()?.getClientUrl?.();

        if (!clientUrl) {
          alert("D365 context not found.");
          return;
        }

        const cleanGuid = value =>
          String(value || "")
            .replace(/[{}]/g, "")
            .trim()
            .toLowerCase();

        const escapeOData = value =>
          String(value || "").replace(/'/g, "''");

        const getLabel = label =>
          label?.UserLocalizedLabel?.Label ||
          label?.LocalizedLabels?.[0]?.Label ||
          "";

        const formatted = (row, field) =>
          row?.[`${field}@OData.Community.Display.V1.FormattedValue`] ??
          row?.[field] ??
          "";

        const requestJson = async url => {
          const response = await fetch(url, {
            method: "GET",
            credentials: "same-origin",
            headers: {
              Accept: "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0",
              Prefer:
                'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=5000'
            }
          });

          const text = await response.text();
          let body;

          try {
            body = text ? JSON.parse(text) : null;
          } catch {
            body = text;
          }

          if (!response.ok) {
            throw new Error(
              body?.error?.message ||
              `HTTP ${response.status} ${response.statusText}`
            );
          }

          return body;
        };

        const apiGet = relativeUrl =>
          requestJson(`${clientUrl}/api/data/v9.2/${relativeUrl}`);

        const apiGetAll = async relativeUrl => {
          const rows = [];
          let nextUrl = relativeUrl.startsWith("http")
            ? relativeUrl
            : `${clientUrl}/api/data/v9.2/${relativeUrl}`;

          while (nextUrl) {
            const body = await requestJson(nextUrl);
            rows.push(...(body?.value || []));
            nextUrl = body?.["@odata.nextLink"] || null;
          }

          return rows;
        };

        const COMPONENT_FAMILIES = {
          table: {
            label: "Table & Table Components",
            mode: "table"
          },
          webResource: {
            label: "Web Resource / JavaScript / HTML / CSS",
            mode: "entity",
            entitySet: "webresourceset",
            idField: "webresourceid",
            nameFields: ["name", "displayname"],
            select: [
              "webresourceid",
              "name",
              "displayname",
              "webresourcetype"
            ]
          },
          pluginAssembly: {
            label: "Plugin Assembly",
            mode: "entity",
            entitySet: "pluginassemblies",
            idField: "pluginassemblyid",
            nameFields: ["name"],
            select: [
              "pluginassemblyid",
              "name",
              "version",
              "publickeytoken",
              "isolationmode"
            ]
          },
          pluginType: {
            label: "Plugin Type",
            mode: "entity",
            entitySet: "plugintypes",
            idField: "plugintypeid",
            nameFields: ["typename", "friendlyname", "name"],
            select: [
              "plugintypeid",
              "typename",
              "friendlyname",
              "name",
              "assemblyname"
            ]
          },
          pluginStep: {
            label: "Plugin Step",
            mode: "entity",
            entitySet: "sdkmessageprocessingsteps",
            idField: "sdkmessageprocessingstepid",
            nameFields: ["name"],
            select: [
              "sdkmessageprocessingstepid",
              "name",
              "stage",
              "mode",
              "rank",
              "statecode"
            ]
          },
          workflow: {
            label: "Workflow / Process / Action / Business Rule",
            mode: "entity",
            entitySet: "workflows",
            idField: "workflowid",
            nameFields: ["name", "uniquename"],
            select: [
              "workflowid",
              "name",
              "uniquename",
              "category",
              "primaryentity",
              "statecode"
            ]
          },
          customApi: {
            label: "Custom API",
            mode: "entity",
            entitySet: "customapis",
            idField: "customapiid",
            nameFields: ["name", "uniquename"],
            select: [
              "customapiid",
              "name",
              "uniquename",
              "bindingtype",
              "boundentitylogicalname"
            ]
          },
          modelDrivenApp: {
            label: "Model-driven App",
            mode: "entity",
            entitySet: "appmodules",
            idField: "appmoduleid",
            nameFields: ["name", "uniquename"],
            select: [
              "appmoduleid",
              "name",
              "uniquename",
              "statecode"
            ]
          },
          siteMap: {
            label: "Site Map",
            mode: "entity",
            entitySet: "sitemaps",
            idField: "sitemapid",
            nameFields: ["sitemapname", "sitemapnameunique"],
            select: [
              "sitemapid",
              "sitemapname",
              "sitemapnameunique"
            ]
          },
          globalChoice: {
            label: "Global Choice",
            mode: "globalChoice"
          },
          environmentVariableDefinition: {
            label: "Environment Variable Definition",
            mode: "entity",
            entitySet: "environmentvariabledefinitions",
            idField: "environmentvariabledefinitionid",
            nameFields: ["displayname", "schemaname"],
            select: [
              "environmentvariabledefinitionid",
              "displayname",
              "schemaname",
              "type",
              "statecode"
            ]
          },
          environmentVariableValue: {
            label: "Environment Variable Value",
            mode: "entity",
            entitySet: "environmentvariablevalues",
            idField: "environmentvariablevalueid",
            nameFields: ["value"],
            select: [
              "environmentvariablevalueid",
              "value",
              "_environmentvariabledefinitionid_value",
              "statecode"
            ]
          },
          securityRole: {
            label: "Security Role",
            mode: "entity",
            entitySet: "roles",
            idField: "roleid",
            nameFields: ["name"],
            select: ["roleid", "name", "statecode"]
          },
          emailTemplate: {
            label: "Email Template",
            mode: "entity",
            entitySet: "templates",
            idField: "templateid",
            nameFields: ["title"],
            select: [
              "templateid",
              "title",
              "templatetypecode",
              "ispersonal"
            ]
          }
        };

        const TABLE_COMPONENT_TYPES = {
          table: "Table",
          column: "Columns",
          relationship: "Relationships",
          form: "Forms",
          view: "System Views",
          chart: "System Charts",
          key: "Alternate Keys",
          localChoice: "Local Choices",
          process: "Processes / Business Rules"
        };

        /*
         * msdyn_componentlayers is a virtual table.
         * It returns rows only when BOTH filters are supplied:
         *   msdyn_componentid
         *   msdyn_solutioncomponentname
         *
         * These names must match the Dataverse solution component names exactly.
         */
        const COMPONENT_TYPE_NAME_MAP = {
          1: "Entity",
          2: "Attribute",
          3: "Relationship",
          4: "AttributePicklistValue",
          5: "AttributeLookupValue",
          6: "ViewAttribute",
          7: "LocalizedLabel",
          9: "OptionSet",
          10: "EntityRelationship",
          13: "ManagedProperty",
          14: "EntityKey",
          16: "Privilege",
          20: "Role",
          21: "RolePrivilege",
          22: "DisplayString",
          24: "Form",
          26: "SavedQuery",
          29: "Workflow",
          31: "Report",
          36: "EmailTemplate",
          44: "DuplicateRule",
          47: "AttributeMap",
          48: "RibbonCommand",
          50: "RibbonCustomization",
          55: "RibbonDiff",
          59: "SavedQueryVisualization",
          60: "SystemForm",
          61: "WebResource",
          62: "SiteMap",
          63: "ConnectionRole",
          66: "CustomControl",
          70: "FieldSecurityProfile",
          71: "FieldPermission",
          80: "AppModule",
          90: "PluginType",
          91: "PluginAssembly",
          92: "SDKMessageProcessingStep",
          93: "SDKMessageProcessingStepImage",
          95: "ServiceEndpoint",
          300: "CanvasApp",
          371: "Connector",
          380: "EnvironmentVariableDefinition",
          381: "EnvironmentVariableValue"
        };


        const overlay = document.createElement("div");
        overlay.id = MODAL_ID;
        overlay.style.cssText = `
          position:fixed; inset:0; z-index:2147483647;
          display:flex; align-items:center; justify-content:center;
          padding:16px; background:rgba(15,23,42,.70);
          backdrop-filter:blur(4px); direction:rtl;
        `;

        const dialog = document.createElement("div");
        dialog.style.cssText = `
          width:min(1500px,97vw); height:min(900px,95vh);
          display:flex; flex-direction:column; overflow:hidden;
          background:#f8fafc; border:1px solid #cbd5e1;
          border-radius:18px; box-shadow:0 30px 90px rgba(0,0,0,.42);
          font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
        `;

        const header = document.createElement("div");
        header.style.cssText = `
          display:flex; align-items:center; justify-content:space-between;
          gap:12px; padding:15px 18px; color:#fff;
          background:linear-gradient(135deg,#7c3aed,#2563eb);
        `;
        header.innerHTML = `
          <div>
            <div style="font-size:18px;font-weight:900">Solution Layers</div>
            <div style="margin-top:3px;font-size:12px;opacity:.86">
              Select any Dataverse component and view its layers only.
            </div>
          </div>
        `;

        const closeTop = document.createElement("button");
        closeTop.textContent = "✕";
        closeTop.style.cssText = `
          width:38px; height:38px; border:1px solid rgba(255,255,255,.35);
          border-radius:10px; color:#fff; background:rgba(255,255,255,.13);
          cursor:pointer; font-size:16px; font-weight:900;
        `;
        header.appendChild(closeTop);

        const content = document.createElement("div");
        content.style.cssText = `
          display:grid; grid-template-columns:minmax(380px,450px) minmax(0,1fr);
          gap:14px; flex:1; min-height:0; padding:14px;
        `;

        const sidebar = document.createElement("div");
        sidebar.style.cssText = `
          display:flex; flex-direction:column; gap:12px; min-height:0;
          padding:14px; overflow:auto; background:#fff;
          border:1px solid #dbe3ef; border-radius:14px;
        `;

        const main = document.createElement("div");
        main.style.cssText = `
          display:flex; flex-direction:column; gap:12px;
          min-width:0; min-height:0;
        `;

        const inputStyle = `
          width:100%; box-sizing:border-box; padding:10px 11px;
          color:#0f172a; background:#fff; border:1px solid #cbd5e1;
          border-radius:10px; outline:none; font-size:13px;
        `;

        const makeField = (labelText, element) => {
          const wrap = document.createElement("div");
          wrap.style.cssText = "display:grid;gap:6px;";

          const label = document.createElement("div");
          label.textContent = labelText;
          label.style.cssText =
            "font-size:12px;font-weight:800;color:#334155;";

          wrap.appendChild(label);
          wrap.appendChild(element);
          return wrap;
        };

        const familySelect = document.createElement("select");
        familySelect.style.cssText = inputStyle;

        Object.entries(COMPONENT_FAMILIES).forEach(([key, config]) => {
          const option = document.createElement("option");
          option.value = key;
          option.textContent = config.label;
          familySelect.appendChild(option);
        });

        const entitySearch = document.createElement("input");
        entitySearch.placeholder =
          "Search table by display/logical/schema name...";
        entitySearch.style.cssText = inputStyle;

        const entitySelect = document.createElement("select");
        entitySelect.size = 7;
        entitySelect.style.cssText = `${inputStyle} min-height:165px;`;

        const tableComponentSelect = document.createElement("select");
        tableComponentSelect.style.cssText = inputStyle;

        Object.entries(TABLE_COMPONENT_TYPES).forEach(([key, label]) => {
          const option = document.createElement("option");
          option.value = key;
          option.textContent = label;
          tableComponentSelect.appendChild(option);
        });

        const componentSearch = document.createElement("input");
        componentSearch.placeholder = "Search existing components...";
        componentSearch.style.cssText = inputStyle;

        const componentSelect = document.createElement("select");
        componentSelect.size = 13;
        componentSelect.style.cssText = `
          ${inputStyle}
          min-height:300px; direction:ltr; text-align:left;
          font-family:Consolas,Monaco,"Courier New",monospace;
          line-height:1.55;
        `;

        const status = document.createElement("div");
        status.style.cssText = `
          min-height:22px; padding:9px 10px; color:#475569;
          background:#f8fafc; border:1px solid #e2e8f0;
          border-radius:10px; font-size:12px; line-height:1.45;
        `;

        const entitySearchField = makeField("2. Search Table", entitySearch);
        const entityField = makeField("Available Tables", entitySelect);
        const tableComponentField = makeField(
          "3. Table Component Type",
          tableComponentSelect
        );

        sidebar.appendChild(makeField("1. Component Type", familySelect));
        sidebar.appendChild(entitySearchField);
        sidebar.appendChild(entityField);
        sidebar.appendChild(tableComponentField);
        sidebar.appendChild(
          makeField("Search Component", componentSearch)
        );
        sidebar.appendChild(
          makeField("Available Components", componentSelect)
        );
        sidebar.appendChild(status);

        const summary = document.createElement("div");
        summary.style.cssText = `
          padding:12px 14px; color:#334155; background:#fff;
          border:1px solid #dbe3ef; border-radius:14px;
          font-size:12px; line-height:1.6;
        `;
        summary.textContent = "Select a component.";

        const resultWrap = document.createElement("div");
        resultWrap.style.cssText = `
          flex:1; min-height:0; overflow:auto; background:#fff;
          border:1px solid #dbe3ef; border-radius:14px;
        `;

        const resultTable = document.createElement("table");
        resultTable.style.cssText = `
          width:100%; border-collapse:collapse; font-size:12px;
          direction:ltr; text-align:left;
        `;
        resultWrap.appendChild(resultTable);

        main.appendChild(summary);
        main.appendChild(resultWrap);

        content.appendChild(sidebar);
        content.appendChild(main);

        const footer = document.createElement("div");
        footer.style.cssText = `
          display:flex; justify-content:flex-end; padding:12px 14px;
          background:#fff; border-top:1px solid #dbe3ef;
        `;

        const searchActiveLayersButton =
          document.createElement("button");
        searchActiveLayersButton.textContent =
          "Search Active Layers";
        searchActiveLayersButton.style.cssText = `
          padding:10px 14px;
          border:none;
          border-radius:10px;
          color:#fff;
          background:#ea580c;
          cursor:pointer;
          font-weight:900;
        `;

        const closeButton = document.createElement("button");
        closeButton.textContent = "Close";
        closeButton.style.cssText = `
          padding:10px 14px; border:1px solid #cbd5e1;
          border-radius:10px; background:#fff; cursor:pointer;
          font-weight:800;
        `;

        footer.appendChild(searchActiveLayersButton);
        footer.appendChild(closeButton);

        dialog.appendChild(header);
        dialog.appendChild(content);
        dialog.appendChild(footer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        let allEntities = [];
        let filteredEntities = [];
        let allComponents = [];
        let filteredComponents = [];
        let selectedEntity = null;
        let loadToken = 0;

        const setStatus = (message, type = "normal") => {
          const colors = {
            normal: "#475569",
            success: "#047857",
            warning: "#b45309",
            error: "#dc2626"
          };
          status.textContent = message;
          status.style.color = colors[type] || colors.normal;
        };

        const renderSelect = (select, rows, textFactory, emptyText) => {
          select.innerHTML = "";

          if (!rows.length) {
            const option = document.createElement("option");
            option.value = "";
            option.textContent = emptyText;
            select.appendChild(option);
            return;
          }

          rows.forEach((row, index) => {
            const option = document.createElement("option");
            option.value = String(index);
            option.textContent = textFactory(row);
            select.appendChild(option);
          });
        };

        const renderEntities = () => {
          renderSelect(
            entitySelect,
            filteredEntities,
            entity =>
              `${entity.displayName || "(No display name)"} | ${entity.logicalName}`,
            "No tables found"
          );
        };

        const componentText = component =>
          [component.displayName, component.name, component.extra]
            .filter(Boolean)
            .join(" | ");

        const renderComponents = () => {
          renderSelect(
            componentSelect,
            filteredComponents,
            componentText,
            "No components found"
          );
        };

        const renderLayers = layers => {
          resultTable.innerHTML = "";

          const columns = [
            ["order", "Order"],
            ["solution", "Solution"],
            ["publisher", "Publisher"],
            ["componentType", "Component Type"],
            ["layerName", "Layer Name"],
            ["overwriteTime", "Overwrite Time"],
            ["hasChanges", "Has Changes"]
          ];

          const thead = document.createElement("thead");
          const headerRow = document.createElement("tr");

          columns.forEach(([, label]) => {
            const th = document.createElement("th");
            th.textContent = label;
            th.style.cssText = `
              position:sticky; top:0; z-index:1; padding:10px 9px;
              color:#fff; background:#0f172a;
              border-bottom:1px solid #334155; white-space:nowrap;
            `;
            headerRow.appendChild(th);
          });

          thead.appendChild(headerRow);
          resultTable.appendChild(thead);

          const tbody = document.createElement("tbody");

          if (!layers.length) {
            const tr = document.createElement("tr");
            const td = document.createElement("td");
            td.colSpan = columns.length;
            td.textContent = "No layers found.";
            td.style.cssText =
              "padding:24px;text-align:center;color:#64748b;";
            tr.appendChild(td);
            tbody.appendChild(tr);
          } else {
            layers.forEach((layer, index) => {
              const tr = document.createElement("tr");
              tr.style.background =
                index % 2 === 0 ? "#fff" : "#f8fafc";

              columns.forEach(([key]) => {
                const td = document.createElement("td");
                td.textContent = String(layer[key] ?? "");
                td.title = td.textContent;
                td.style.cssText = `
                  padding:9px; color:#1e293b;
                  border-bottom:1px solid #e2e8f0;
                  white-space:nowrap; max-width:450px;
                  overflow:hidden; text-overflow:ellipsis;
                `;
                tr.appendChild(td);
              });

              tbody.appendChild(tr);
            });
          }

          resultTable.appendChild(tbody);
        };


        const isActiveLayer = layer => {
          const values = [
            layer?.solution,
            layer?.layerName
          ]
            .filter(Boolean)
            .map(value =>
              String(value).trim().toLowerCase()
            );

          return values.some(value =>
            value === "active" ||
            value === "active layer" ||
            value.includes("active")
          );
        };

        const renderActiveLayerResults = rows => {
          resultTable.innerHTML = "";

          const columns = [
            ["componentDisplay", "Component"],
            ["componentName", "Logical / Unique Name"],
            ["componentFamily", "Selected Type"],
            ["order", "Order"],
            ["solution", "Solution"],
            ["publisher", "Publisher"],
            ["componentType", "Layer Component Type"],
            ["layerName", "Layer Name"],
            ["overwriteTime", "Overwrite Time"]
          ];

          const thead = document.createElement("thead");
          const headerRow = document.createElement("tr");

          columns.forEach(([, label]) => {
            const th = document.createElement("th");
            th.textContent = label;
            th.style.cssText = `
              position:sticky;
              top:0;
              z-index:1;
              padding:10px 9px;
              color:#fff;
              background:#9a3412;
              border-bottom:1px solid #7c2d12;
              white-space:nowrap;
            `;
            headerRow.appendChild(th);
          });

          thead.appendChild(headerRow);
          resultTable.appendChild(thead);

          const tbody = document.createElement("tbody");

          if (!rows.length) {
            const tr = document.createElement("tr");
            const td = document.createElement("td");
            td.colSpan = columns.length;
            td.textContent =
              "No active layers found in the currently loaded components.";
            td.style.cssText =
              "padding:24px;text-align:center;color:#64748b;";
            tr.appendChild(td);
            tbody.appendChild(tr);
          } else {
            rows.forEach((row, index) => {
              const tr = document.createElement("tr");
              tr.style.background =
                index % 2 === 0 ? "#fff7ed" : "#ffedd5";

              columns.forEach(([key]) => {
                const td = document.createElement("td");
                td.textContent = String(row[key] ?? "");
                td.title = td.textContent;
                td.style.cssText = `
                  padding:9px;
                  color:#431407;
                  border-bottom:1px solid #fed7aa;
                  white-space:nowrap;
                  max-width:430px;
                  overflow:hidden;
                  text-overflow:ellipsis;
                `;
                tr.appendChild(td);
              });

              tbody.appendChild(tr);
            });
          }

          resultTable.appendChild(tbody);
        };

        const mapWithConcurrency = async (
          items,
          concurrency,
          worker
        ) => {
          const results = new Array(items.length);
          let nextIndex = 0;

          const runner = async () => {
            while (true) {
              const index = nextIndex++;
              if (index >= items.length) return;

              results[index] = await worker(
                items[index],
                index
              );
            }
          };

          await Promise.all(
            Array.from(
              {
                length: Math.min(
                  concurrency,
                  Math.max(items.length, 1)
                )
              },
              runner
            )
          );

          return results;
        };

        const searchActiveLayers = async () => {
          if (!allComponents.length) {
            setStatus(
              "⚠️ No components are loaded. Select a component type first.",
              "warning"
            );
            return;
          }

          const token = ++loadToken;
          const components = [...allComponents];
          const family =
            COMPONENT_FAMILIES[familySelect.value];

          searchActiveLayersButton.disabled = true;
          searchActiveLayersButton.textContent =
            "Searching Active Layers...";

          summary.innerHTML = `
            <b>Active layer scan:</b>
            ${components.length} loaded component(s)
          `;

          renderActiveLayerResults([]);

          let completed = 0;
          const activeRows = [];

          try {
            await mapWithConcurrency(
              components,
              4,
              async component => {
                if (token !== loadToken) return;

                try {
                  const layers =
                    await getAllLayers(component);

                  const activeLayers =
                    layers.filter(isActiveLayer);

                  activeLayers.forEach(layer => {
                    activeRows.push({
                      componentDisplay:
                        component.displayName ||
                        component.name ||
                        component.id,
                      componentName:
                        component.name || "",
                      componentFamily:
                        familySelect.value === "table"
                          ? TABLE_COMPONENT_TYPES[
                              tableComponentSelect.value
                            ]
                          : family.label,
                      ...layer
                    });
                  });
                } catch (error) {
                  console.warn(
                    "Active layer scan failed for component:",
                    component,
                    error
                  );
                } finally {
                  completed++;

                  if (token === loadToken) {
                    setStatus(
                      `⏳ Scanned ${completed}/${components.length}; ` +
                      `found ${activeRows.length} active layer(s)...`
                    );
                  }
                }
              }
            );

            if (token !== loadToken) return;

            activeRows.sort((a, b) => {
              const componentCompare =
                String(a.componentDisplay || "")
                  .localeCompare(
                    String(b.componentDisplay || "")
                  );

              if (componentCompare !== 0) {
                return componentCompare;
              }

              return (
                Number(a.order ?? 0) -
                Number(b.order ?? 0)
              );
            });

            renderActiveLayerResults(activeRows);

            summary.innerHTML = `
              <b>Active layer scan:</b>
              ${components.length} component(s)
              &nbsp;&nbsp;|&nbsp;&nbsp;
              <b>Active layers:</b>
              ${activeRows.length}
            `;

            setStatus(
              activeRows.length
                ? `✅ Found ${activeRows.length} active layer(s).`
                : "✅ Scan completed. No active layers were found.",
              "success"
            );

            console.table(activeRows);
          } finally {
            searchActiveLayersButton.disabled = false;
            searchActiveLayersButton.textContent =
              "Search Active Layers";
          }
        };

        const loadEntities = async () => {
          const response = await apiGet(
            "EntityDefinitions?" +
            "$select=MetadataId,LogicalName,SchemaName,ObjectTypeCode,DisplayName"
          );

          allEntities = (response.value || [])
            .filter(row => row.MetadataId && row.LogicalName)
            .map(row => ({
              id: cleanGuid(row.MetadataId),
              logicalName: row.LogicalName,
              schemaName: row.SchemaName || "",
              objectTypeCode: row.ObjectTypeCode,
              displayName: getLabel(row.DisplayName),
              raw: row
            }))
            .sort((a, b) =>
              (a.displayName || a.logicalName).localeCompare(
                b.displayName || b.logicalName
              )
            );

          filteredEntities = [...allEntities];
          renderEntities();
        };

        const loadTableComponents = async () => {
          if (!selectedEntity) return [];

          const type = tableComponentSelect.value;
          const logicalName = escapeOData(selectedEntity.logicalName);

          if (type === "table") {
            return [{
              id: selectedEntity.id,
              name: selectedEntity.logicalName,
              displayName: selectedEntity.displayName,
              extra: selectedEntity.schemaName,
              raw: selectedEntity.raw
            }];
          }

          if (type === "column" || type === "localChoice") {
            const response = await apiGet(
              `EntityDefinitions(LogicalName='${logicalName}')/Attributes?` +
              "$select=MetadataId,LogicalName,SchemaName,DisplayName,AttributeType"
            );

            let rows = (response.value || [])
              .filter(row => row.MetadataId)
              .map(row => ({
                id: cleanGuid(row.MetadataId),
                name: row.LogicalName || "",
                displayName: getLabel(row.DisplayName),
                extra: row.SchemaName || row.AttributeType || "",
                raw: row
              }));

            if (type === "localChoice") {
              rows = rows.filter(row =>
                ["Picklist", "State", "Status", "MultiSelectPicklist"]
                  .includes(row.raw?.AttributeType)
              );
            }

            return rows;
          }

          if (type === "relationship") {
            const response = await apiGet(
              `EntityDefinitions(LogicalName='${logicalName}')?` +
              "$select=LogicalName" +
              "&$expand=" +
              "OneToManyRelationships($select=MetadataId,SchemaName,ReferencedEntity,ReferencingEntity)," +
              "ManyToOneRelationships($select=MetadataId,SchemaName,ReferencedEntity,ReferencingEntity)," +
              "ManyToManyRelationships($select=MetadataId,SchemaName,Entity1LogicalName,Entity2LogicalName)"
            );

            const rows = [
              ...(response.OneToManyRelationships || []).map(row => ({
                ...row,
                relType: "1:N"
              })),
              ...(response.ManyToOneRelationships || []).map(row => ({
                ...row,
                relType: "N:1"
              })),
              ...(response.ManyToManyRelationships || []).map(row => ({
                ...row,
                relType: "N:N"
              }))
            ];

            const unique = new Map();

            rows.forEach(row => {
              const id = cleanGuid(row.MetadataId);
              if (!id || unique.has(id)) return;

              const related = [
                row.ReferencedEntity,
                row.ReferencingEntity,
                row.Entity1LogicalName,
                row.Entity2LogicalName
              ]
                .filter(Boolean)
                .filter((value, index, array) =>
                  array.indexOf(value) === index
                )
                .join(" ↔ ");

              unique.set(id, {
                id,
                name: row.SchemaName || "",
                displayName: row.relType,
                extra: related,
                raw: row
              });
            });

            return [...unique.values()];
          }

          if (type === "form") {
            const response = await apiGet(
              "systemforms?" +
              "$select=formid,name,uniquename,type,formactivationstate,objecttypecode" +
              `&$filter=objecttypecode eq '${logicalName}'` +
              "&$orderby=name asc"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.formid),
              name: row.name || row.uniquename || "",
              displayName: formatted(row, "type") || `Type ${row.type}`,
              extra:
                formatted(row, "formactivationstate") ||
                String(row.formactivationstate ?? ""),
              raw: row
            }));
          }

          if (type === "view") {
            const response = await apiGet(
              "savedqueries?" +
              "$select=savedqueryid,name,returnedtypecode,querytype,statecode" +
              `&$filter=returnedtypecode eq '${logicalName}'` +
              "&$orderby=name asc"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.savedqueryid),
              name: row.name || "",
              displayName:
                formatted(row, "querytype") ||
                `Query Type ${row.querytype}`,
              extra: row.statecode === 0 ? "Active" : "Inactive",
              raw: row
            }));
          }

          if (type === "chart") {
            const response = await apiGet(
              "savedqueryvisualizations?" +
              "$select=savedqueryvisualizationid,name,primaryentitytypecode" +
              `&$filter=primaryentitytypecode eq '${logicalName}'` +
              "&$orderby=name asc"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.savedqueryvisualizationid),
              name: row.name || "",
              displayName: "System Chart",
              extra: "",
              raw: row
            }));
          }

          if (type === "key") {
            const response = await apiGet(
              `EntityDefinitions(LogicalName='${logicalName}')/Keys?` +
              "$select=MetadataId,LogicalName,SchemaName,DisplayName,KeyAttributes"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.MetadataId),
              name: row.LogicalName || row.SchemaName || "",
              displayName: getLabel(row.DisplayName),
              extra: (row.KeyAttributes || []).join(", "),
              raw: row
            }));
          }

          if (type === "process") {
            const response = await apiGet(
              "workflows?" +
              "$select=workflowid,name,uniquename,category,primaryentity,statecode" +
              `&$filter=primaryentity eq '${logicalName}'` +
              "&$orderby=name asc"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.workflowid),
              name: row.name || row.uniquename || "",
              displayName:
                formatted(row, "category") ||
                `Category ${row.category}`,
              extra: row.statecode === 1 ? "Active" : "Draft",
              raw: row
            }));
          }

          return [];
        };

        const loadGlobalComponents = async familyKey => {
          const config = COMPONENT_FAMILIES[familyKey];

          if (config.mode === "globalChoice") {
            const response = await apiGet(
              "GlobalOptionSetDefinitions?" +
              "$select=MetadataId,Name,DisplayName"
            );

            return (response.value || []).map(row => ({
              id: cleanGuid(row.MetadataId),
              name: row.Name || "",
              displayName: getLabel(row.DisplayName),
              extra: "",
              raw: row
            }));
          }

          const rows = await apiGetAll(
            `${config.entitySet}?$select=${config.select.join(",")}`
          );

          return rows
            .filter(row => row[config.idField])
            .map(row => ({
              id: cleanGuid(row[config.idField]),
              name:
                config.nameFields
                  .map(field => row[field])
                  .find(Boolean) || "",
              displayName:
                config.nameFields
                  .slice(1)
                  .map(field => row[field])
                  .find(Boolean) || "",
              extra:
                row.assemblyname ||
                row.primaryentity ||
                row.schemaname ||
                formatted(row, "webresourcetype") ||
                formatted(row, "category") ||
                "",
              raw: row
            }));
        };

        const getSolutionComponents = async componentId => {
          const id = cleanGuid(componentId);

          if (!id) {
            return [];
          }

          const result = await apiGetAll(
            "solutioncomponents?" +
            "$select=objectid,componenttype" +
            `&$filter=objectid eq ${id}`
          );

          return result || [];
        };

        const getFallbackComponentTypes = component => {
          const familyKey = familySelect.value;
          const tableType = tableComponentSelect?.value;

          const familyTypes = {
            webResource: [61],
            pluginAssembly: [91],
            pluginType: [90],
            pluginStep: [92],
            workflow: [29],
            customApi: [29],
            modelDrivenApp: [80],
            siteMap: [62],
            globalChoice: [9],
            environmentVariableDefinition: [380],
            environmentVariableValue: [381],
            securityRole: [20],
            emailTemplate: [36]
          };

          if (familyKey !== "table") {
            return familyTypes[familyKey] || [];
          }

          const tableTypes = {
            table: [1],
            column: [2],
            relationship: [3, 10],
            form: [60, 24],
            view: [26],
            chart: [59],
            key: [14],
            localChoice: [2, 9],
            process: [29]
          };

          return tableTypes[tableType] || [];
        };

        const getAllLayers = async component => {
          const componentId = cleanGuid(component?.id);

          if (!componentId) {
            throw new Error("Component ID is missing.");
          }

          /*
           * The virtual table requires the real solution component type
           * together with the object ID. Query solutioncomponent first.
           */
          const solutionComponents =
            await getSolutionComponents(componentId);

          const typeCodes = new Set(
            solutionComponents
              .map(item => Number(item.componenttype))
              .filter(Number.isFinite)
          );

          /*
           * Some metadata subcomponents don't have a direct row in every
           * solution. Use the selected UI type only as a fallback.
           */
          if (!typeCodes.size) {
            getFallbackComponentTypes(component)
              .forEach(code => typeCodes.add(code));
          }

          const queries = [];
          const rows = [];

          for (const typeCode of typeCodes) {
            const typeName =
              COMPONENT_TYPE_NAME_MAP[typeCode];

            if (!typeName) {
              queries.push({
                objectId: componentId,
                componentType: typeCode,
                skipped: true,
                reason: "No component type name mapping"
              });
              continue;
            }

            const filter =
              `msdyn_solutioncomponentname eq '${escapeOData(typeName)}'` +
              " and " +
              `msdyn_componentid eq '${escapeOData(componentId)}'`;

            try {
              const result = await apiGetAll(
                "msdyn_componentlayers?" +
                "$select=" +
                [
                  "msdyn_componentlayerid",
                  "msdyn_componentid",
                  "msdyn_name",
                  "msdyn_order",
                  "msdyn_solutionname",
                  "msdyn_publishername",
                  "msdyn_solutioncomponentname",
                  "msdyn_overwritetime",
                  "msdyn_changes"
                ].join(",") +
                `&$filter=${encodeURIComponent(filter)}` +
                "&$orderby=msdyn_order asc"
              );

              queries.push({
                objectId: componentId,
                componentType: typeCode,
                componentTypeName: typeName,
                found: result.length
              });

              rows.push(...result);
            } catch (error) {
              queries.push({
                objectId: componentId,
                componentType: typeCode,
                componentTypeName: typeName,
                found: 0,
                error: error?.message || String(error)
              });
            }
          }

          const unique = new Map();

          rows.forEach(row => {
            const layerId =
              cleanGuid(row.msdyn_componentlayerid);

            const key =
              layerId ||
              [
                row.msdyn_componentid,
                row.msdyn_solutioncomponentname,
                row.msdyn_solutionname,
                row.msdyn_order,
                row.msdyn_overwritetime
              ].join("|");

            if (!unique.has(key)) {
              unique.set(key, {
                order: row.msdyn_order,
                solution:
                  row.msdyn_solutionname || "",
                publisher:
                  row.msdyn_publishername || "",
                componentType:
                  row.msdyn_solutioncomponentname || "",
                layerName:
                  row.msdyn_name || "",
                overwriteTime:
                  row.msdyn_overwritetime || "",
                hasChanges:
                  Boolean(row.msdyn_changes),
                componentId:
                  row.msdyn_componentid || "",
                layerId
              });
            }
          });

          const layers = [...unique.values()]
            .sort((a, b) => {
              const orderDifference =
                Number(a.order ?? 0) -
                Number(b.order ?? 0);

              if (orderDifference !== 0) {
                return orderDifference;
              }

              return String(a.overwriteTime || "")
                .localeCompare(
                  String(b.overwriteTime || "")
                );
            });

          console.log("Exact solution layer lookup:", {
            selectedComponent: component,
            solutionComponents,
            typeCodes: [...typeCodes],
            queries,
            layers
          });

          return layers;
        };

        const inspectComponent = async component => {
          const token = ++loadToken;

          setStatus(
            `⏳ Loading layers for ${
              component.displayName || component.name || component.id
            }...`
          );

          try {
            const layers = await getAllLayers(component);
            if (token !== loadToken) return;

            summary.innerHTML = `
              <b>Component:</b>
              ${component.displayName || component.name || "-"}
              &nbsp;&nbsp;|&nbsp;&nbsp;
              <b>GUID:</b>
              <span dir="ltr">${component.id}</span>
              &nbsp;&nbsp;|&nbsp;&nbsp;
              <b>Layers:</b> ${layers.length}
            `;

            renderLayers(layers);

            setStatus(
              layers.length
                ? `✅ Found ${layers.length} layer(s).`
                : "⚠️ No layers found for this component.",
              layers.length ? "success" : "warning"
            );

            console.table(layers);
          } catch (error) {
            console.error(error);
            renderLayers([]);
            setStatus(`❌ ${error.message || error}`, "error");
          }
        };

        const refreshComponents = async () => {
          const token = ++loadToken;
          const familyKey = familySelect.value;
          const family = COMPONENT_FAMILIES[familyKey];

          componentSearch.value = "";
          componentSelect.disabled = true;
          componentSelect.innerHTML = "<option>Loading...</option>";
          renderLayers([]);
          summary.textContent = "Select a component.";

          try {
            if (family.mode === "table") {
              if (!selectedEntity) {
                allComponents = [];
                filteredComponents = [];
                renderComponents();
                setStatus("Choose a table.", "warning");
                return;
              }

              allComponents = await loadTableComponents();
            } else {
              allComponents = await loadGlobalComponents(familyKey);
            }

            if (token !== loadToken) return;

            allComponents.sort((a, b) =>
              componentText(a).localeCompare(componentText(b))
            );

            filteredComponents = [...allComponents];
            renderComponents();
            componentSelect.disabled = false;

            setStatus(
              `✅ Loaded ${allComponents.length} component(s).`,
              "success"
            );

            if (allComponents.length === 1) {
              componentSelect.selectedIndex = 0;
              await inspectComponent(allComponents[0]);
            }
          } catch (error) {
            if (token !== loadToken) return;

            console.error(error);
            allComponents = [];
            filteredComponents = [];
            renderComponents();
            setStatus(`❌ ${error.message || error}`, "error");
          }
        };

        const updateMode = async () => {
          const tableMode =
            COMPONENT_FAMILIES[familySelect.value].mode === "table";

          entitySearchField.style.display = tableMode ? "grid" : "none";
          entityField.style.display = tableMode ? "grid" : "none";
          tableComponentField.style.display = tableMode ? "grid" : "none";

          await refreshComponents();
        };

        const close = () => overlay.remove();
        closeTop.onclick = close;
        closeButton.onclick = close;

        familySelect.addEventListener("change", updateMode);

        entitySearch.addEventListener("input", () => {
          const search = entitySearch.value.trim().toLowerCase();

          filteredEntities = !search
            ? [...allEntities]
            : allEntities.filter(entity =>
                [
                  entity.displayName,
                  entity.logicalName,
                  entity.schemaName
                ]
                  .filter(Boolean)
                  .some(value =>
                    value.toLowerCase().includes(search)
                  )
              );

          renderEntities();
        });

        entitySelect.addEventListener("change", async () => {
          selectedEntity =
            filteredEntities[Number(entitySelect.value)] || null;

          if (!selectedEntity) return;

          entitySearch.value =
            selectedEntity.displayName || selectedEntity.logicalName;

          await refreshComponents();
        });

        tableComponentSelect.addEventListener(
          "change",
          refreshComponents
        );

        componentSearch.addEventListener("input", () => {
          const search = componentSearch.value.trim().toLowerCase();

          filteredComponents = !search
            ? [...allComponents]
            : allComponents.filter(component =>
                [
                  component.displayName,
                  component.name,
                  component.extra,
                  component.id
                ]
                  .filter(Boolean)
                  .some(value =>
                    String(value).toLowerCase().includes(search)
                  )
              );

          renderComponents();
          renderLayers([]);
          summary.textContent = "Select a component.";
        });

        componentSelect.addEventListener("change", async () => {
          const component =
            filteredComponents[Number(componentSelect.value)] || null;

          if (component) {
            await inspectComponent(component);
          }
        });

        searchActiveLayersButton.addEventListener(
          "click",
          searchActiveLayers
        );

        renderLayers([]);

        try {
          setStatus("⏳ Loading tables...");
          await loadEntities();
          await updateMode();
        } catch (error) {
          console.error(error);
          setStatus(`❌ ${error.message || error}`, "error");
        }
      }
    });
  });


document
  .getElementById("openSystemUserUi")
  .addEventListener("click", async () => {
    const tab = await __d365GetActiveTab();
    if (!tab?.id) return;

    await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: false },
      world: "MAIN",
      func: async () => {
        const MODAL_ID = "__rv_open_systemuser";
        document.getElementById(MODAL_ID)?.remove();

        const globalContext =
          window.Xrm?.Utility?.getGlobalContext?.();

        const clientUrl = globalContext?.getClientUrl?.();

        if (!clientUrl) {
          alert("D365 context not found.");
          return;
        }

        const cleanGuid = (value) =>
          String(value || "")
            .replace(/[{}]/g, "")
            .trim()
            .toLowerCase();

        const escapeOData = (value) =>
          String(value || "").replace(/'/g, "''");

        const html = (value) =>
          String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");

        const requestJson = async (url) => {
          const response = await fetch(url, {
            method: "GET",
            credentials: "same-origin",
            headers: {
              Accept: "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0",
              Prefer:
                'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
            }
          });

          const text = await response.text();
          let body = null;

          try {
            body = text ? JSON.parse(text) : null;
          } catch {
            body = text;
          }

          if (!response.ok) {
            throw new Error(
              body?.error?.message ||
                `HTTP ${response.status} ${response.statusText}`
            );
          }

          return body;
        };

        const openUser = async (userId) => {
          const id = cleanGuid(userId);
          if (!id) return;

          try {
            await window.Xrm.Navigation.openForm({
              entityName: "systemuser",
              entityId: id,
              openInNewWindow: true
            });
          } catch (error) {
            console.error(error);
            alert(error?.message || "Could not open the user record.");
          }
        };

        const style = document.createElement("style");
        style.textContent = `
          #${MODAL_ID}, #${MODAL_ID} * { box-sizing:border-box; }

          #${MODAL_ID} {
            position:fixed;
            inset:0;
            z-index:2147483647;
            display:flex;
            align-items:center;
            justify-content:center;
            padding:18px;
            direction:rtl;
            background:rgba(15,23,42,.62);
            backdrop-filter:blur(5px);
            font-family:system-ui,-apple-system,"Segoe UI",Arial,sans-serif;
          }

          #${MODAL_ID} .osu-dialog {
            width:min(900px,96vw);
            max-height:min(760px,94vh);
            display:flex;
            flex-direction:column;
            overflow:hidden;
            color:#111827;
            background:#ffffff;
            border:1px solid #cbd5e1;
            border-radius:18px;
            box-shadow:0 28px 90px rgba(0,0,0,.34);
          }

          #${MODAL_ID} .osu-header {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:14px;
            padding:17px 19px;
            background:#ffffff;
            border-bottom:1px solid #e2e8f0;
          }

          #${MODAL_ID} .osu-title {
            color:#0f172a;
            font-size:20px;
            font-weight:950;
          }

          #${MODAL_ID} .osu-subtitle {
            margin-top:3px;
            color:#64748b;
            font-size:12px;
          }

          #${MODAL_ID} .osu-close {
            width:40px;
            height:40px;
            padding:0;
            cursor:pointer;
            color:#111827;
            background:#ffffff;
            border:1px solid #cbd5e1;
            border-radius:11px;
            font-size:17px;
            font-weight:900;
          }

          #${MODAL_ID} .osu-body {
            display:flex;
            flex-direction:column;
            gap:14px;
            min-height:0;
            padding:16px;
            overflow:auto;
            background:#f8fafc;
          }

          #${MODAL_ID} .osu-me-card {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:14px;
            padding:14px;
            background:#ffffff;
            border:1px solid #dbe3ec;
            border-radius:14px;
          }

          #${MODAL_ID} .osu-me-title {
            color:#0f172a;
            font-size:14px;
            font-weight:950;
          }

          #${MODAL_ID} .osu-me-value {
            margin-top:4px;
            color:#64748b;
            font-size:12px;
            direction:ltr;
            text-align:right;
          }

          #${MODAL_ID} .osu-search-panel {
            padding:14px;
            background:#ffffff;
            border:1px solid #dbe3ec;
            border-radius:14px;
          }

          #${MODAL_ID} .osu-label {
            margin-bottom:7px;
            color:#334155;
            font-size:12px;
            font-weight:900;
          }

          #${MODAL_ID} .osu-search-row {
            display:grid;
            grid-template-columns:minmax(0,1fr) auto;
            gap:9px;
          }

          #${MODAL_ID} .osu-input {
            width:100%;
            min-width:0;
            padding:11px 12px;
            color:#111827;
            background:#ffffff;
            border:1px solid #b8c3d1;
            border-radius:10px;
            outline:none;
            font-size:13px;
          }

          #${MODAL_ID} .osu-input:focus {
            border-color:#0f172a;
            box-shadow:0 0 0 2px rgba(15,23,42,.09);
          }

          #${MODAL_ID} .osu-help {
            margin-top:7px;
            color:#64748b;
            font-size:11px;
          }

          #${MODAL_ID} .osu-btn {
            padding:10px 14px;
            cursor:pointer;
            color:#111827;
            background:#ffffff;
            border:1px solid #b8c3d1;
            border-radius:10px;
            font-size:12px;
            font-weight:900;
          }

          #${MODAL_ID} .osu-btn:hover {
            background:#f1f5f9;
          }

          #${MODAL_ID} .osu-primary {
            color:#ffffff;
            background:#111827;
            border-color:#111827;
          }

          #${MODAL_ID} .osu-primary:hover {
            background:#000000;
          }

          #${MODAL_ID} .osu-status {
            min-height:20px;
            color:#64748b;
            font-size:12px;
          }

          #${MODAL_ID} .osu-results {
            display:flex;
            flex-direction:column;
            gap:8px;
          }

          #${MODAL_ID} .osu-user {
            width:100%;
            display:grid;
            grid-template-columns:minmax(180px,1.25fr) minmax(180px,1fr) minmax(160px,.9fr) auto;
            gap:12px;
            align-items:center;
            padding:12px 13px;
            cursor:pointer;
            text-align:right;
            color:#111827;
            background:#ffffff;
            border:1px solid #dbe3ec;
            border-radius:12px;
          }

          #${MODAL_ID} .osu-user:hover {
            border-color:#94a3b8;
            box-shadow:0 7px 18px rgba(15,23,42,.07);
            transform:translateY(-1px);
          }

          #${MODAL_ID} .osu-name {
            overflow:hidden;
            text-overflow:ellipsis;
            white-space:nowrap;
            color:#0f172a;
            font-size:13px;
            font-weight:950;
          }

          #${MODAL_ID} .osu-domain,
          #${MODAL_ID} .osu-email,
          #${MODAL_ID} .osu-bu {
            overflow:hidden;
            text-overflow:ellipsis;
            white-space:nowrap;
            color:#475569;
            font-size:11px;
            direction:ltr;
            text-align:left;
          }

          #${MODAL_ID} .osu-badge {
            display:inline-flex;
            align-items:center;
            justify-content:center;
            padding:4px 8px;
            border-radius:999px;
            font-size:10px;
            font-weight:900;
            white-space:nowrap;
          }

          #${MODAL_ID} .osu-active {
            color:#047857;
            background:#ecfdf5;
            border:1px solid #a7f3d0;
          }

          #${MODAL_ID} .osu-disabled {
            color:#b91c1c;
            background:#fef2f2;
            border:1px solid #fecaca;
          }

          #${MODAL_ID} .osu-empty {
            padding:28px;
            color:#64748b;
            text-align:center;
            background:#ffffff;
            border:1px dashed #cbd5e1;
            border-radius:12px;
          }

          @media (max-width:700px) {
            #${MODAL_ID} {
              padding:0;
            }

            #${MODAL_ID} .osu-dialog {
              width:100vw;
              max-height:100vh;
              height:100vh;
              border-radius:0;
            }

            #${MODAL_ID} .osu-me-card {
              align-items:stretch;
              flex-direction:column;
            }

            #${MODAL_ID} .osu-search-row,
            #${MODAL_ID} .osu-user {
              grid-template-columns:1fr;
            }

            #${MODAL_ID} .osu-btn {
              width:100%;
            }
          }
        `;

        document.head.appendChild(style);

        const currentUserId = cleanGuid(
          globalContext?.userSettings?.userId
        );

        const currentUserName =
          globalContext?.userSettings?.userName || "Current user";

        const overlay = document.createElement("div");
        overlay.id = MODAL_ID;
        overlay.innerHTML = `
          <div class="osu-dialog">
            <div class="osu-header">
              <div>
                <div class="osu-title">Open System User</div>
                <div class="osu-subtitle">Open the logged-in user or search by name, domain or email</div>
              </div>
              <button class="osu-close" id="osuClose">✕</button>
            </div>

            <div class="osu-body">
              <div class="osu-me-card">
                <div>
                  <div class="osu-me-title">Logged-in user</div>
                  <div class="osu-me-value">${html(currentUserName)} · ${html(currentUserId)}</div>
                </div>
                <button class="osu-btn osu-primary" id="osuOpenMe">Open my user</button>
              </div>

              <div class="osu-search-panel">
                <div class="osu-label">Search system users</div>
                <div class="osu-search-row">
                  <input
                    class="osu-input"
                    id="osuSearch"
                    placeholder="Name, domain, email or user GUID"
                    autocomplete="off"
                  />
                  <button class="osu-btn osu-primary" id="osuSearchButton">Search</button>
                </div>
                <div class="osu-help">
                  Searches fullname, domainname and internalemailaddress. Pasting a GUID opens the user directly.
                </div>
              </div>

              <div class="osu-status" id="osuStatus">Enter at least 2 characters.</div>
              <div class="osu-results" id="osuResults"></div>
            </div>
          </div>
        `;

        document.body.appendChild(overlay);

        const searchInput = overlay.querySelector("#osuSearch");
        const searchButton = overlay.querySelector("#osuSearchButton");
        const results = overlay.querySelector("#osuResults");
        const status = overlay.querySelector("#osuStatus");

        const setStatus = (message, type = "normal") => {
          const colors = {
            normal: "#64748b",
            success: "#047857",
            warning: "#b45309",
            error: "#b91c1c"
          };

          status.textContent = message;
          status.style.color = colors[type] || colors.normal;
        };

        const isGuid = (value) =>
          /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
            cleanGuid(value)
          );

        const renderUsers = (users) => {
          results.innerHTML = "";

          if (!users.length) {
            results.innerHTML = `
              <div class="osu-empty">No matching system users were found.</div>
            `;
            return;
          }

          users.forEach((user) => {
            const businessUnit =
              user[
                "_businessunitid_value@OData.Community.Display.V1.FormattedValue"
              ] || "";

            const button = document.createElement("button");
            button.type = "button";
            button.className = "osu-user";
            button.innerHTML = `
              <div>
                <div class="osu-name" title="${html(user.fullname || "")}">
                  ${html(user.fullname || "(No name)")}
                </div>
                <div class="osu-bu" title="${html(businessUnit)}">
                  ${html(businessUnit || "No business unit")}
                </div>
              </div>

              <div class="osu-domain" title="${html(user.domainname || "")}">
                ${html(user.domainname || "No domain")}
              </div>

              <div class="osu-email" title="${html(user.internalemailaddress || "")}">
                ${html(user.internalemailaddress || "No email")}
              </div>

              <span class="osu-badge ${user.isdisabled ? "osu-disabled" : "osu-active"}">
                ${user.isdisabled ? "Disabled" : "Active"}
              </span>
            `;

            button.addEventListener("click", () =>
              openUser(user.systemuserid)
            );

            results.appendChild(button);
          });
        };

        const searchUsers = async () => {
          const value = searchInput.value.trim();

          if (!value) {
            results.innerHTML = "";
            setStatus("Enter a name, domain, email or user GUID.", "warning");
            return;
          }

          if (isGuid(value)) {
            setStatus("Opening user record...", "success");
            await openUser(value);
            return;
          }

          if (value.length < 2) {
            results.innerHTML = "";
            setStatus("Enter at least 2 characters.", "warning");
            return;
          }

          searchButton.disabled = true;
          searchButton.textContent = "Searching...";
          results.innerHTML = `<div class="osu-empty">Loading users...</div>`;
          setStatus("Searching system users...");

          try {
            const safeValue = escapeOData(value);
            const filter = [
              `contains(fullname,'${safeValue}')`,
              `contains(domainname,'${safeValue}')`,
              `contains(internalemailaddress,'${safeValue}')`
            ].join(" or ");

            const query =
              `${clientUrl}/api/data/v9.2/systemusers?` +
              `$select=systemuserid,fullname,domainname,internalemailaddress,isdisabled,_businessunitid_value&` +
              `$filter=${encodeURIComponent(filter)}&` +
              `$orderby=fullname asc&` +
              `$top=50`;

            const response = await requestJson(query);
            const users = response?.value || [];

            renderUsers(users);
            setStatus(
              users.length
                ? `${users.length} user(s) found. Click a row to open the record.`
                : "No users found.",
              users.length ? "success" : "warning"
            );
          } catch (error) {
            console.error(error);
            results.innerHTML = `
              <div class="osu-empty" style="color:#b91c1c">
                ${html(error?.message || "Search failed.")}
              </div>
            `;
            setStatus(error?.message || "Search failed.", "error");
          } finally {
            searchButton.disabled = false;
            searchButton.textContent = "Search";
          }
        };

        overlay
          .querySelector("#osuClose")
          .addEventListener("click", () => overlay.remove());

        overlay
          .querySelector("#osuOpenMe")
          .addEventListener("click", () => openUser(currentUserId));

        searchButton.addEventListener("click", searchUsers);

        searchInput.addEventListener("keydown", (event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            searchUsers();
          }
        });

        searchInput.focus();
      }
    });
  });


