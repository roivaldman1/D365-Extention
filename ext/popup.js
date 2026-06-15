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

document.getElementById("retrieveMultiple").addEventListener("click", async () => {
  const tab = await __d365GetActiveTab();
  if (!tab?.id) return;

  await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: false },
    world: "MAIN",
    func: () => {
      // ---------- helpers ----------
      const normalizeFilter = (f) => {
        if (!f) return "";
        f = f.replace(/[“”]/g, '"').replace(/[‘’]/g, "'");
        f = f.replace(/"([^"]*)"/g, "'$1'");
        return f.trim();
      };

      const safeString = (v) => {
        if (v == null) return "";
        if (typeof v === "string") return v;
        if (typeof v === "number" || typeof v === "boolean") return String(v);
        try { return JSON.stringify(v); } catch (e) { return String(v); }
      };

      const getShownVal = (row, key) => {
        const fv = row?.[`${key}@OData.Community.Display.V1.FormattedValue`];
        const v = (fv != null) ? fv : row?.[key];
        return v;
      };

      const escapePipes = (s) => String(s ?? "").replace(/\|/g, "\\|");

      // Parses the "Columns" input:
      // - allows plain: col1,col2,col3
      // - allows advanced: col1,col2&$expand=nav($select=name)&$orderby=createdon desc
      const parseColumnsAndExtra = (input) => {
        const raw = (input || "").trim();
        if (!raw) return { cols: [], extraParts: [] };

        // Split on first '&' (keep the rest as extra query)
        const ampIndex = raw.indexOf("&");
        let colsPart = raw;
        let extra = "";

        if (ampIndex !== -1) {
          colsPart = raw.slice(0, ampIndex).trim();
          extra = raw.slice(ampIndex + 1).trim(); // everything after &
        }

        const cols = colsPart
          .split(",")
          .map(s => s.trim())
          .filter(Boolean);

        const extraParts = extra
          ? extra.split("&").map(s => s.trim()).filter(Boolean)
          : [];

        return { cols, extraParts };
      };

      // ---------- modal ----------
      document.getElementById("__d365helper_modal")?.remove();

      const overlay = document.createElement("div");
      overlay.id = "__d365helper_modal";
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,.35);
        z-index: 2147483647; display: flex; align-items: center; justify-content: center; padding: 16px;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(980px, 96vw); background: #fff; border-radius: 14px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35); overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      const header = document.createElement("div");
      header.style.cssText = `padding: 12px 14px; font-weight: 800; border-bottom: 1px solid #e5e7eb;`;
      header.textContent = "D365 RetrieveMultiple (WebApi)";

      const body = document.createElement("div");
      body.style.cssText = `padding: 12px 14px; display: grid; gap: 10px;`;

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

      const selectInput = document.createElement("input");
      selectInput.placeholder =
        "Columns (comma). You can also append: & $expand=... & $orderby=...  (example: col1,col2&$expand=nav($select=name))";
      selectInput.style.cssText = inputStyle;

      const filterInput = document.createElement("input");
      filterInput.placeholder = "Filter (without $filter=) e.g. statecode eq 0 and contains(fullname,'Roi')";
      filterInput.style.cssText = inputStyle;

      const topInput = document.createElement("input");
      topInput.placeholder = "Top (optional) e.g. 25";
      topInput.style.cssText = inputStyle;

      const status = document.createElement("div");
      status.style.cssText = `font-size: 12px; color: #374151;`;

      const resultTa = document.createElement("textarea");
      resultTa.readOnly = true;
      resultTa.placeholder = "Results will appear here…";
      resultTa.style.cssText = `
        width: 100%;
        height: 360px;
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

      body.appendChild(mkRow("Entity", entityInput));
      body.appendChild(mkRow("Columns (+ optional & $expand=...)", selectInput));
      body.appendChild(mkRow("Filter (optional)", filterInput));
      body.appendChild(mkRow("Top (optional)", topInput));
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

      const btnRun = btn("Run");
      btnRun.style.border = "none";
      btnRun.style.background = "#111827";
      btnRun.style.color = "#fff";

      const close = () => overlay.remove();
      btnClose.onclick = close;
      

      btnCopy.onclick = async () => {
        try {
          await navigator.clipboard.writeText(resultTa.value || "");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        } catch (e) {
          resultTa.focus(); resultTa.select(); document.execCommand("copy");
          btnCopy.textContent = "Copied ✅";
          setTimeout(() => (btnCopy.textContent = "Copy"), 900);
        }
      };

      btnRun.onclick = async () => {
        const entity = (entityInput.value || "").trim();
        const filter = normalizeFilter(filterInput.value || "");
        const topStr = (topInput.value || "").trim();

        status.textContent = "";
        resultTa.value = "";

        if (!entity) { status.textContent = "❌ Entity is required."; return; }

        const { cols, extraParts } = parseColumnsAndExtra(selectInput.value || "");

        const top = topStr ? parseInt(topStr, 10) : null;
        if (topStr && (!Number.isFinite(top) || top <= 0)) {
          status.textContent = "❌ Top must be a positive number.";
          return;
        }

        const Xrm = window.Xrm;
        const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;
        if (!webApi?.retrieveMultipleRecords) {
          status.textContent = "❌ Xrm.WebApi.retrieveMultipleRecords not available.";
          return;
        }

        // build query string (supports $expand and any extra $... parts)
        const params = [];
        if (cols.length > 0) params.push(`$select=${encodeURIComponent(cols.join(","))}`);
        if (filter) params.push(`$filter=${encodeURIComponent(filter)}`);
        if (top) params.push(`$top=${encodeURIComponent(String(top))}`);

        // Append extra parts like: $expand=... or $orderby=... or $select=...
        // NOTE: keep it generic, but safe (encode key+value)
        for (const p of extraParts) {
          const part = p.replace(/^\?/, "").trim();
          if (!part) continue;

          // Allow "$expand=..." (most common)
          const eqIdx = part.indexOf("=");
          if (eqIdx === -1) {
            params.push(encodeURIComponent(part));
            continue;
          }

          const key = part.slice(0, eqIdx).trim();
          const val = part.slice(eqIdx + 1).trim();
          if (!key) continue;

          params.push(`${encodeURIComponent(key)}=${encodeURIComponent(val)}`);
        }

        const query = params.length ? `?${params.join("&")}` : "";

        status.textContent = "⏳ Running…";

        try {
          const res = await webApi.retrieveMultipleRecords(entity, query);
          const rows = res?.entities || [];

          const lines = [];
          lines.push(`Entity: ${entity}`);
          lines.push(`Query: ${query || "(none)"}  (no $select => ALL columns)`);
          lines.push(`Returned: ${rows.length}`);
          lines.push("");

          if (!rows.length) {
            lines.push("(no rows)");
            resultTa.value = lines.join("\n");
            status.textContent = "✅ Done (0 rows).";
            return;
          }

          // Columns to show:
          // - if user chose columns => show them
          // - else show keys from first row (up to 25)
          const shownCols = (cols.length > 0)
            ? cols
            : Object.keys(rows[0]).filter(k => !k.startsWith("@")).slice(0, 25);

          // Pretty table with dynamic widths (better reading)
          const colWidths = shownCols.map((c) => {
            const headerW = c.length;
            const maxCell = Math.max(
              ...rows.slice(0, 200).map(r => safeString(getShownVal(r, c)).length)
            );
            return Math.min(Math.max(headerW, maxCell, 6), 40); // cap width 40
          });

          const pad = (s, w) => {
            s = safeString(s);
            if (s.length > w) return s.slice(0, Math.max(0, w - 1)) + "…";
            return (s + " ".repeat(w)).slice(0, w);
          };

          lines.push(
            shownCols.map((c, i) => pad(escapePipes(c), colWidths[i])).join(" | ")
          );
          lines.push(
            shownCols.map((_, i) => "-".repeat(colWidths[i])).join("-+-")
          );

          for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const vals = shownCols.map((c, idx) => {
              const v = getShownVal(r, c);
              return pad(escapePipes(v), colWidths[idx]);
            });
            lines.push(vals.join(" | "));
            if (i >= 199) { lines.push("... (truncated to 200 rows)"); break; }
          }

          // JSON dump ALL rows (for discovery) - cap to 50 for sanity
          lines.push("\n--- ALL RECORDS (JSON, up to 50) ---\n");
          try {
            const take = rows.slice(0, 50);
            lines.push(JSON.stringify(take, null, 2));
            if (rows.length > 50) lines.push(`\n... (${rows.length - 50} more not shown)`);
          } catch (e) {
            lines.push(String(rows));
          }

          resultTa.value = lines.join("\n");
          status.textContent = `✅ Done (${rows.length} rows).`;
          resultTa.focus();
          resultTa.select();
        } catch (err) {
          status.textContent = "❌ Failed.";
          resultTa.value =
            "ERROR:\n" +
            (err?.message || err?.toString?.() || "Unknown error") +
            "\n\nTip: Put ONLY columns in Columns. If you add expand, append like: & $expand=nav($select=name)\n" +
            "Tip: filter must be valid OData. For strings use single quotes: firstname eq 'Roi'";
        }
      };

      footer.appendChild(btnClose);
      footer.appendChild(btnCopy);
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
            <input id="__rolesUserSearch" type="text" placeholder="חפש לפי שם / יוזר / מייל"
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

            <select id="__rolesRoleSelect" size="8"
              style="width:100%;padding:8px;border:1px solid #ccc;border-radius:10px;box-sizing:border-box;"></select>

            <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;">
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

      const closeModal = () => overlay.remove();
      document.getElementById("__rolesClose").onclick = closeModal;

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

      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const BASE_URL = `${clientUrl}/api/data/v9.2`;

      let allUsers = [];
      let allRoles = [];
      let businessUnitsMap = {};
      let currentSelectedUserId = null;

      function normalizeGuid(id) {
        return String(id || "").replace(/[{}]/g, "").trim();
      }

      function escapeODataString(str) {
        return String(str ?? "").replace(/'/g, "''");
      }

      async function fetchJSON(url, options = {}) {
        const res = await fetch(url, {
          ...options,
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json",
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
        let all = [];
        let nextUrl = url;

        while (nextUrl) {
          const data = await fetchJSON(nextUrl);
          all = all.concat(data?.value || []);
          nextUrl = data?.["@odata.nextLink"] || null;
        }

        return all;
      }

      async function loadUsers() {
        const users = await fetchAllPages(
          `${BASE_URL}/systemusers?$select=systemuserid,fullname,domainname,isdisabled,internalemailaddress&$orderby=fullname asc`
        );

        allUsers = users.map(u => {
          const domain = u.domainname || "";
          const username = domain ? domain.replace(/@mac\.org\.il$/i, "") : "";

          return {
            id: u.systemuserid,
            fullname: u.fullname || "",
            domainname: domain,
            internalemailaddress: u.internalemailaddress || "",
            isdisabled: !!u.isdisabled,
            username
          };
        });
      }

      async function loadBusinessUnitsMap() {
        const businessUnits = await fetchAllPages(
          `${BASE_URL}/businessunits?$select=businessunitid,name,_parentbusinessunitid_value&$orderby=name asc`
        );

        businessUnitsMap = {};
        for (const bu of businessUnits) {
          businessUnitsMap[bu.businessunitid] = bu.name || "";
        }

        return businessUnits;
      }

      async function loadRoles() {
        await loadBusinessUnitsMap();

        const roles = await fetchAllPages(
          `${BASE_URL}/roles?$select=roleid,name,_businessunitid_value&$orderby=name asc`
        );

        allRoles = roles
          .filter(r => r.name)
          .map(r => ({
            id: r.roleid,
            name: r.name || "",
            businessunitid: r._businessunitid_value || ""
          }))
          .sort((a, b) => {
            const nameCmp = (a.name || "").localeCompare(b.name || "", "he");
            if (nameCmp !== 0) return nameCmp;
            return (businessUnitsMap[a.businessunitid] || "").localeCompare(businessUnitsMap[b.businessunitid] || "", "he");
          });
      }

      function getCurrentUserId() {
        return normalizeGuid(Xrm.Utility.getGlobalContext().userSettings.userId);
      }

      function getSelectedUserId() {
        if (useCurrentUserCheckbox.checked) return getCurrentUserId();

        const selectedUser = userSelect.options[userSelect.selectedIndex];
        if (!selectedUser) throw new Error("צריך לבחור משתמש מהרשימה או לסמן 'בצע עליי'.");

        return selectedUser.dataset.userid;
      }

      function getSelectedUserLabel() {
        if (useCurrentUserCheckbox.checked) {
          return "משתמש נבחר: המשתמש המחובר";
        }

        const selectedUser = userSelect.options[userSelect.selectedIndex];
        if (!selectedUser) return "לא נבחר משתמש";

        return `משתמש נבחר: ${selectedUser.dataset.fullname || ""} | ${selectedUser.dataset.domainname || "ללא domain"} | ${selectedUser.dataset.isdisabled === "true" ? "לא פעיל" : "פעיל"}`;
      }

      function toggleUserSelectionState() {
        const disabled = useCurrentUserCheckbox.checked;

        userSearchInput.disabled = disabled;
        userSelect.disabled = disabled;
        userSearchInput.style.opacity = disabled ? "0.6" : "1";
        userSelect.style.opacity = disabled ? "0.6" : "1";
      }

      function renderUsers(searchText = "") {
        const q = searchText.trim().toLowerCase();
        userSelect.innerHTML = "";

        const filtered = allUsers.filter(u => {
          if (!q) return true;
          return (
            (u.fullname || "").toLowerCase().includes(q) ||
            (u.domainname || "").toLowerCase().includes(q) ||
            (u.username || "").toLowerCase().includes(q) ||
            (u.internalemailaddress || "").toLowerCase().includes(q)
          );
        });

        for (const user of filtered) {
          const option = document.createElement("option");
          option.value = user.id;
          option.textContent = `${user.fullname || "(ללא שם)"} | ${user.username || user.domainname || "ללא domain"} | ${user.isdisabled ? "לא פעיל" : "פעיל"}`;
          option.dataset.userid = user.id;
          option.dataset.fullname = user.fullname || "";
          option.dataset.domainname = user.domainname || "";
          option.dataset.email = user.internalemailaddress || "";
          option.dataset.isdisabled = String(user.isdisabled);
          userSelect.appendChild(option);
        }

        if (filtered.length > 0) userSelect.selectedIndex = 0;
      }

      function renderRoles(searchText = "", buFilter = "") {
        const q = searchText.trim().toLowerCase();
        roleSelect.innerHTML = "";

        const filtered = allRoles.filter(r => {
          if (buFilter && r.businessunitid !== buFilter) return false;
          if (!q) return true;
          return (r.name || "").toLowerCase().includes(q);
        });

        for (const role of filtered) {
          const buName = businessUnitsMap[role.businessunitid] || "";
          const option = document.createElement("option");
          option.value = role.name;
          option.textContent = buName ? `${role.name} | ${buName}` : role.name;
          option.dataset.roleid = role.id;
          roleSelect.appendChild(option);
        }

        if (filtered.length > 0) roleSelect.selectedIndex = 0;
      }

      function populateBuFilter() {
        const buIds = [...new Set(allRoles.map(r => r.businessunitid))];
        const buList = buIds
          .map(id => ({ id, name: businessUnitsMap[id] || "" }))
          .filter(bu => bu.name)
          .sort((a, b) => a.name.localeCompare(b.name, "he"));

        for (const bu of buList) {
          const opt = document.createElement("option");
          opt.value = bu.id;
          opt.textContent = bu.name;
          buFilterSelect.appendChild(opt);
        }
      }

      async function getUserRoles(userId) {
        const userUrl =
          `${BASE_URL}/systemusers(${userId})` +
          `?$select=fullname,domainname,isdisabled` +
          `&$expand=systemuserroles_association($select=roleid,name,_businessunitid_value),` +
          `teammembership_association($select=teamid,name)`;

        const userData = await fetchJSON(userUrl);

        const directRoles = (userData.systemuserroles_association || []).map(r => ({
          roleid: r.roleid,
          name: r.name || "",
          businessunitid: r._businessunitid_value || "",
          source: "ישיר"
        }));

        const teams = userData.teammembership_association || [];
        const teamRolesArrays = await Promise.all(teams.map(async team => {
          try {
            const teamData = await fetchJSON(
              `${BASE_URL}/teams(${team.teamid})?$select=name` +
              `&$expand=teamroles_association($select=roleid,name,_businessunitid_value)`
            );
            return (teamData.teamroles_association || []).map(r => ({
              roleid: r.roleid,
              name: r.name || "",
              businessunitid: r._businessunitid_value || "",
              source: `צוות: ${team.name || ""}`
            }));
          } catch {
            return [];
          }
        }));

        const roleMap = new Map();
        for (const r of [...directRoles, ...teamRolesArrays.flat()]) {
          if (!roleMap.has(r.roleid)) {
            roleMap.set(r.roleid, r);
          }
        }

        return {
          user: {
            fullname: userData.fullname || "",
            domainname: userData.domainname || "",
            isdisabled: !!userData.isdisabled
          },
          roles: Array.from(roleMap.values())
        };
      }

      async function addUserRole(userId, roleId) {
        const roleData = await fetchJSON(`${BASE_URL}/roles(${roleId})?$select=roleid,name`);

        await fetchJSON(`${BASE_URL}/systemusers(${userId})/systemuserroles_association/$ref`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            "@odata.id": `${BASE_URL}/roles(${roleData.roleid})`
          })
        });

        return roleData;
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
          const nameCompare = (a.name || "").localeCompare(b.name || "", "he");
          if (nameCompare !== 0) return nameCompare;
          return (businessUnitsMap[a.businessunitid] || "").localeCompare(businessUnitsMap[b.businessunitid] || "", "he");
        });

        for (const role of sortedRoles) {
          const row = document.createElement("div");
          row.style.cssText = `
            display:grid;
            grid-template-columns:1.5fr 1fr 140px;
            padding:12px;
            border-bottom:1px solid #eee;
            align-items:center;
          `;

          const roleNameCell = document.createElement("div");
          const sourceLabel = role.source && role.source !== "ישיר"
            ? ` <span style="font-size:12px;color:#666;">(${role.source})</span>`
            : "";
          roleNameCell.innerHTML = `${role.name || ""}${sourceLabel}`;

          const buCell = document.createElement("div");
          buCell.textContent = businessUnitsMap[role.businessunitid] || role.businessunitid || "";

          const actionCell = document.createElement("div");

          if (role.source === "ישיר") {
            const removeBtn = document.createElement("button");
            removeBtn.textContent = "הסר";
            removeBtn.style.cssText = `
              border:none;
              background:#d13438;
              color:white;
              border-radius:8px;
              padding:8px 12px;
              cursor:pointer;
              font-size:14px;
            `;

            removeBtn.addEventListener("click", async () => {
              if (!confirm(`להסיר את התפקיד "${role.name}" מהמשתמש?`)) return;

              removeBtn.disabled = true;
              statusBox.textContent = `מסיר את התפקיד "${role.name}"...`;

              try {
                await removeUserRole(userId, role.roleid);
                statusBox.textContent = `✅ התפקיד "${role.name}" הוסר בהצלחה`;
                await refreshUserRoles(userId, selectedUserBox.textContent);
              } catch (err) {
                statusBox.textContent = `❌ שגיאה בהסרה: ${err.message}`;
              } finally {
                removeBtn.disabled = false;
              }
            });

            actionCell.appendChild(removeBtn);
          } else {
            const span = document.createElement("span");
            span.style.cssText = "font-size:12px;color:#888;";
            span.textContent = "דרך צוות";
            actionCell.appendChild(span);
          }

          row.appendChild(roleNameCell);
          row.appendChild(buCell);
          row.appendChild(actionCell);
          rowsBox.appendChild(row);
        }
      }

      async function refreshUserRoles(userId, userLabel) {
        currentSelectedUserId = userId;
        rowsBox.innerHTML = "";
        selectedUserBox.textContent = userLabel || getSelectedUserLabel();
        statusBox.textContent = "טוען תפקידי אבטחה...";

        const result = await getUserRoles(userId);
        renderRolesRows(result.roles, userId);

        statusBox.textContent =
          `✅ נמצאו ${result.roles.length} תפקידי אבטחה עבור ${result.user.fullname}` +
          (result.user.domainname ? ` (${result.user.domainname})` : "");
      }

      try {
        statusBox.textContent = "טוען משתמשים, תפקידים ויחידות עסקיות...";
        await Promise.all([loadUsers(), loadRoles()]);
        renderUsers();
        renderRoles();
        populateBuFilter();
        statusBox.textContent = `✅ נטענו ${allUsers.length} משתמשים ו-${allRoles.length} תפקידים`;
      } catch (err) {
        statusBox.textContent = `❌ שגיאה בטעינה: ${err.message}`;
        return;
      }

      useCurrentUserCheckbox.addEventListener("change", toggleUserSelectionState);
      toggleUserSelectionState();

      userSearchInput.addEventListener("input", () => renderUsers(userSearchInput.value));
      roleSearchInput.addEventListener("input", () => renderRoles(roleSearchInput.value, buFilterSelect.value));
      buFilterSelect.addEventListener("change", () => renderRoles(roleSearchInput.value, buFilterSelect.value));

      clearBtn.addEventListener("click", () => {
        userSearchInput.value = "";
        roleSearchInput.value = "";
        buFilterSelect.value = "";
        renderUsers();
        renderRoles();
        currentSelectedUserId = null;
        selectedUserBox.textContent = "לא נבחר משתמש";
        rowsBox.innerHTML = "";
        statusBox.textContent = `✅ נטענו ${allUsers.length} משתמשים ו-${allRoles.length} תפקידים`;
      });

      showBtn.addEventListener("click", async () => {
        showBtn.disabled = true;

        try {
          const userId = getSelectedUserId();
          await refreshUserRoles(userId, getSelectedUserLabel());
        } catch (err) {
          statusBox.textContent = `❌ שגיאה: ${err.message}`;
        } finally {
          showBtn.disabled = false;
        }
      });

      addBtn.addEventListener("click", async () => {
        const selectedRole = roleSelect.options[roleSelect.selectedIndex];

        if (!selectedRole) {
          statusBox.textContent = "צריך לבחור תפקיד מהרשימה.";
          return;
        }

        addBtn.disabled = true;

        try {
          const userId = getSelectedUserId();
          const roleId = selectedRole.dataset.roleid;
          const roleName = selectedRole.value;

          statusBox.textContent = `מקצה את התפקיד "${roleName}"...`;

          const role = await addUserRole(userId, roleId);

          statusBox.textContent = `✅ התפקיד "${role.name}" הוקצה בהצלחה`;
          await refreshUserRoles(userId, getSelectedUserLabel());
        } catch (err) {
          statusBox.textContent = `❌ שגיאה בהוספה: ${err.message}`;
        } finally {
          addBtn.disabled = false;
        }
      });
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

      const escapeXml = (value) =>
        String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&apos;");

      const getLabel = (displayName, fallback) =>
        displayName?.UserLocalizedLabel?.Label ||
        displayName?.LocalizedLabels?.[0]?.Label ||
        fallback ||
        "";

      const fetchAll = async (url) => {
        const all = [];

        while (url) {
          const res = await fetch(url, {
            headers: {
              Accept: "application/json",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0"
            }
          });

          const data = await res.json();

          if (!res.ok) {
            throw new Error(data?.error?.message || "Fetch failed");
          }

          all.push(...(data.value || []));
          url = data["@odata.nextLink"] || null;
        }

        return all;
      };

      const state = {
        entities: [],
        fields: [],
        entity: null,
        filterType: "and",
        selectedColumns: [],
        conditions: []
      };

      const overlay = document.createElement("div");
      overlay.id = "__d365_fetch_builder";
      overlay.style.cssText = `
        position: fixed;
        inset: 0;
        z-index: 2147483647;
        background: rgba(0,0,0,.35);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
        direction: ltr;
      `;

      const box = document.createElement("div");
      box.style.cssText = `
        width: min(1200px, 96vw);
        height: min(880px, 94vh);
        background: #fff;
        border-radius: 16px;
        box-shadow: 0 18px 50px rgba(0,0,0,.35);
        display: flex;
        flex-direction: column;
        overflow: hidden;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      `;

      box.innerHTML = `
        <div style="padding:14px 16px;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;align-items:center;">
          <div style="font-weight:900;font-size:16px;">🧱 Fetch Builder</div>
          <button id="fbClose" style="border:1px solid #cbd5e1;background:#fff;border-radius:10px;padding:8px 12px;font-weight:700;cursor:pointer;">Close</button>
        </div>

        <div style="padding:14px 16px;display:grid;grid-template-columns:1fr 1fr;gap:14px;overflow:auto;">
          <div style="display:grid;gap:8px;">
            <label style="font-weight:800;">Entity</label>
            <input id="fbEntitySearch" placeholder="Search entity by display name / logical name"
              style="border:1px solid #cbd5e1;border-radius:10px;padding:10px;" />
            <select id="fbEntitySelect" size="8"
              style="border:1px solid #cbd5e1;border-radius:10px;padding:8px;"></select>
          </div>

          <div style="display:grid;gap:8px;">
            <label style="font-weight:800;">Columns</label>
            <input id="fbFieldSearch" placeholder="Search fields by display name / schema / logical name"
              style="border:1px solid #cbd5e1;border-radius:10px;padding:10px;" />

            <div id="fbColumnsSelect"
              style="border:1px solid #cbd5e1;border-radius:10px;padding:8px;height:180px;overflow:auto;display:grid;gap:6px;">
            </div>

            <div id="fbSelectedColumns" style="font-size:12px;color:#475569;">
              Selected: 0
            </div>
          </div>

          <div style="grid-column:1 / -1;display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
            <label style="font-weight:800;">Filter type:</label>

            <select id="fbFilterType" style="border:1px solid #cbd5e1;border-radius:10px;padding:8px;">
              <option value="and">AND</option>
              <option value="or">OR</option>
            </select>

            <button id="fbAddCondition"
              style="border:none;background:#2563eb;color:#fff;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer;">
              + Add condition
            </button>

            <button id="fbGenerate"
              style="border:none;background:#16a34a;color:#fff;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer;">
              Generate FetchXML
            </button>

            <button id="fbCopy"
              style="border:1px solid #cbd5e1;background:#fff;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer;">
              Copy
            </button>
          </div>

          <div style="grid-column:1 / -1;">
            <div style="font-weight:800;margin-bottom:8px;">Conditions</div>
            <div id="fbConditions" style="display:grid;gap:10px;"></div>
          </div>

          <div style="grid-column:1 / -1;">
            <label style="font-weight:800;">FetchXML</label>
            <textarea id="fbOutput" readonly
              style="margin-top:8px;width:100%;height:260px;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:10px;padding:10px;font-family:Consolas,Monaco,'Courier New',monospace;font-size:12px;direction:ltr;text-align:left;"></textarea>
          </div>
        </div>
      `;

      overlay.appendChild(box);
      document.body.appendChild(overlay);

      const $ = (selector) => box.querySelector(selector);

      const entitySearch = $("#fbEntitySearch");
      const entitySelect = $("#fbEntitySelect");
      const fieldSearch = $("#fbFieldSearch");
      const columnsSelect = $("#fbColumnsSelect");
      const selectedColumnsLabel = $("#fbSelectedColumns");
      const conditionsWrap = $("#fbConditions");
      const output = $("#fbOutput");

      $("#fbClose").onclick = () => overlay.remove();

      const getOperatorsByType = (type) => {
        const t = String(type || "").toLowerCase();

        if (t.includes("string") || t.includes("memo")) {
          return [
            ["eq", "Equals"],
            ["ne", "Not Equals"],
            ["like", "Contains"],
            ["not-like", "Does Not Contain"],
            ["null", "Is Empty"],
            ["not-null", "Is Not Empty"]
          ];
        }

        if (t.includes("datetime")) {
          return [
            ["on", "On"],
            ["on-or-after", "On Or After"],
            ["on-or-before", "On Or Before"],
            ["last-x-days", "Last X Days"],
            ["next-x-days", "Next X Days"],
            ["null", "Is Empty"],
            ["not-null", "Is Not Empty"]
          ];
        }

        if (
          t.includes("picklist") ||
          t.includes("state") ||
          t.includes("status") ||
          t.includes("boolean")
        ) {
          return [
            ["eq", "Equals"],
            ["ne", "Not Equals"],
            ["null", "Is Empty"],
            ["not-null", "Is Not Empty"]
          ];
        }

        if (
          t.includes("integer") ||
          t.includes("decimal") ||
          t.includes("double") ||
          t.includes("money")
        ) {
          return [
            ["eq", "Equals"],
            ["ne", "Not Equals"],
            ["gt", "Greater Than"],
            ["ge", "Greater Or Equal"],
            ["lt", "Less Than"],
            ["le", "Less Or Equal"],
            ["null", "Is Empty"],
            ["not-null", "Is Not Empty"]
          ];
        }

        if (t.includes("lookup") || t.includes("customer") || t.includes("owner")) {
          return [
            ["eq", "Equals GUID"],
            ["ne", "Not Equals GUID"],
            ["eq-userid", "Current User"],
            ["null", "Is Empty"],
            ["not-null", "Is Not Empty"]
          ];
        }

        return [
          ["eq", "Equals"],
          ["ne", "Not Equals"],
          ["null", "Is Empty"],
          ["not-null", "Is Not Empty"]
        ];
      };

      const renderEntities = () => {
        const q = entitySearch.value.trim().toLowerCase();
        entitySelect.innerHTML = "";

        state.entities
          .filter((e) => {
            const text = `${e.label} ${e.logicalName} ${e.schemaName}`.toLowerCase();
            return !q || text.includes(q);
          })
          .sort((a, b) =>
            (a.label || a.logicalName).localeCompare(b.label || b.logicalName)
          )
          .forEach((e) => {
            const opt = document.createElement("option");
            opt.value = e.logicalName;
            opt.textContent = `${e.label || e.logicalName} (${e.logicalName})`;
            entitySelect.appendChild(opt);
          });
      };

      const renderFields = () => {
        const q = fieldSearch.value.trim().toLowerCase();
        columnsSelect.innerHTML = "";

        const filtered = state.fields
          .filter((f) => {
            const text = `${f.label} ${f.logicalName} ${f.schemaName}`.toLowerCase();
            return !q || text.includes(q);
          })
          .sort((a, b) =>
            (a.label || a.logicalName).localeCompare(b.label || b.logicalName)
          );

        filtered.forEach((f) => {
          const row = document.createElement("label");
          row.style.cssText = `
            display:flex;
            align-items:center;
            gap:8px;
            cursor:pointer;
            padding:6px;
            border-radius:8px;
            font-size:13px;
          `;

          row.onmouseenter = () => (row.style.background = "#f8fafc");
          row.onmouseleave = () => (row.style.background = "");

          const cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = state.selectedColumns.includes(f.logicalName);

          cb.onchange = () => {
            if (cb.checked) {
              if (!state.selectedColumns.includes(f.logicalName)) {
                state.selectedColumns.push(f.logicalName);
              }
            } else {
              state.selectedColumns = state.selectedColumns.filter(
                (x) => x !== f.logicalName
              );
            }

            selectedColumnsLabel.textContent = `Selected: ${state.selectedColumns.length}`;
          };

          const text = document.createElement("span");
          text.textContent = `${f.label || f.logicalName} (${f.logicalName})`;

          row.appendChild(cb);
          row.appendChild(text);
          columnsSelect.appendChild(row);
        });

        selectedColumnsLabel.textContent = `Selected: ${state.selectedColumns.length}`;
      };

      const getField = (logicalName) =>
        state.fields.find((f) => f.logicalName === logicalName);

      const renderConditions = () => {
        conditionsWrap.innerHTML = "";

        state.conditions.forEach((condition, index) => {
          const row = document.createElement("div");
          row.style.cssText = `
            display:grid;
            grid-template-columns: 1.3fr .8fr 1.2fr auto;
            gap:8px;
            align-items:center;
          `;

          const fieldSelect = document.createElement("select");
          fieldSelect.style.cssText =
            "border:1px solid #cbd5e1;border-radius:10px;padding:8px;";

          state.fields
            .sort((a, b) =>
              (a.label || a.logicalName).localeCompare(b.label || b.logicalName)
            )
            .forEach((f) => {
              const opt = document.createElement("option");
              opt.value = f.logicalName;
              opt.textContent = `${f.label || f.logicalName} (${f.logicalName})`;
              if (condition.field === f.logicalName) opt.selected = true;
              fieldSelect.appendChild(opt);
            });

          const operatorSelect = document.createElement("select");
          operatorSelect.style.cssText =
            "border:1px solid #cbd5e1;border-radius:10px;padding:8px;";

          const selectedField = getField(condition.field);
          const operators = getOperatorsByType(selectedField?.type);

          operators.forEach(([value, label]) => {
            const opt = document.createElement("option");
            opt.value = value;
            opt.textContent = label;
            if (condition.operator === value) opt.selected = true;
            operatorSelect.appendChild(opt);
          });

          const valueWrap = document.createElement("div");

          const renderValueInput = () => {
            valueWrap.innerHTML = "";

            const field = getField(condition.field);
            const operator = condition.operator;

            if (operator === "null" || operator === "not-null" || operator === "eq-userid") {
              const span = document.createElement("span");
              span.textContent = "No value needed";
              span.style.cssText = "color:#64748b;font-size:13px;";
              valueWrap.appendChild(span);
              return;
            }

            if (field?.options?.length) {
              const select = document.createElement("select");
              select.style.cssText =
                "width:100%;border:1px solid #cbd5e1;border-radius:10px;padding:8px;";

              field.options.forEach((o) => {
                const opt = document.createElement("option");
                opt.value = o.value;
                opt.textContent = `${o.label} (${o.value})`;
                if (String(condition.value) === String(o.value)) opt.selected = true;
                select.appendChild(opt);
              });

              select.onchange = () => {
                condition.value = select.value;
              };

              valueWrap.appendChild(select);

              if (!condition.value && field.options[0]) {
                condition.value = field.options[0].value;
              }

              return;
            }

            const input = document.createElement("input");
            input.value = condition.value || "";
            input.placeholder = "Value";
            input.style.cssText =
              "width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:10px;padding:8px;";

            if (String(field?.type || "").toLowerCase().includes("datetime")) {
              input.type = "date";
            }

            input.oninput = () => {
              condition.value = input.value;
            };

            valueWrap.appendChild(input);
          };

          const removeBtn = document.createElement("button");
          removeBtn.textContent = "Remove";
          removeBtn.style.cssText = `
            border:1px solid #fecaca;
            background:#fff;
            color:#b91c1c;
            border-radius:10px;
            padding:8px 10px;
            font-weight:700;
            cursor:pointer;
          `;

          fieldSelect.onchange = () => {
            condition.field = fieldSelect.value;
            condition.operator = "eq";
            condition.value = "";
            renderConditions();
          };

          operatorSelect.onchange = () => {
            condition.operator = operatorSelect.value;
            condition.value = "";
            renderConditions();
          };

          removeBtn.onclick = () => {
            state.conditions.splice(index, 1);
            renderConditions();
          };

          row.appendChild(fieldSelect);
          row.appendChild(operatorSelect);
          row.appendChild(valueWrap);
          row.appendChild(removeBtn);

          conditionsWrap.appendChild(row);
          renderValueInput();
        });
      };

      const loadEntities = async () => {
        entitySelect.innerHTML = "<option>Loading entities...</option>";

        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        const rows = await fetchAll(
          `${clientUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,SchemaName,DisplayName,ObjectTypeCode&$filter=IsPrivate eq false`
        );

        state.entities = rows.map((e) => ({
          logicalName: e.LogicalName,
          schemaName: e.SchemaName,
          label: getLabel(e.DisplayName, e.LogicalName)
        }));

        renderEntities();
      };

      const loadOptionsForField = async (entityLogicalName, field) => {
        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        const type = String(field.type || "").toLowerCase();

        let cast = null;

        if (type.includes("picklist")) {
          cast = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata";
        } else if (type.includes("status")) {
          cast = "Microsoft.Dynamics.CRM.StatusAttributeMetadata";
        } else if (type.includes("state")) {
          cast = "Microsoft.Dynamics.CRM.StateAttributeMetadata";
        } else if (type.includes("boolean")) {
          cast = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata";
        }

        if (!cast) return;

        try {
          const url =
            `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')` +
            `/Attributes(LogicalName='${field.logicalName}')/${cast}?$expand=OptionSet`;

          const res = await fetch(url, {
            headers: {
              Accept: "application/json",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0"
            }
          });

          const data = await res.json();

          if (!res.ok) return;

          if (type.includes("boolean")) {
            const trueOption = data.OptionSet?.TrueOption;
            const falseOption = data.OptionSet?.FalseOption;

            field.options = [
              {
                value: falseOption?.Value ?? 0,
                label: getLabel(falseOption?.Label, "No")
              },
              {
                value: trueOption?.Value ?? 1,
                label: getLabel(trueOption?.Label, "Yes")
              }
            ];
          } else {
            field.options = (data.OptionSet?.Options || []).map((o) => ({
              value: o.Value,
              label: getLabel(o.Label, String(o.Value))
            }));
          }
        } catch {
          field.options = [];
        }
      };

      const loadFields = async (entityLogicalName) => {
        columnsSelect.innerHTML = "Loading fields...";
        conditionsWrap.innerHTML = "";
        output.value = "";

        state.entity = entityLogicalName;
        state.fields = [];
        state.selectedColumns = [];
        state.conditions = [];

        selectedColumnsLabel.textContent = "Selected: 0";

        const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        const url =
          `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
          `?$select=LogicalName,SchemaName,DisplayName,AttributeType`;

        const attrs = await fetchAll(url);

        state.fields = attrs
          .filter(
            (a) =>
              a.LogicalName &&
              !a.LogicalName.startsWith("_") &&
              !a.LogicalName.includes("@")
          )
          .map((a) => ({
            logicalName: a.LogicalName,
            schemaName: a.SchemaName,
            label: getLabel(a.DisplayName, a.LogicalName),
            type: a.AttributeType,
            options: []
          }));

        const optionFields = state.fields.filter((f) => {
          const t = String(f.type || "").toLowerCase();
          return (
            t.includes("picklist") ||
            t.includes("status") ||
            t.includes("state") ||
            t.includes("boolean")
          );
        });

        await Promise.all(
          optionFields.map((field) => loadOptionsForField(entityLogicalName, field))
        );

        renderFields();
      };

      const buildFetchXml = () => {
        if (!state.entity) {
          output.value = "Please select entity.";
          return;
        }

        const columns =
          state.selectedColumns.length > 0
            ? state.selectedColumns
            : [state.fields[0]?.logicalName].filter(Boolean);

        const attrsXml = columns
          .map((name) => `    <attribute name="${escapeXml(name)}" />`)
          .join("\n");

        const conditionsXml = state.conditions
          .filter((c) => c.field && c.operator)
          .map((c) => {
            if (["null", "not-null", "eq-userid"].includes(c.operator)) {
              return `      <condition attribute="${escapeXml(c.field)}" operator="${escapeXml(c.operator)}" />`;
            }

            const value =
              c.operator === "like" || c.operator === "not-like"
                ? `%${c.value || ""}%`
                : c.value || "";

            return `      <condition attribute="${escapeXml(c.field)}" operator="${escapeXml(c.operator)}" value="${escapeXml(value)}" />`;
          })
          .join("\n");

        const filterXml = conditionsXml
          ? `    <filter type="${escapeXml(state.filterType)}">
${conditionsXml}
    </filter>`
          : "";

        output.value = `<fetch top="50">
  <entity name="${escapeXml(state.entity)}">
${attrsXml}
${filterXml}
  </entity>
</fetch>`;
      };

      entitySearch.oninput = renderEntities;
      fieldSearch.oninput = renderFields;

      entitySelect.onchange = async () => {
        try {
          await loadFields(entitySelect.value);
        } catch (e) {
          output.value = `Failed loading fields:\n${e.message || e}`;
        }
      };

      $("#fbFilterType").onchange = (e) => {
        state.filterType = e.target.value;
      };

      $("#fbAddCondition").onclick = () => {
        if (!state.entity || !state.fields.length) {
          alert("Select entity first.");
          return;
        }

        state.conditions.push({
          field: state.fields[0].logicalName,
          operator: "eq",
          value: ""
        });

        renderConditions();
      };

      $("#fbGenerate").onclick = buildFetchXml;

      $("#fbCopy").onclick = async () => {
        await navigator.clipboard.writeText(output.value);
      };

      try {
        await loadEntities();
      } catch (e) {
        output.value = `Failed loading metadata:\n${e.message || e}`;
      }
    }
  });
});