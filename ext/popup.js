// popup.js

document.getElementById("ribbondebug").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !tab?.url) return;

  if (tab.url.includes("ribbondebug=true")) return;

  const joiner = tab.url.includes("?") ? "&" : "?";
  chrome.tabs.update(tab.id, { url: tab.url + joiner + "ribbondebug=true" });
});

document.getElementById("tabsname").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
      overlay.addEventListener("click", (e) => {
        if (e.target === overlay) close();
      });

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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
document.getElementById("openDefaultView").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !tab.url) return;

  // ask entity
  const entity = prompt("Enter entity logical name (example: contact, incident, account):");
  if (!entity) return;

  const cleanEntity = entity.trim().toLowerCase();
  if (!cleanEntity) return;

  // get base org url + current appid from the current tab
  const u = new URL(tab.url);
  const orgUrl = `${u.protocol}//${u.host}`;
  const appid = u.searchParams.get("appid"); // can be null

  // build default view url
  // Note: # is optional; without it D365 still opens the default view for that entity
  const url =
    `${orgUrl}/main.aspx?` +
    (appid ? `appid=${encodeURIComponent(appid)}&` : "") +
    `pagetype=entitylist&etn=${encodeURIComponent(cleanEntity)}`;

  // open in new tab
  chrome.tabs.create({ url });
});

document.getElementById("getRolePermissions").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
      header.textContent = "Role Permissions By Entity";

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

      const roleInput = document.createElement("input");
      roleInput.placeholder = "Role name (e.g. מנהל מערכת)";
      roleInput.style.cssText = inputStyle;

      const entityInput = document.createElement("input");
      entityInput.placeholder = "Entity logical name (e.g. contact)";
      entityInput.style.cssText = inputStyle;

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

      body.appendChild(mkRow("Role Name", roleInput));
      body.appendChild(mkRow("Entity Logical Name", entityInput));
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
        const cols = ["roleName", "entityName", "permission", "depth", "privilegeName"];
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
          { key: "roleName", label: "Role" },
          { key: "entityName", label: "Entity" },
          { key: "permission", label: "Permission" },
          { key: "depth", label: "Depth" },
          { key: "privilegeName", label: "Privilege" }
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
          td.textContent = "No permissions found.";
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
              max-width: 380px;
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

        const roleName = (roleInput.value || "").trim();
        const entityLogicalName = (entityInput.value || "").trim().toLowerCase();

        if (!roleName || !entityLogicalName) {
          status.textContent = "❌ Role name and entity logical name are required.";
          summary.textContent = "Missing input.";
          return;
        }

        const Xrm = window.Xrm;
        const webApi = Xrm?.WebApi || Xrm?.WebApi?.online;

        if (!Xrm || !webApi?.retrieveMultipleRecords) {
          status.textContent = "❌ Xrm.WebApi.retrieveMultipleRecords not available. Open a D365 page.";
          summary.textContent = "D365 context not found.";
          return;
        }

        status.textContent = "⏳ Loading…";

        try {
          const depthMap = {
            0: "None",
            1: "User",
            2: "Business Unit",
            4: "Parent:Child Business Unit",
            8: "Organization"
          };

          const knownRights = [
            "Create",
            "Read",
            "Write",
            "Delete",
            "Append",
            "AppendTo",
            "Assign",
            "Share"
          ];

          const clean = (v) => (v || "").replace(/[{}]/g, "").toLowerCase();

          const roleResult = await webApi.retrieveMultipleRecords(
            "role",
            `?$select=roleid,name&$filter=name eq '${roleName.replace(/'/g, "''")}'`
          );

          if (!roleResult.entities.length) {
            status.textContent = "⚠️ Role not found.";
            summary.textContent = `Role not found: ${roleName}`;
            rawTa.value = JSON.stringify([], null, 2);
            renderTable([]);
            return;
          }

          const roleRows = [];

          for (const role of roleResult.entities) {
            const roleId = clean(role.roleid);

            const fetchXml = `
<fetch distinct="true">
  <entity name="role">
    <attribute name="roleid" />
    <attribute name="name" />
    <filter>
      <condition attribute="roleid" operator="eq" value="${roleId}" />
    </filter>
    <link-entity name="roleprivileges" from="roleid" to="roleid" alias="rp" intersect="true">
      <attribute name="privilegedepthmask" />
      <link-entity name="privilege" from="privilegeid" to="privilegeid" alias="p">
        <attribute name="name" />
      </link-entity>
    </link-entity>
  </entity>
</fetch>`;

            const privilegeResult = await webApi.retrieveMultipleRecords(
              "role",
              "?fetchXml=" + encodeURIComponent(fetchXml)
            );

            const matchedPermissions = [];

            for (const row of privilegeResult.entities) {
              const privilegeName = row["p.name"];
              const depthValue = row["rp.privilegedepthmask"];

              if (!privilegeName || depthValue == null) continue;

              const matchedRight = knownRights.find(
                right => privilegeName.toLowerCase() === `prv${right}${entityLogicalName}`.toLowerCase()
              );

              if (matchedRight) {
                matchedPermissions.push({
                  roleName: role.name,
                  entityName: entityLogicalName,
                  permission: matchedRight,
                  privilegeName,
                  depthValue,
                  depthLabel: depthMap[depthValue] || `Unknown (${depthValue})`
                });
              }
            }

            roleRows.push({
              roleId,
              roleName: role.name,
              entityName: entityLogicalName,
              permissions: matchedPermissions.sort((a, b) => a.permission.localeCompare(b.permission))
            });
          }

          const flatRows = roleRows.flatMap(r =>
            r.permissions.map(p => ({
              roleName: p.roleName,
              entityName: p.entityName,
              permission: p.permission,
              depth: p.depthLabel,
              privilegeName: p.privilegeName
            }))
          );

          lastRows = flatRows;
          rawTa.value = JSON.stringify(flatRows, null, 2);
          renderTable(flatRows);

          summary.innerHTML = `
            <b>Role:</b> ${escapeHtml(roleName)}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Entity:</b> ${escapeHtml(entityLogicalName)}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Permissions:</b> ${flatRows.length}
          `;

          status.textContent = flatRows.length
            ? `✅ Done (${flatRows.length} permissions).`
            : "⚠️ No permissions found.";

          console.log(`Role: ${roleName}`);
          console.log(`Entity: ${entityLogicalName}`);
          console.table(flatRows);
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

      roleInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") runGet();
      });

      entityInput.addEventListener("keydown", (e) => {
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

      roleInput.focus();
    }
  });
});
document.getElementById("getSystemRolesByPrivilege").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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

          const roleUrl = `${clientUrl}/api/data/v9.2/roles?$select=name,roleid`;
          const roleRes = await fetch(roleUrl, {
            method: "GET",
            headers: {
              "Accept": "application/json",
              "OData-Version": "4.0",
              "OData-MaxVersion": "4.0"
            },
            credentials: "same-origin"
          });

          if (!roleRes.ok) {
            throw new Error(`Role lookup failed (${roleRes.status})`);
          }

          const roleData = await roleRes.json();
          const allRoles = roleData.value || [];

          const checkRole = async (role) => {
            try {
              const url = `${clientUrl}/api/data/v9.2/RetrieveRolePrivilegesRole(RoleId=${role.roleid})`;
              const res = await fetch(url, {
                method: "GET",
                headers: {
                  "Accept": "application/json",
                  "OData-Version": "4.0",
                  "OData-MaxVersion": "4.0"
                },
                credentials: "same-origin"
              });

              if (!res.ok) return null;

              const data = await res.json();
              const privs = data.RolePrivileges || [];

              const hit = privs.find(p => normalizeGuid(p.PrivilegeId) === targetPrivId);
              if (hit && getDepthValue(hit.Depth) === wantedDepthValue) {
                return role.name;
              }
            } catch (e) {
              return null;
            }
            return null;
          };

          const rawResults = await Promise.all(allRoles.map(r => checkRole(r)));
          const uniqueRoleNames = [...new Set(rawResults.filter(name => name !== null))].sort();

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
document.getElementById("assignSecurityRoleUi").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
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
        width: min(980px, 96vw);
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
          <h2 style="margin:0;font-size:22px;">הקצאת תפקיד אבטחה למשתמש</h2>
          <button id="__assignRoleClose" style="
            border:none;
            background:#f3f3f3;
            border-radius:10px;
            padding:8px 12px;
            cursor:pointer;
            font-size:16px;
          ">✖</button>
        </div>
        <div style="margin-bottom:14px;padding:10px 12px;background:#f7f7f7;border-radius:10px;">
          <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:600;">
            <input id="__assignRoleUseCurrentUser" type="checkbox" />
            בצע עליי (המשתמש המחובר)
          </label>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
          <div>
            <label style="display:block;font-weight:600;margin-bottom:8px;">חיפוש משתמש</label>
            <input id="__assignRoleUserSearch" type="text" placeholder="חפש לפי שם / יוזר / מייל"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:10px;box-sizing:border-box;" />
            <select id="__assignRoleUserSelect" size="16"
              style="width:100%;padding:8px;border:1px solid #ccc;border-radius:10px;box-sizing:border-box;"></select>
          </div>

          <div>
            <label style="display:block;font-weight:600;margin-bottom:8px;">חיפוש תפקיד אבטחה</label>
            <input id="__assignRoleRoleSearch" type="text" placeholder="חפש תפקיד"
              style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:10px;margin-bottom:10px;box-sizing:border-box;" />
            <select id="__assignRoleRoleSelect" size="16"
              style="width:100%;padding:8px;border:1px solid #ccc;border-radius:10px;box-sizing:border-box;"></select>
          </div>
        </div>

        <div style="display:flex;gap:10px;justify-content:flex-start;margin-top:18px;">
          <button id="__assignRoleSubmit" style="
            border:none;
            background:#0f6cbd;
            color:white;
            border-radius:10px;
            padding:10px 18px;
            cursor:pointer;
            font-size:15px;
          ">הקצה תפקיד</button>

          <button id="__assignRoleCancel" style="
            border:none;
            background:#eaeaea;
            color:#222;
            border-radius:10px;
            padding:10px 18px;
            cursor:pointer;
            font-size:15px;
          ">סגור</button>
        </div>

        <div id="__assignRoleStatus" style="
          margin-top:14px;
          padding:10px 12px;
          background:#f7f7f7;
          border-radius:10px;
          min-height:22px;
          white-space:pre-wrap;
          font-size:14px;
        "></div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      const closeModal = () => overlay.remove();
      document.getElementById("__assignRoleClose").onclick = closeModal;
      document.getElementById("__assignRoleCancel").onclick = closeModal;
      overlay.addEventListener("click", (e) => {
        if (e.target === overlay) closeModal();
      });

      const userSearchInput = document.getElementById("__assignRoleUserSearch");
      const userSelect = document.getElementById("__assignRoleUserSelect");
      const roleSearchInput = document.getElementById("__assignRoleRoleSearch");
      const roleSelect = document.getElementById("__assignRoleRoleSelect");
      const submitBtn = document.getElementById("__assignRoleSubmit");
      const statusBox = document.getElementById("__assignRoleStatus");
      const useCurrentUserCheckbox = document.getElementById("__assignRoleUseCurrentUser");
      const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
      const BASE_URL = `${clientUrl}/api/data/v9.2`;

      let allUsers = [];
      let allRoles = [];
  async function assignRoleToCurrentUser(roleId) {
        const currentUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, "");

        const user = await fetchJSON(
          `${BASE_URL}/systemusers(${currentUserId})?$select=systemuserid,fullname,domainname,isdisabled`
        );

        const roleData = await fetchJSON(
          `${BASE_URL}/roles(${roleId})?$select=roleid,name`
        );
      if (!roleData?.roleid) throw new Error(`תפקיד לא נמצא: ${roleId}`);

        const res = await fetch(
          `${BASE_URL}/systemusers(${user.systemuserid})/systemuserroles_association/$ref`,
          {
            method: "POST",
            headers: {
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
              "Accept": "application/json",
              "Content-Type": "application/json"
            },
            credentials: "same-origin",
            body: JSON.stringify({
              "@odata.id": `${BASE_URL}/roles(${roleData.roleid})`
                })
              }
            );

        if (!res.ok) {
          let message = `${res.status}`;
          try {
            const err = await res.json();
            message = err?.error?.message || message;
          } catch (_) {}
          throw new Error(message);
        }

        return { user, role: roleData };
      }
      function escapeODataString(str) {
        return String(str ?? "").replace(/'/g, "''");
      }
      function toggleUserSelectionState() {
        const disabled = useCurrentUserCheckbox.checked;

        userSearchInput.disabled = disabled;
        userSelect.disabled = disabled;

        userSearchInput.style.opacity = disabled ? "0.6" : "1";
        userSelect.style.opacity = disabled ? "0.6" : "1";
      }
      useCurrentUserCheckbox.addEventListener("change", toggleUserSelectionState);
      toggleUserSelectionState();
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

        return await res.json();
      }

      async function fetchAllPages(url) {
        let all = [];
        let nextUrl = url;

        while (nextUrl) {
          const data = await fetchJSON(nextUrl);
          all = all.concat(data.value || []);
          nextUrl = data["@odata.nextLink"] || null;
        }

        return all;
      }

      async function assignRoleByName_ByUsername(username, roleName) {
        username = username + "@mac.org.il";

        const users = await fetchJSON(
          `${BASE_URL}/systemusers?$filter=domainname eq '${escapeODataString(username)}'&$select=systemuserid,fullname,domainname,isdisabled`
        );
        if (!users.value.length) throw new Error(`משתמש לא נמצא: ${username}`);
        const user = users.value[0];

        const roles = await fetchJSON(
          `${BASE_URL}/roles?$filter=name eq '${escapeODataString(roleName)}'&$select=roleid,name`
        );
        if (!roles.value.length) throw new Error(`תפקיד לא נמצא: ${roleName}`);
        const role = roles.value[0];

        const res = await fetch(
          `${BASE_URL}/systemusers(${user.systemuserid})/systemuserroles_association/$ref`,
          {
            method: "POST",
            headers: {
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
              "Accept": "application/json",
              "Content-Type": "application/json"
            },
            credentials: "same-origin",
            body: JSON.stringify({
              "@odata.id": `${BASE_URL}/roles(${role.roleid})`
            })
          }
        );

        if (!res.ok) {
          let message = `${res.status}`;
          try {
            const err = await res.json();
            message = err?.error?.message || message;
          } catch (_) {}
          throw new Error(message);
        }

        return { user, role };
      }

      async function loadUsers() {
        const users = await fetchAllPages(
          `${BASE_URL}/systemusers?$select=systemuserid,fullname,domainname,isdisabled,internalemailaddress&$orderby=fullname asc`
        );

        allUsers = users.map(u => {
          const domain = u.domainname || "";
          const email = u.internalemailaddress || "";
          const username = domain ? domain.replace(/@mac\.org\.il$/i, "") : "";

          return {
            id: u.systemuserid,
            fullname: u.fullname || "",
            domainname: domain,
            internalemailaddress: email,
            isdisabled: !!u.isdisabled,
            username: username
          };
        });
      }

      async function loadRoles() {
        const businessUnits = await fetchAllPages(
          `${BASE_URL}/businessunits?$select=businessunitid,name,_parentbusinessunitid_value`
        );

        const rootBusinessUnit = businessUnits.find(bu => !bu._parentbusinessunitid_value);
        if (!rootBusinessUnit) {
          throw new Error("לא נמצאה יחידה עסקית ראשית");
        }

        const rootBusinessUnitId = rootBusinessUnit.businessunitid;

        const roles = await fetchAllPages(
          `${BASE_URL}/roles?$select=roleid,name,_businessunitid_value&$orderby=name asc`
        );

        allRoles = roles
          .filter(r => r.name && r._businessunitid_value === rootBusinessUnitId)
          .sort((a, b) => (a.name || "").localeCompare(b.name || "", "he"))
          .map(r => ({
            id: r.roleid,
            name: r.name || "",
            businessunitid: r._businessunitid_value
          }));
      }
async function assignRoleById_ToUsername(username, roleId) {
  username = username + "@mac.org.il";

  const users = await fetchJSON(
    `${BASE_URL}/systemusers?$filter=domainname eq '${escapeODataString(username)}'&$select=systemuserid,fullname,domainname,isdisabled`
  );
  if (!users.value.length) throw new Error(`משתמש לא נמצא: ${username}`);
  const user = users.value[0];

  const roleData = await fetchJSON(
    `${BASE_URL}/roles(${roleId})?$select=roleid,name`
  );
  if (!roleData?.roleid) throw new Error(`תפקיד לא נמצא: ${roleId}`);

  const res = await fetch(
    `${BASE_URL}/systemusers(${user.systemuserid})/systemuserroles_association/$ref`,
    {
      method: "POST",
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      credentials: "same-origin",
      body: JSON.stringify({
        "@odata.id": `${BASE_URL}/roles(${roleData.roleid})`
      })
    }
  );

  if (!res.ok) {
    let message = `${res.status}`;
    try {
      const err = await res.json();
      message = err?.error?.message || message;
    } catch (_) {}
    throw new Error(message);
  }

  return { user, role: roleData };
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
          option.value = user.username || user.domainname || user.fullname;

          const statusText = user.isdisabled ? "לא פעיל" : "פעיל";
          const userLabel = user.username || user.domainname || "ללא domainname";

          option.textContent = `${user.fullname || "(ללא שם)"} | ${userLabel} | ${statusText}`;
          option.dataset.domainname = user.domainname || "";
          option.dataset.userid = user.id;
          option.dataset.isdisabled = String(user.isdisabled);
          option.dataset.email = user.internalemailaddress || "";
          userSelect.appendChild(option);
        }

        if (filtered.length > 0) userSelect.selectedIndex = 0;
      }

      function renderRoles(searchText = "") {
        const q = searchText.trim().toLowerCase();
        roleSelect.innerHTML = "";

        const filtered = allRoles.filter(r => {
          if (!q) return true;
          return (r.name || "").toLowerCase().includes(q);
        });

        for (const role of filtered) {
          const option = document.createElement("option");
          option.value = role.name;
          option.textContent = role.name;
          option.dataset.roleid = role.id;
          roleSelect.appendChild(option);
        }

        if (filtered.length > 0) roleSelect.selectedIndex = 0;
      }

      try {
        statusBox.textContent = "טוען את כל המשתמשים (פעילים ולא פעילים) ואת כל התפקידים...";
        await Promise.all([loadUsers(), loadRoles()]);
        renderUsers();
        renderRoles();

        const activeCount = allUsers.filter(u => !u.isdisabled).length;
        const disabledCount = allUsers.filter(u => u.isdisabled).length;

        statusBox.textContent =
          `✅ נטענו ${allUsers.length} משתמשים (${activeCount} פעילים, ${disabledCount} לא פעילים)\n` +
          `✅ נטענו ${allRoles.length} תפקידים`;
      } catch (err) {
        statusBox.textContent = `❌ שגיאה בטעינה: ${err.message}`;
        return;
      }

      userSearchInput.addEventListener("input", () => renderUsers(userSearchInput.value));
      roleSearchInput.addEventListener("input", () => renderRoles(roleSearchInput.value));

      userSearchInput.addEventListener("keydown", (e) => {
        if (e.key === "ArrowDown") {
          userSelect.focus();
          if (userSelect.options.length && userSelect.selectedIndex < 0) {
            userSelect.selectedIndex = 0;
          }
        }
      });

      roleSearchInput.addEventListener("keydown", (e) => {
        if (e.key === "ArrowDown") {
          roleSelect.focus();
          if (roleSelect.options.length && roleSelect.selectedIndex < 0) {
            roleSelect.selectedIndex = 0;
          }
        }
      });

      submitBtn.addEventListener("click", async () => {
  const selectedRole = roleSelect.options[roleSelect.selectedIndex];

  if (!selectedRole) {
    statusBox.textContent = "צריך לבחור תפקיד מהרשימה.";
    return;
  }

  const useCurrentUser = useCurrentUserCheckbox.checked;
  const roleId = selectedRole.dataset.roleid;
  const roleName = selectedRole.value;

  let selectedUser = null;
  let username = null;

  if (!useCurrentUser) {
    selectedUser = userSelect.options[userSelect.selectedIndex];

    if (!selectedUser) {
      statusBox.textContent = "צריך לבחור משתמש מהרשימה או לסמן 'בצע עליי'.";
      return;
    }

    username = selectedUser.value;

    if (!username) {
      statusBox.textContent = "למשתמש שנבחר אין domainname / username תקין.";
      return;
    }
  }

  submitBtn.disabled = true;

  if (useCurrentUser) {
    statusBox.textContent = `מקצה את התפקיד "${roleName}" למשתמש המחובר...`;
  } else {
    statusBox.textContent = `מקצה את התפקיד "${roleName}" למשתמש "${username}"...`;
  }

  try {
    let result;

    if (useCurrentUser) {
      result = await assignRoleToCurrentUser(roleId);
    } else {
      result = await assignRoleById_ToUsername(username, roleId);
    }

    statusBox.textContent =
      `✅ התפקיד הוקצה בהצלחה\n` +
      `משתמש: ${result.user.fullname} (${result.user.domainname || "ללא domainname"})\n` +
      `סטטוס: ${result.user.isdisabled ? "לא פעיל" : "פעיל"}\n` +
      `תפקיד: ${result.role.name}`;
  } catch (err) {
    statusBox.textContent = `❌ שגיאה: ${err.message}`;
  } finally {
    submitBtn.disabled = false;
  }
});
    }
  });
});