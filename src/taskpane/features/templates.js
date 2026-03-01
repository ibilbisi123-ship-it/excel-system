/* global document, fetch, Excel */

export function setupTemplateDropdown(updateStatus, addMessage) {
  const templateDropdownBtn = document.getElementById("templateDropdownBtn");
  const templateMenu = document.getElementById("templateMenu");
  if (!templateDropdownBtn || !templateMenu) return;

  const toggleMenu = (open) => {
    const isOpen = open !== undefined ? !!open : templateMenu.classList.contains("hidden");
    templateMenu.classList.toggle("hidden", !isOpen);
    templateDropdownBtn.setAttribute("aria-expanded", String(isOpen));
  };

  templateDropdownBtn.addEventListener("click", async () => {
    toggleMenu();
    if (templateMenu.dataset.loaded !== "1") {
      await loadTemplateList(updateStatus, addMessage);
      templateMenu.dataset.loaded = "1";
    }
  });

  document.addEventListener("click", (e) => {
    if (!templateMenu.contains(e.target) && e.target !== templateDropdownBtn) {
      templateMenu.classList.add("hidden");
      templateDropdownBtn.setAttribute("aria-expanded", "false");
    }
  });
}

async function loadTemplateList(updateStatus, addMessage) {
  try {
    const container = document.getElementById("templateMenuItems");
    const empty = document.getElementById("templateMenuEmpty");
    if (!container) return;
    container.innerHTML = "";
    const resp = await fetch("assets/templates/templates.json", { cache: "no-store" });
    if (!resp.ok) throw new Error("Missing templates.json");
    const list = await resp.json();
    const items = Array.isArray(list) ? list : list?.templates || [];
    if (!items.length) {
      if (empty) empty.classList.remove("hidden");
      return;
    }
    if (empty) empty.classList.add("hidden");
    for (const t of items) {
      const name = t.name || t.title || t.file || "Template";
      const file = t.file || t.path;
      if (!file) continue;
      const btn = document.createElement("button");
      btn.className = "btn btn--outline";
      btn.style.display = "block";
      btn.style.width = "100%";
      btn.style.textAlign = "left";
      btn.style.margin = "4px 0";
      btn.textContent = name;
      btn.addEventListener("click", async () => {
        const menu = document.getElementById("templateMenu");
        if (menu) menu.classList.add("hidden");
        await insertWorkbookTemplateFromUrl(`assets/templates/${file}`, updateStatus, addMessage);
      });
      container.appendChild(btn);
    }
  } catch (e) {
    console.error("Failed to load templates list", e);
    if (updateStatus) updateStatus("Failed to load templates", "error");
  }
}

async function insertWorkbookTemplateFromUrl(relativeUrl, updateStatus, addMessage) {
  try {
    if (typeof Excel === "undefined") {
      updateStatus && updateStatus("Excel APIs not available", "error");
      return;
    }
    updateStatus && updateStatus("Downloading template...", "processing");
    const resp = await fetch(relativeUrl);
    if (!resp.ok) throw new Error(`Failed to fetch ${relativeUrl}`);
    const blob = await resp.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const base64 = await arrayBufferToBase64(arrayBuffer);

    updateStatus && updateStatus("Inserting template...", "processing");
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      if (typeof wb.insertWorksheetsFromBase64 === "function") {
        wb.insertWorksheetsFromBase64(base64, null, "After", active);
        await context.sync();
        return;
      }
      if (typeof wb.addFromBase64 === "function") {
        wb.addFromBase64(base64);
        await context.sync();
        return;
      }
      throw new Error("Excel version lacks base64 import APIs; update Office.");
    });

    updateStatus && updateStatus("Template inserted", "success");
    addMessage &&
      addMessage("Template workbook has been inserted into the current workbook.", "ai");
  } catch (e) {
    console.error("Insert template from URL failed", e);
    updateStatus && updateStatus("Failed to insert template", "error");
  }
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return Promise.resolve(btoa(binary));
}
