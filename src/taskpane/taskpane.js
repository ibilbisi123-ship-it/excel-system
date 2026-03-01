/* global document, window, fetch, Office, Excel, console, localStorage, setTimeout */

import { checkLicense, activateLicense, clearLicense } from "./features/licenseGate.js";

// Version injected by webpack DefinePlugin from version.json
const APP_VERSION = typeof process !== "undefined" && process.env && process.env.APP_VERSION
  ? process.env.APP_VERSION
  : "1.0.0";

// Helper functions for geometry (moved out for safety)
// Helper functions for geometry (moved out for safety)
function calculateLengthHelper(points, width, height) {
  if (!points || !Array.isArray(points) || points.length < 2) return 0;
  let total = 0;
  for (let i = 1; i < points.length; i++) {
    const p1 = points[i - 1];
    const p2 = points[i];

    // Loosely check for valid point-like objects
    if (!p1 || !p2 || p1.x == null || p2.x == null) continue;

    // Coerce to number
    const x1 = Number(p1.x);
    const y1 = Number(p1.y);
    const x2 = Number(p2.x);
    const y2 = Number(p2.y);

    if (isNaN(x1) || isNaN(y1) || isNaN(x2) || isNaN(y2)) continue;

    const dx = (x2 - x1) * width;
    const dy = (y2 - y1) * height;
    total += Math.sqrt(dx * dx + dy * dy);
  }
  return total;
}

function calculateAreaHelper(points, width, height) {
  if (!points || !Array.isArray(points) || points.length < 3) return 0;
  let area = 0;
  for (let i = 0; i < points.length; i++) {
    const p1 = points[i];
    const p2 = points[(i + 1) % points.length];

    if (!p1 || !p2 || p1.x == null || p2.x == null) continue;

    const x1 = Number(p1.x);
    const y1 = Number(p1.y);
    const x2 = Number(p2.x);
    const y2 = Number(p2.y);

    if (isNaN(x1) || isNaN(y1) || isNaN(x2) || isNaN(y2)) continue;

    area += (x1 * width * y2 * height);
    area -= (x2 * width * y1 * height);
  }
  return Math.abs(area) / 2.0;
}

// Excel AI Assistant - Main Application Logic (adapted from addon test/app.js)
class ExcelAIAssistant {
  constructor() {
    this.config = {
      apiEndpoint: "https://api.openai.com/v1/chat/completions",
      apiKey: "", // Set via localStorage or window.OPENAI_API_KEY
      model: "gpt-4.1-mini", // Hardcoded model
      directActions: true, // Direct actions enabled
      promptId: "pmpt_68cc105e392c81948a63cef7980a3ac50002125793d33290", // Hardcoded Prompt ID
      promptVersion: "4", // Default prompt version (can be overridden)
    };
    this.chatHistory = [];
    this.currentSelection = "A1:C10";
    this.excelData = null;
    this.isProcessing = false;
    // Store last parsed action for manual apply
    this.lastParsedAction = null;

    // Default to hardcoded key; allow explicit override from window/localStorage if present
    const overrideKey = this.detectApiKey();
    if (overrideKey) {
      this.config.apiKey = overrideKey;
    }

    // Allow overriding prompt version via window/localStorage
    const overridePromptVersion = this.detectPromptVersion();
    if (overridePromptVersion) {
      this.config.promptVersion = String(overridePromptVersion);
    }
    // Expose a small helper to set the key at runtime via console
    window.setAIKey = (k) => {
      try {
        const v = String(k || "").trim();
        if (!v) throw new Error("Empty key");
        localStorage.setItem("ai_api_key", v);
        this.config.apiKey = v;
        this.updateStatus("API key saved", "success");
        return true;
      } catch (e) {
        console.error("Failed to save API key", e);
        this.updateStatus("Failed to save API key", "error");
        return false;
      }
    };

    // Expose helper to set prompt version at runtime
    window.setPromptVersion = (v) => {
      try {
        const sv = String(v || "").trim();
        if (!sv) throw new Error("Empty version");
        localStorage.setItem("ai_prompt_version", sv);
        this.config.promptVersion = sv;
        this.updateStatus(`Prompt version set to ${sv}`, "success");
        return true;
      } catch (e) {
        console.error("Failed to save prompt version", e);
        this.updateStatus("Failed to save prompt version", "error");
        return false;
      }
    };

    this.init();
  }

  detectApiKey() {
    try {
      const fromWin = (window.OPENAI_API_KEY || "").trim();
      if (fromWin) return fromWin;
      const fromLS = (localStorage.getItem("ai_api_key") || "").trim();
      if (fromLS) return fromLS;
    } catch {
      /* no-op */
    }
    return "";
  }

  detectPromptVersion() {
    try {
      const fromWin = (window.OPENAI_PROMPT_VERSION || "").trim();
      if (fromWin) return fromWin;
      const fromLS = (localStorage.getItem("ai_prompt_version") || "").trim();
      if (fromLS) return fromLS;
    } catch {
      /* no-op */
    }
    return "";
  }

  async init() {
    try {
      this.showLoading(false);
      await this.initializeOffice();

      // ── License gate ──
      const licenseValid = await checkLicense();
      if (!licenseValid) {
        this.showLicenseGate();
        return; // Block the rest of init until license is activated
      }
      this.hideLicenseGate();

      this.loadConfiguration();
      this.setupEventListeners();
      this.updateStatus("Ready", "success");
      // Try to update current selection from Excel if available
      this.tryUpdateSelectionFromExcel();

      // ── Check for updates (non-blocking) ──
      this.checkForAppUpdate();
    } catch (error) {
      console.error("Failed to initialize application:", error);
      this.updateStatus("Initialization failed", "error");
      this.showLoading(false);
    }
  }

  /** Check GitHub releases for a newer version and show the update banner if found. */
  async checkForAppUpdate() {
    try {
      const res = await fetch("/api/check-update");
      if (!res.ok) return;
      const data = await res.json();
      if (data.updateAvailable) {
        // Check if user already dismissed this version
        const dismissed = localStorage.getItem("update_dismissed_version");
        if (dismissed === data.latestVersion) return;

        this.showUpdateBanner(data);
      }
    } catch (e) {
      console.log("[Update] Client-side update check skipped:", e.message);
    }
  }

  /** Show the update notification banner. */
  showUpdateBanner(data) {
    const banner = document.getElementById("updateBanner");
    const textEl = document.getElementById("updateBannerText");
    const versionEl = document.getElementById("updateBannerVersion");
    const dismissBtn = document.getElementById("dismissUpdateBtn");

    if (!banner) return;

    if (textEl) textEl.textContent = `Update available! Restart the server to apply v${data.latestVersion}.`;
    if (versionEl) versionEl.textContent = `v${data.currentVersion} → v${data.latestVersion}`;

    banner.classList.remove("hidden");

    if (dismissBtn) {
      dismissBtn.addEventListener("click", () => {
        banner.classList.add("hidden");
        localStorage.setItem("update_dismissed_version", data.latestVersion);
      }, { once: true });
    }
  }

  showLicenseGate(reason) {
    // Stop periodic re-check while gate is visible
    this.stopLicenseRecheck();

    const gate = document.getElementById("licenseGate");
    const taskPane = document.querySelector(".task-pane");
    if (gate) gate.classList.remove("license-gate--hidden");
    if (taskPane) taskPane.style.display = "none";

    const activateBtn = document.getElementById("activateLicenseBtn");
    const keyInput = document.getElementById("licenseKeyInput");
    const errorEl = document.getElementById("licenseError");

    // Show a reason message if the license was revoked mid-session
    if (reason && errorEl) {
      errorEl.textContent = reason;
    }

    const doActivate = async () => {
      if (!keyInput || !activateBtn) return;
      const key = keyInput.value.trim();
      if (!key) {
        if (errorEl) errorEl.textContent = "Please enter a license key.";
        return;
      }
      activateBtn.disabled = true;
      activateBtn.textContent = "Validating…";
      if (errorEl) errorEl.textContent = "";

      const result = await activateLicense(key);
      if (result.valid) {
        this.hideLicenseGate();
        // Continue init if not already set up
        if (!this._appReady) {
          this.loadConfiguration();
          this.setupEventListeners();
          this._appReady = true;
        }
        this.updateStatus("Ready", "success");
        this.tryUpdateSelectionFromExcel();
      } else {
        activateBtn.disabled = false;
        activateBtn.textContent = "Activate";
        const messages = {
          INVALID: "Invalid license key. Please check and try again.",
          EXPIRED: "This license has expired.",
          RATE_LIMITED: "Too many attempts. Please wait and try again.",
          SERVER_ERROR: "Server error. Please try again later.",
          NETWORK_ERROR: "Network error. Check your connection.",
          EMPTY_KEY: "Please enter a license key.",
        };
        if (errorEl) errorEl.textContent = messages[result.status] || "Activation failed. Please try again.";
      }
    };

    if (activateBtn) {
      activateBtn.addEventListener("click", doActivate);
    }
    if (keyInput) {
      keyInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          doActivate();
        }
      });
      // Auto-focus
      setTimeout(() => keyInput.focus(), 300);
    }
  }

  hideLicenseGate() {
    const gate = document.getElementById("licenseGate");
    const taskPane = document.querySelector(".task-pane");
    if (gate) gate.classList.add("license-gate--hidden");
    if (taskPane) taskPane.style.display = "";
    // Start periodic re-check (every 5 minutes)
    this.startLicenseRecheck();
  }

  /** Re-validate the license every 5 minutes; lock app if invalid. */
  startLicenseRecheck() {
    this.stopLicenseRecheck(); // clear any existing interval
    const INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
    this._licenseRecheckId = setInterval(async () => {
      try {
        const valid = await checkLicense();
        if (!valid) {
          console.warn("License re-check failed — locking app.");
          clearLicense();
          this.showLicenseGate("Your license is no longer valid. It may have expired or been revoked. Please re-enter a valid key.");
        }
      } catch (e) {
        console.error("License re-check error:", e);
      }
    }, INTERVAL_MS);
  }

  /** Stop periodic license re-checks. */
  stopLicenseRecheck() {
    if (this._licenseRecheckId) {
      clearInterval(this._licenseRecheckId);
      this._licenseRecheckId = null;
    }
  }

  async initializeOffice() {
    // If Office.js is available, wait for Office readiness; otherwise, continue
    if (typeof Office !== "undefined" && Office.onReady) {
      try {
        await Office.onReady(async (info) => {
          if (info.host === Office.HostType.Excel) {
            // Assuming these elements exist in the HTML for the add-in
            const sideloadMsg = document.getElementById("sideload-msg");
            if (sideloadMsg) sideloadMsg.style.display = "none";
            const appBody = document.getElementById("app-body");
            if (appBody) appBody.style.display = "flex";

            // These event listeners are typically set up in setupEventListeners,
            // but if they are critical for Office.onReady, they can be here.
            // For this change, we'll assume they are handled elsewhere or are placeholders.
            // document.getElementById("runMhrCostCalc").onclick = runMhrCostCalc; // runMhrCostCalc is not defined in this scope
            // document.getElementById("runPriceCalc").onclick = runPriceCalc; // runPriceCalc is not defined in this scope

            // setupTabs(); // setupTabs is not defined in this scope
            await this.setupAutoCleanup(); // Call the new method
          }
        });
      } catch {
        // ignore
      }
    } else {
      // small delay similar to original simulated init
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
  }

  // Persistent event listener setup for auto-cleanup
  async setupAutoCleanup() {
    try {
      await Excel.run(async (context) => {
        const ws = context.workbook.worksheets.getActiveWorksheet();
        ws.onChanged.add(this.handleSheetChange.bind(this)); // Bind 'this' to the handler
        await context.sync();
        console.log("Auto-cleanup event listener registered.");
      });
    } catch (e) {
      console.error("Failed to register event listener", e);
    }
  }

  async handleSheetChange(event) {
    // Only process likely edits
    if (
      event.changeType === Excel.DataChangeType.rangeEdited ||
      event.changeType === Excel.DataChangeType.unknown
    ) {
      await Excel.run(async (context) => {
        const ws = context.workbook.worksheets.getActiveWorksheet();
        // Robust check: Ensure we are looking at the right address
        const rng = ws.getRange(event.address);
        rng.load("values");
        await context.sync();

        let dirty = false;
        const newValues = rng.values.map((row) => {
          return row.map((val) => {
            // Pattern: "Number || Description"
            if (typeof val === "string" && val.includes(" || ")) {
              const parts = val.split(" || ");
              const numPart = parts[0].trim();
              // Check if first part looks like a number (allow decimals)
              if (/^-?\d+(\.\d+)?$/.test(numPart)) {
                dirty = true;
                return numPart;
              }
            }
            return val;
          });
        });

        if (dirty) {
          // Turn off events temporarily to avoid infinite loop?
          // However, writing a simple number won't re-trigger the " || " check, so it terminates naturally.
          rng.values = newValues;
          await context.sync();
        }
      }).catch((err) => console.error("Error in auto-cleanup handler:", err));
    }
  }

  setupEventListeners() {
    const configToggle = document.getElementById("configToggle");
    const configContent = document.getElementById("configContent");
    if (configToggle && configContent) {
      configToggle.addEventListener("click", () => {
        configContent.classList.toggle("show");
      });
    }

    const saveConfigBtn = document.getElementById("saveConfig");
    if (saveConfigBtn) {
      saveConfigBtn.addEventListener("click", () => {
        this.saveConfiguration();
      });
    }

    const readSelectionBtn = document.getElementById("readSelection");
    if (readSelectionBtn) {
      readSelectionBtn.addEventListener("click", () => {
        this.readExcelSelection();
      });
    }

    const insertResponseBtn = document.getElementById("insertResponse");
    if (insertResponseBtn) {
      insertResponseBtn.addEventListener("click", () => {
        this.insertResponseToExcel();
      });
    }

    const applyLastActionBtn = document.getElementById("applyLastAction");
    if (applyLastActionBtn) {
      applyLastActionBtn.addEventListener("click", () => this.applyLastAction());
    }

    const chatInput = document.getElementById("chatInput");
    const sendBtn = document.getElementById("sendBtn");
    if (sendBtn) {
      sendBtn.addEventListener("click", () => {
        this.sendMessage();
      });
    }
    if (chatInput) {
      chatInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          this.sendMessage();
        }
      });
      chatInput.addEventListener("input", () => {
        this.autoResizeTextarea(chatInput);
      });
    }

    // Sidebar tab switching
    try {
      const sidebarButtons = Array.from(document.querySelectorAll(".sidebar-item .sidebar-link"));
      const sidebarItems = Array.from(document.querySelectorAll(".sidebar-item"));
      const panels = Array.from(document.querySelectorAll(".tab-panel"));
      const activate = (targetSel) => {
        if (!targetSel) return;
        const targetEl = document.querySelector(targetSel);
        if (!targetEl) return;
        panels.forEach((p) => p.classList.toggle("active", p === targetEl));
        sidebarItems.forEach((li) =>
          li.classList.toggle("active", li.getAttribute("data-target") === targetSel)
        );
      };
      sidebarButtons.forEach((btn) => {
        btn.addEventListener("click", () => {
          const li = btn.closest(".sidebar-item");
          if (!li) return;
          const target = li.getAttribute("data-target");
          activate(target);
        });
      });
      // Default to chat tab active
      const activeItem = document.querySelector(".sidebar-item.active");
      if (activeItem) activate(activeItem.getAttribute("data-target"));
    } catch {
      // non-fatal
    }

    // Settings: Direct actions toggle
    const chkDirectActions = document.getElementById("chkDirectActions");
    if (chkDirectActions) {
      chkDirectActions.checked = !!this.config.directActions;
      chkDirectActions.addEventListener("change", (e) => {
        this.config.directActions = !!e.target.checked;
        this.updateStatus(
          `Direct actions ${this.config.directActions ? "enabled" : "disabled"}`,
          "success"
        );
      });
    }

    // Settings: Prompt version display
    const pv = document.getElementById("promptVersionDisplay");
    if (pv) {
      pv.textContent = String(this.config.promptVersion || "");
    }

    // Sidebar: show app version
    const versionEl = document.getElementById("appVersionDisplay");
    if (versionEl) {
      versionEl.textContent = `v${APP_VERSION}`;
    }

    // Sidebar overlay open/close
    const toggleSidebarBtn = document.getElementById("toggleSidebar"); // close
    const openSidebarBtn = document.getElementById("openSidebarBtn"); // open
    const sidebarOverlay = document.querySelector(".sidebar--overlay");
    const sidebarBackdrop = document.getElementById("sidebarBackdrop");
    const setOverlayState = (open) => {
      if (!sidebarOverlay || !sidebarBackdrop) return;
      sidebarOverlay.classList.toggle("open", !!open);
      sidebarBackdrop.classList.toggle("hidden", !open);
      if (openSidebarBtn) openSidebarBtn.setAttribute("aria-expanded", String(!!open));
      if (toggleSidebarBtn) toggleSidebarBtn.setAttribute("aria-expanded", String(!!open));
    };
    if (openSidebarBtn) openSidebarBtn.addEventListener("click", () => setOverlayState(true));
    if (toggleSidebarBtn) toggleSidebarBtn.addEventListener("click", () => setOverlayState(false));
    if (sidebarBackdrop) sidebarBackdrop.addEventListener("click", () => setOverlayState(false));

    // Excel context show/hide
    const toggleExcelBtn = document.getElementById("toggleExcelContext");
    const excelContextPanel = document.getElementById("excelContextPanel");
    if (toggleExcelBtn && excelContextPanel) {
      toggleExcelBtn.addEventListener("click", () => {
        const hidden = excelContextPanel.classList.toggle("hidden");
        toggleExcelBtn.setAttribute("aria-expanded", String(!hidden));
      });
    }

    const clearChatBtn = document.getElementById("clearChat");
    if (clearChatBtn) {
      clearChatBtn.addEventListener("click", () => {
        this.clearChat();
      });
    }

    // Slot 3: Cable Filler run button
    const runCableFillerBtn = document.getElementById("runCableFiller");
    if (runCableFillerBtn) {
      runCableFillerBtn.addEventListener("click", async () => {
        try {
          const mod = await import(
            /* webpackChunkName: "feature-cable-filler" */ "./features/cableFiller.js"
          );
          await mod.runCableFillerOnActiveSheet(
            this.updateStatus.bind(this),
            this.addMessage.bind(this)
          );
        } catch (e) {
          console.error("Failed to load cable filler feature", e);
          this.updateStatus("Failed to load feature", "error");
        }
      });
    }

    // Slot 4: Insert template into new sheet
    // Templates dropdown: lazy-load feature module to keep bundle small
    const templateDropdownBtn = document.getElementById("templateDropdownBtn");
    if (templateDropdownBtn) {
      templateDropdownBtn.addEventListener(
        "click",
        async () => {
          try {
            const mod = await import(
              /* webpackChunkName: "feature-templates" */ "./features/templates.js"
            );
            // setupTemplateDropdown will rewire clicks and load menu on first open
            mod.setupTemplateDropdown(this.updateStatus.bind(this), this.addMessage.bind(this));
          } catch (e) {
            console.error("Failed to load templates feature", e);
            this.updateStatus("Failed to load feature", "error");
          }
        },
        { once: true }
      );
    }

    // Slot: Mhr & Cost calculator
    const runMhrCostBtn = document.getElementById("runMhrCostCalc");
    if (runMhrCostBtn) {
      runMhrCostBtn.addEventListener("click", async () => {
        try {
          if (runMhrCostBtn.disabled) return;
          const originalLabel = (runMhrCostBtn.textContent || "").trim();
          runMhrCostBtn.disabled = true;
          runMhrCostBtn.classList.add("btn--loading");
          runMhrCostBtn.setAttribute("aria-busy", "true");
          runMhrCostBtn.setAttribute(
            "data-original-label",
            originalLabel || "Process current sheet"
          );
          runMhrCostBtn.textContent = "Processing…";
          const mod = await import(
            /* webpackChunkName: "feature-mhr-cost" */ "./features/mhrCost.js"
          );
          await mod.runMhrCostOnActiveSheet(
            this.updateStatus.bind(this),
            this.addMessage.bind(this)
          );
        } catch (e) {
          console.error("Failed to load Mhr & Cost feature", e);
          this.updateStatus("Failed to load feature", "error");
        } finally {
          // Restore button state
          runMhrCostBtn.disabled = false;
          runMhrCostBtn.classList.remove("btn--loading");
          runMhrCostBtn.removeAttribute("aria-busy");
          // Flash a short 'Done' state, then restore original label
          try {
            const lbl =
              runMhrCostBtn.getAttribute("data-original-label") || "Process current sheet";
            runMhrCostBtn.textContent = "Done";
            setTimeout(() => {
              try {
                if (document.body.contains(runMhrCostBtn)) {
                  runMhrCostBtn.textContent = lbl;
                  runMhrCostBtn.removeAttribute("data-original-label");
                }
              } catch { }
            }, 600);
          } catch { }
        }
      });
    }

    const learnMhrBtn = document.getElementById("learnFromExcelBtn");
    if (learnMhrBtn) {
      learnMhrBtn.addEventListener("click", () => {
        this.learnFromExcel();
      });
    }

    // Slot: Price calculator (my2.db)
    const runPriceCalcBtn = document.getElementById("runPriceCalc");
    if (runPriceCalcBtn) {
      runPriceCalcBtn.addEventListener("click", async () => {
        try {
          if (runPriceCalcBtn.disabled) return;
          const originalLabel = (runPriceCalcBtn.textContent || "").trim();
          runPriceCalcBtn.disabled = true;
          runPriceCalcBtn.classList.add("btn--loading");
          runPriceCalcBtn.setAttribute("aria-busy", "true");
          runPriceCalcBtn.setAttribute("data-original-label", originalLabel || "Calculate Price");
          runPriceCalcBtn.textContent = "Processing…";
          const mod = await import(
            /* webpackChunkName: "feature-price-calc" */ "./features/priceCalc.js"
          );
          await mod.runPriceOnActiveSheet(this.updateStatus.bind(this), this.addMessage.bind(this));
        } catch (e) {
          console.error("Failed to load Price feature", e);
          this.updateStatus("Failed to load feature", "error");
        } finally {
          runPriceCalcBtn.disabled = false;
          runPriceCalcBtn.classList.remove("btn--loading");
          runPriceCalcBtn.removeAttribute("aria-busy");
          try {
            const lbl = runPriceCalcBtn.getAttribute("data-original-label") || "Calculate Price";
            runPriceCalcBtn.textContent = "Done";
            setTimeout(() => {
              try {
                if (document.body.contains(runPriceCalcBtn)) {
                  runPriceCalcBtn.textContent = lbl;
                  runPriceCalcBtn.removeAttribute("data-original-label");
                }
              } catch { }
            }, 600);
          } catch { }
        }
      });
    }

    // Slot: Formulas copy-to-clipboard
    const formulaButtons = Array.from(document.querySelectorAll(".formula-btn"));
    if (formulaButtons.length) {
      const copyToClipboard = async (text) => {
        try {
          if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(text);
            return true;
          }
        } catch { }
        try {
          const textarea = document.createElement("textarea");
          textarea.value = text;
          textarea.style.position = "fixed";
          textarea.style.opacity = "0";
          document.body.appendChild(textarea);
          textarea.select();
          document.execCommand("copy");
          document.body.removeChild(textarea);
          return true;
        } catch (e) {
          console.error("Clipboard copy failed", e);
          return false;
        }
      };
      const showToast = (msg) => {
        const el = document.getElementById("toast");
        if (!el) return;
        el.textContent = msg || "Copied to clipboard";
        el.classList.add("show");
        clearTimeout(this._toastTimer);
        this._toastTimer = setTimeout(() => el.classList.remove("show"), 1600);
      };
      formulaButtons.forEach((btn) => {
        btn.addEventListener("click", async () => {
          const fx = btn.getAttribute("data-formula") || "";
          if (!fx) return;
          const ok = await copyToClipboard(fx);
          if (ok) {
            this.updateStatus("Formula copied", "success");
            showToast("Copied!");
          } else {
            this.updateStatus("Copy failed", "error");
          }
        });
      });
    }

    // Config panel removed; no Import/Sync UI listeners

    // Reference Finder functionality
    const runReferenceSearchBtn = document.getElementById("runReferenceSearchBtn");
    if (runReferenceSearchBtn) {
      runReferenceSearchBtn.addEventListener("click", () => {
        this.runReferenceSearch();
      });
    }

    const browseFolderBtn = document.getElementById("browseFolderBtn");
    if (browseFolderBtn) {
      browseFolderBtn.addEventListener("click", async () => {
        try {
          this.updateStatus("Waiting for folder selection...", "processing");
          const response = await fetch("/api/browse-folder");
          if (!response.ok) throw new Error("Failed to open folder browser");
          const data = await response.json();
          if (data.folderPath) {
            const input = document.getElementById("referenceFolderPath");
            if (input) input.value = data.folderPath;
            this.updateStatus("Folder selected", "success");
          } else {
            this.updateStatus("Folder selection cancelled", "ready");
          }
        } catch (e) {
          console.error("Browse folder error:", e);
          this.updateStatus("Error browsing folder", "error");
        }
      });
    }

    // keyboard shortcuts
    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        const active = document.activeElement;
        if (active && active.id === "chatInput") this.sendMessage();
      }
      if (e.key === "Escape") {
        const active = document.activeElement;
        if (active && active.id === "chatInput") {
          active.value = "";
          this.autoResizeTextarea(active);
        }
      }
    });
    this.setupMtoListeners();

    // License: Deactivate button
    const deactivateBtn = document.getElementById("deactivateLicenseBtn");
    if (deactivateBtn) {
      deactivateBtn.addEventListener("click", () => {
        clearLicense();
        // Re-show the gate
        this.showLicenseGate();
      });
    }
  }

  loadConfiguration() {
    // Configuration is hardcoded now; nothing to load
  }

  saveConfiguration() {
    // Configuration is hardcoded now; nothing to save
  }

  updateConfigFromUI() {
    // Configuration is hardcoded now; ignore UI
  }

  updateConfigUI() {
    // Config panel removed; nothing to render
  }

  // Templates code moved to ./features/templates.js (lazy-loaded)

  async insertTemplateToNewSheet() {
    try {
      if (typeof Excel === "undefined") {
        this.updateStatus("Excel APIs not available", "error");
        return;
      }
      this.updateStatus("Creating template sheet...", "processing");
      await Excel.run(async (context) => {
        const wb = context.workbook;
        const sheets = wb.worksheets;
        sheets.load("items/name");
        await context.sync();

        // Generate a unique name like Template, Template 2, Template 3, ...
        const base = "Template";
        const existingNames = new Set(sheets.items.map((s) => s.name));
        let name = base;
        let i = 2;
        while (existingNames.has(name)) {
          name = `${base} ${i++}`;
        }

        const sheet = sheets.add(name);
        sheet.activate();

        // Prepare a simple template: headers + sample formulas
        const headers = ["Item", "Category", "Quantity", "Unit Price", "Total", "Notes"];
        const headerRange = sheet.getRange("A1:F1");
        headerRange.values = [headers];
        headerRange.format.fill.color = "#1f4e78"; // dark blue
        headerRange.format.font.color = "#ffffff";
        headerRange.format.font.bold = true;

        // Put some placeholder rows
        const rows = 20; // pre-allocate 20 rows
        const bodyRange = sheet.getRange(`A2:F${rows + 1}`);
        bodyRange.values = Array.from({ length: rows }, () => ["", "", 0, 0, 0, ""]);

        // Add formula for Total = Quantity * Unit Price in column E
        const totalRange = sheet.getRange(`E2:E${rows + 1}`);
        totalRange.formulas = Array.from({ length: rows }, (_, idx) => [
          [`=C${idx + 2}*D${idx + 2}`],
        ]);

        // Add a table for better UX
        const used = sheet.getUsedRange();
        used.load(["address"]);
        await context.sync();
        const table = wb.tables.add(used, true /*hasHeaders*/);
        table.name = name.replace(/\s+/g, "_") + "_Table";
        table.style = "TableStyleMedium9";

        // Freeze header row and set column widths
        sheet.freezePanes.freezeRows(1);
        sheet.getRange("A:F").format.columnWidth = 18;
        sheet.getRange("E:E").numberFormat = [["$#,##0.00"]];
        sheet.getRange("D:D").numberFormat = [["$#,##0.00"]];

        // Add a summary row below
        const sumRow = rows + 2;
        const sumRange = sheet.getRange(`D${sumRow}:E${sumRow}`);
        sumRange.values = [["Total:", 0]];
        sheet.getRange(`E${sumRow}`).formulas = [[`=SUM(E2:E${rows + 1})`]];
        sheet.getRange(`D${sumRow}:E${sumRow}`).format.font.bold = true;
        sheet.getRange(`E${sumRow}`).numberFormat = [["$#,##0.00"]];

        // Format categories as dropdown (data validation) - example list inline
        try {
          const categoryRange = sheet.getRange(`B2:B${rows + 1}`);
          categoryRange.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: '"Materials,Services,Travel,Other"',
            },
          };
        } catch {
          // Data validation may not be supported in some contexts
        }

        await context.sync();
      });
      this.updateStatus("Template sheet created", "success");
      this.addMessage("Inserted a new 'Template' worksheet with headers and formulas.", "ai");
    } catch (e) {
      console.error("Insert template failed", e);
      this.updateStatus("Failed to insert template", "error");
    }
  }

  // cable filler logic lazy-loaded from ./features/cableFiller.js

  async learnFromExcel() {
    try {
      if (typeof Excel === "undefined") {
        this.updateStatus("Excel context not found", "error");
        return;
      }
      this.updateStatus("Learning from Excel...", "processing");
      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "rowIndex", "rowCount", "columnIndex", "worksheet"]);
        await context.sync();
        const ws = range.worksheet;
        const values = range.values;
        if (!values || !values.length) {
          return { success: false, error: "Selection is empty" };
        }

        // Check if any cell has data
        let hasData = false;
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            if (values[r][c] !== "" && values[r][c] != null) {
              hasData = true;
              break;
            }
          }
          if (hasData) break;
        }
        if (!hasData) {
          return { success: false, error: "Selected cells are empty" };
        }

        const usedRange = ws.getUsedRange();
        usedRange.load(["address", "rowIndex", "rowCount", "columnCount", "values", "columnIndex"]);
        await context.sync();
        let descColIndex = -1;
        const uRows = usedRange.values;
        // Find header row in used range
        for (let r = 0; r < uRows.length; r++) {
          for (let c = 0; c < uRows[r].length; c++) {
            if (String(uRows[r][c]).trim().toLowerCase() === "description") {
              // c is relative to the start of the used range
              descColIndex = c + usedRange.columnIndex;
              break;
            }
          }
          if (descColIndex !== -1) break;
        }
        if (descColIndex === -1) {
          return { success: false, error: "Could not find 'Description' header" };
        }

        // For each row in the selection
        const itemsToLearn = [];
        const startRow = range.rowIndex;

        // We only care about the rows in the selection.
        // We assume the user selected the MHR Value column(s).
        // If multiple columns are selected, we'll take the first non-empty value in that row's selection?
        // Or strictly the first column of selection? Let's use the first column of selection for values.

        for (let i = 0; i < range.rowCount; i++) {
          const val = values[i][0]; // First column of selection
          if (val === "" || val == null) continue;

          // Row index in sheet
          const sheetRow = startRow + i;

          // Get description from the same row
          const descCell = ws.getCell(sheetRow, descColIndex);
          descCell.load("values");
          itemsToLearn.push({ row: sheetRow, val: val, descCell: descCell });
        }

        if (itemsToLearn.length === 0) {
          return { success: false, error: "No valid values found in selection" };
        }

        // Sync to get descriptions
        await context.sync();

        const payload = itemsToLearn.map(item => ({
          value: item.val,
          description: item.descCell.values[0][0]
        })).filter(p => p.description && String(p.description).trim() !== "");

        return { success: true, payload };
      });

      if (!result.success) {
        this.updateStatus(result.error, "error");
        this.addMessage(`Error: ${result.error}`, "ai");
        return;
      }

      const payload = result.payload;
      if (payload.length === 0) {
        this.updateStatus("No descriptions found", "error");
        this.addMessage("Found values but no corresponding descriptions.", "ai");
        return;
      }

      let successCount = 0;
      let failCount = 0;

      this.updateStatus(`Learning ${payload.length} items...`, "processing");

      // Send them sequentially or parallel? Parallel is faster.
      await Promise.all(payload.map(async (item) => {
        try {
          const response = await fetch("/api/mhr/learn", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ description: item.description, value: item.value }),
          });
          if (response.ok) {
            const json = await response.json();
            if (json.success) successCount++;
            else failCount++;
          } else {
            failCount++;
          }
        } catch (e) {
          failCount++;
        }
      }));

      if (successCount > 0) {
        this.updateStatus(`Learned ${successCount} items!`, "success");
        this.addMessage(`Successfully learned ${successCount} items. ${failCount > 0 ? `(${failCount} failed)` : ""}`, "ai");
      } else {
        this.updateStatus("Learning failed", "error");
        this.addMessage("Failed to learn any items.", "ai");
      }
    } catch (e) {
      console.error("Learn error", e);
      this.updateStatus("Learn failed", "error");
      this.addMessage("Learn operation failed. See console.", "ai");
    }
  }

  async tryUpdateSelectionFromExcel() {
    try {
      if (typeof Excel === "undefined") return;
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        this.currentSelection = range.address || this.currentSelection;
        const el = document.getElementById("selectedRange");
        if (el) el.textContent = this.currentSelection;
      });
    } catch {
      // ignore if not available
    }
  }

  async readExcelSelection() {
    try {
      this.updateStatus("Reading Excel data...", "processing");
      // Attempt real Excel read; fallback to simulation
      if (typeof Excel !== "undefined") {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["address", "values"]);
          await context.sync();
          this.currentSelection = range.address;
          this.excelData = { range: range.address, values: range.values };
          const el = document.getElementById("selectedRange");
          if (el) el.textContent = this.currentSelection;
        });
      } else {
        await this.simulateExcelRead();
      }

      const message = `I've read the data from ${this.currentSelection}. Here's what I found:`;
      this.addMessage(message, "ai");
      this.updateStatus("Data read successfully", "success");
    } catch (error) {
      console.error("Failed to read Excel data:", error);
      this.updateStatus("Failed to read Excel data", "error");
    }
  }

  async simulateExcelRead() {
    return new Promise((resolve) => {
      setTimeout(() => {
        this.excelData = {
          range: this.currentSelection,
          values: [
            ["Product", "Revenue", "Units"],
            ["Product A", 8500, 120],
            ["Product B", 12400, 85],
            ["Product C", 6800, 200],
          ],
        };
        resolve();
      }, 400);
    });
  }

  async insertResponseToExcel() {
    try {
      if (!this.chatHistory.length) {
        this.updateStatus("No response to insert", "error");
        return;
      }
      this.updateStatus("Inserting to Excel...", "processing");
      const last = this.chatHistory[this.chatHistory.length - 1];
      if (typeof Excel !== "undefined") {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.values = [[last.text]];
          await context.sync();
        });
      } else {
        await this.simulateExcelWrite();
      }
      this.updateStatus("Response inserted successfully", "success");
      this.addMessage("I've inserted the analysis results into your Excel worksheet.", "ai");
    } catch (error) {
      console.error("Failed to insert to Excel:", error);
      this.updateStatus("Failed to insert to Excel", "error");
    }
  }

  async simulateExcelWrite() {
    return new Promise((resolve) => setTimeout(resolve, 300));
  }

  async sendMessage() {
    const input = document.getElementById("chatInput");
    if (!input) return;
    const message = input.value.trim();
    if (!message || this.isProcessing) return;

    try {
      // Clear input UI immediately
      input.value = "";
      this.autoResizeTextarea(input);
      // Echo user message
      this.addMessage(message, "user");

      // Try to parse and (optionally) execute from the user's message first
      {
        const userAction = this.parseExcelAction(message);
        if (userAction) {
          this.lastParsedAction = userAction;
          if (this.config.directActions && typeof Excel !== "undefined") {
            try {
              await this.executeExcelAction(userAction);
              // For read actions, executeExcelAction posts the result message itself
              if (userAction.type === "read" || userAction.type === "readRange") {
                this.updateStatus("Read completed", "success");
                return;
              }
              // For write/formula actions, report a concise confirmation
              this.addMessage(
                `Done. I wrote ${userAction.type === "formula" ? "the formula" : '"' + userAction.value + '"'} to ${userAction.address || userAction.range || "selection"}.`,
                "ai"
              );
              this.updateStatus("Action executed", "success");
              return; // No need to call API when direct action succeeded
            } catch (e) {
              console.error("Direct action from user message failed", e);
              // continue to API flow
            }
          } else {
            this.updateStatus(
              "Parsed action ready. Click 'Apply last action' to run it.",
              "success"
            );
          }
        }
      }

      if (!this.config.apiKey) {
        this.addMessage(
          "I can run direct Excel actions without an AI key. To enable AI replies, set your OpenAI key by running window.setAIKey('sk-...') in the Console, or store it in localStorage as 'ai_api_key'.",
          "ai"
        );
        this.updateStatus("API key missing (direct actions still work)", "error");
        return;
      }

      this.isProcessing = true;
      this.updateStatus("Processing your request...", "processing");
      this.showLoading(true);
      const response = await this.callGPTAPI(message);
      // Try to parse for direct Excel action if allowed
      {
        const action = this.parseExcelAction(response);
        if (action) {
          this.lastParsedAction = action;
          if (this.config.directActions && typeof Excel !== "undefined") {
            try {
              await this.executeExcelAction(action);
              if (action.type === "read" || action.type === "readRange") {
                // For read, the execution already posted the readout; still show the AI response above
                this.addMessage(response, "ai");
                this.updateStatus("Read completed", "success");
                return;
              }
              this.addMessage(
                `${response}\n\n(I've executed this in your sheet: ${action.type} ${action.value ? '"' + action.value + '" ' : ""}at ${action.address || action.range
                })`,
                "ai"
              );
              this.updateStatus("Action executed", "success");
              return;
            } catch (e) {
              console.error("Failed to execute parsed action", e);
              // fall through to plain message
            }
          } else {
            this.updateStatus(
              "Parsed action ready from AI. Click 'Apply last action' to run it.",
              "success"
            );
          }
        }
      }
      this.addMessage(response, "ai");
      this.updateStatus("Response received", "success");
    } catch (error) {
      console.error("Failed to send message:", error);
      const msg = error && error.message ? error.message : "Unknown error";
      this.addMessage(`Error: ${msg}`, "ai");
      this.updateStatus(msg, "error");
    } finally {
      this.isProcessing = false;
      this.showLoading(false);
    }
  }

  async applyLastAction() {
    try {
      if (!this.lastParsedAction) {
        this.updateStatus("No parsed action available", "error");
        return;
      }
      if (typeof Excel === "undefined") {
        this.updateStatus("Excel APIs not available", "error");
        return;
      }
      await this.executeExcelAction(this.lastParsedAction);
      const where = this.lastParsedAction.address || this.lastParsedAction.range || "selection";
      this.updateStatus(`Applied action at ${where}`, "success");
      this.addMessage(`Applied last action at ${where}.`, "ai");
    } catch (e) {
      console.error("Apply last action failed", e);
      this.updateStatus("Failed to apply last action", "error");
    }
  }

  // Parse simple actionable instructions like:
  // - write "Hi" to F9
  // - put 123 in A1
  // - enter formula =SUM(A1:A5) into B1
  // Returns one of:
  //  - { type: 'write'|'formula', address: 'F9', value: 'Hi'|'=SUM(...)' }
  //  - { type: 'rangeFill', range: 'A1:C3', value: '1' }
  //  - { type: 'rangeWrite', range: 'A1:B2', values: [[...],[...]] }
  //  - { type: 'read', address: 'B9' } or { type: 'readRange', range: 'A1:C3' }
  parseExcelAction(text) {
    if (!text) return null;
    // const lowered = text.toLowerCase();

    // 1) write/put/type "value" to/into/in CELL
    const writeQuoted =
      /(write|put|type|enter)\s+["“”]([\s\S]*?)["“”]\s+(to|into|in)\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})/i;
    const m1 = text.match(writeQuoted);
    if (m1) return { type: "write", value: m1[2], address: m1[4].toUpperCase() };

    // 2) write/put/type VALUE (unquoted single word/number) to/into/in CELL
    const writeBare =
      /(write|put|type|enter)\s+([^"\n\r\t]+?)\s+(to|into|in)\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})/i;
    const m2 = text.match(writeBare);
    if (m2) return { type: "write", value: m2[2].trim(), address: m2[4].toUpperCase() };

    // 3) formula cases: formula ... into CELL or set CELL to =...
    const formulaInto =
      /(formula|set)\s+(?:=)?([^\n\r]+?)\s+(?:into|in|to)\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})/i;

    // 2b) set CELL to "value" (plain write)
    const setToQuoted = /set\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})\s+to\s+["“”]([\s\S]*?)["“”]/i;
    const m2b = text.match(setToQuoted);
    if (m2b) return { type: "write", value: m2b[2], address: m2b[1].toUpperCase() };

    // 2c) set CELL to VALUE (bare)
    const setToBare = /set\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})\s+to\s+([^"\n\r\t]+)/i;
    const m2c = text.match(setToBare);
    if (m2c) return { type: "write", value: m2c[2].trim(), address: m2c[1].toUpperCase() };
    const m3 = text.match(formulaInto);
    if (m3) {
      const val = m3[2].trim();
      const formula = val.startsWith("=") ? val : "=" + val;
      return { type: "formula", value: formula, address: m3[3].toUpperCase() };
    }

    // 4) If the text contains explicit steps like "Click on cell F9 ... type \"Hi\"" try to infer
    const stepWrite = /cell\s+([a-z]{1,3}[0-9]{1,7}).*?type\s+["“”]([\s\S]*?)["“”]/i;
    const m4 = text.match(stepWrite);
    if (m4) return { type: "write", value: m4[2], address: m4[1].toUpperCase() };

    // 5) Range fills: "fill A1:C3 with 1"
    const fillRange =
      /(fill|populate)\s+(?:range\s+)?([a-z]{1,3}[0-9]{1,7}:[a-z]{1,3}[0-9]{1,7})\s+with\s+["“”]?([^\n\r]+?)["“”]?/i;
    const m5 = text.match(fillRange);
    if (m5) return { type: "rangeFill", range: m5[2].toUpperCase(), value: m5[3].trim() };

    // 6) Range write with matrix: "write [[1,2],[3,4]] to A1:B2"
    const writeMatrix =
      /(write|put|enter)\s+(\[\[[\s\S]*?\]\])\s+(?:to|into|in)\s+(?:range\s+)?([a-z]{1,3}[0-9]{1,7}:[a-z]{1,3}[0-9]{1,7})/i;
    const m6 = text.match(writeMatrix);
    if (m6) {
      const matrix = this.tryParseMatrix(m6[2]);
      if (matrix) return { type: "rangeWrite", range: m6[3].toUpperCase(), values: matrix };
    }

    // 7) set range A1:C3 to 1
    const setRangeTo =
      /set\s+(?:range\s+)?([a-z]{1,3}[0-9]{1,7}:[a-z]{1,3}[0-9]{1,7})\s+to\s+["“”]?([^\n\r]+?)["“”]?/i;
    const m7 = text.match(setRangeTo);
    if (m7) return { type: "rangeFill", range: m7[1].toUpperCase(), value: m7[2].trim() };

    // 8) Read single cell: "read cell B9" / "what's in B9" / "get B9"
    const readCell = /(read|get|what\s*'?s\s*in)\s+(?:cell\s+)?([a-z]{1,3}[0-9]{1,7})/i;
    const m8 = text.match(readCell);
    if (m8) return { type: "read", address: m8[2].toUpperCase() };

    // 9) Read range: "read range A1:C3" / "get A1:C3"
    const readRange = /(read|get)\s+(?:range\s+)?([a-z]{1,3}[0-9]{1,7}:[a-z]{1,3}[0-9]{1,7})/i;
    const m9 = text.match(readRange);
    if (m9) return { type: "readRange", range: m9[2].toUpperCase() };

    // Otherwise, not recognized
    return null;
  }

  tryParseMatrix(str) {
    try {
      // Normalize quotes
      const normalized = str.replace(/“|”/g, '"').replace(/'/g, '"');
      const arr = JSON.parse(normalized);
      if (Array.isArray(arr) && arr.every((r) => Array.isArray(r))) {
        // Coerce numbers
        return arr.map((row) =>
          row.map((v) => {
            const num = Number(v);
            return !isNaN(num) && /^-?\d+(\.\d+)?$/.test(String(v)) ? num : v;
          })
        );
      }
    } catch {
      /* no-op */
    }
    return null;
  }

  async executeExcelAction(action) {
    if (!action || typeof Excel === "undefined") return;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      if (action.type === "read" && action.address) {
        const range = sheet.getRange(action.address);
        range.load(["values", "address"]);
        await context.sync();
        const value = Array.isArray(range.values) && range.values[0] ? range.values[0][0] : "";
        this.addMessage(
          `Cell ${range.address} = ${value === undefined ? "" : String(value)}`,
          "ai"
        );
        return;
      }

      if (action.type === "readRange" && action.range) {
        const rng = sheet.getRange(action.range);
        rng.load(["values", "address", "rowCount", "columnCount"]);
        await context.sync();
        const preview = (rng.values || [])
          .slice(0, 10)
          .map((row) => row.join("\t"))
          .join("\n");
        this.addMessage(
          `Range ${rng.address} (${rng.rowCount}x${rng.columnCount}):\n${preview}`,
          "ai"
        );
        return;
      }

      if (action.address) {
        const range = sheet.getRange(action.address);
        if (action.type === "formula") {
          range.formulas = [[action.value]];
        } else {
          let v = action.value;
          const num = Number(v);
          if (!isNaN(num) && /^-?\d+(\.\d+)?$/.test(String(v))) v = num;
          range.values = [[v]];
        }
        await context.sync();
        return;
      }

      if (action.type === "rangeFill" && action.range) {
        const rng = sheet.getRange(action.range);
        rng.load(["rowCount", "columnCount"]);
        await context.sync();
        let v = action.value;
        const num = Number(v);
        if (!isNaN(num) && /^-?\d+(\.\d+)?$/.test(String(v))) v = num;
        const values = Array.from({ length: rng.rowCount }, () =>
          Array.from({ length: rng.columnCount }, () => v)
        );
        rng.values = values;
        await context.sync();
        return;
      }

      if (action.type === "rangeWrite" && action.range && action.values) {
        const rng = sheet.getRange(action.range);
        rng.load(["rowCount", "columnCount"]);
        await context.sync();
        const rows = rng.rowCount,
          cols = rng.columnCount;
        const padded = [];
        for (let r = 0; r < rows; r++) {
          const srcRow = action.values[r] || [];
          const dstRow = [];
          for (let c = 0; c < cols; c++) {
            let cell = srcRow[c];
            const num = Number(cell);
            if (!isNaN(num) && /^-?\d+(\.\d+)?$/.test(String(cell))) cell = num;
            dstRow.push(cell !== undefined ? cell : "");
          }
          padded.push(dstRow);
        }
        rng.values = padded;
        await context.sync();
        return;
      }
    });
  }

  async callGPTAPI(userMessage) {
    const systemPrompt =
      "You are an AI assistant specialized in Excel and data analysis. You help users with formulas, data analysis, charts, and spreadsheet automation. Be concise but thorough in your responses.";
    try {
      // Build chat-style messages once
      let contextMessage = userMessage;
      if (this.excelData) {
        contextMessage = `Based on the Excel data in range ${this.excelData.range}:\n${JSON.stringify(
          this.excelData.values,
          null,
          2
        )}\n\nUser question: ${userMessage}`;
      }
      // Build a compact conversation transcript (last 6 messages) for context
      const historyText = this.chatHistory
        .slice(-6)
        .map((m) => `${m.type === "user" ? "User" : "Assistant"}: ${m.text}`)
        .join("\n");
      const messages = [
        { role: "system", content: systemPrompt },
        {
          role: "user",
          content: historyText ? `${historyText}\n\nUser: ${contextMessage}` : contextMessage,
        },
      ];

      // Primary: Responses API with prompt ID
      try {
        const responsesEndpoint = "https://api.openai.com/v1/responses";
        const responseBody = {
          model: this.config.model,
          prompt: {
            id: this.config.promptId,
            version: this.config.promptVersion || "1",
          },
          input: messages[messages.length - 1].content,
          instructions: systemPrompt,
          temperature: 0.7,
          max_output_tokens: 1500,
        };
        const res = await fetch(responsesEndpoint, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${this.config.apiKey}`,
          },
          body: JSON.stringify(responseBody),
        });
        if (!res.ok) {
          const errorData = await res.json().catch(() => ({}));
          const msg =
            errorData.error?.message ||
            errorData.message ||
            `API request failed with status ${res.status}`;
          throw new Error(msg);
        }
        const data = await res.json();
        let text = "";
        if (typeof data?.output_text === "string") {
          text = data.output_text;
        } else if (data?.choices && data.choices[0]?.message?.content) {
          text = data.choices[0].message.content;
        } else if (Array.isArray(data?.output)) {
          const parts = data.output
            .map((p) => {
              if (typeof p === "string") return p;
              if (Array.isArray(p?.content)) {
                return p.content
                  .map((c) => (typeof c === "string" ? c : c?.text || c?.content || ""))
                  .join(" ");
              }
              return p?.content || p?.text || "";
            })
            .join(" ");
          text = parts.trim();
        }
        if (!text) throw new Error("Invalid API response format");
        return text.trim();
      } catch (primaryErr) {
        // Fallback: Chat Completions (without prompt ID)
        console.warn("Responses API failed, falling back to Chat Completions:", primaryErr);
        const ccEndpoint = "https://api.openai.com/v1/chat/completions";
        const ccBody = {
          model: this.config.model,
          messages,
          temperature: 0.7,
        };
        const res2 = await fetch(ccEndpoint, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${this.config.apiKey}`,
          },
          body: JSON.stringify(ccBody),
        });
        if (!res2.ok) {
          const err2 = await res2.json().catch(() => ({}));
          const msg2 =
            err2.error?.message || err2.message || `API request failed with status ${res2.status}`;
          throw new Error(msg2);
        }
        const data2 = await res2.json();
        const text2 = data2?.choices?.[0]?.message?.content;
        if (!text2) throw new Error("Invalid Chat Completions response format");
        return String(text2).trim();
      }
    } catch (error) {
      console.error("GPT API call failed:", error);
      if (
        String(error.message).includes("401") ||
        String(error.message).toLowerCase().includes("unauthorized")
      ) {
        throw new Error("Invalid API key. Please check your configuration.");
      } else if (String(error.message).includes("429")) {
        throw new Error("Rate limit exceeded. Please try again later.");
      } else if (String(error.message).toLowerCase().includes("fetch")) {
        throw new Error("Network error. Please check your internet connection.");
      }
      throw error;
    }
  }

  addMessage(text, type) {
    const messagesContainer = document.getElementById("chatMessages");
    if (!messagesContainer) return;
    const messageDiv = document.createElement("div");
    messageDiv.className = `message ${type}-message`;
    const avatar = document.createElement("div");
    avatar.className = `message-avatar ${type}-avatar`;
    const avatarIcon =
      type === "user"
        ? '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>'
        : '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>';
    avatar.innerHTML = avatarIcon;
    const content = document.createElement("div");
    content.className = "message-content";
    const messageText = document.createElement("div");
    messageText.className = "message-text";
    messageText.textContent = text;
    const messageTime = document.createElement("div");
    messageTime.className = "message-time";
    messageTime.textContent = new Date().toLocaleTimeString();
    content.appendChild(messageText);
    content.appendChild(messageTime);
    messageDiv.appendChild(avatar);
    messageDiv.appendChild(content);
    messagesContainer.appendChild(messageDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
    this.chatHistory.push({ text, type, timestamp: new Date() });
  }

  clearChat() {
    const messagesContainer = document.getElementById("chatMessages");
    if (!messagesContainer) return;
    messagesContainer.innerHTML = `
      <div class="message ai-message">
        <div class="message-avatar ai-avatar">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="5"></circle>
            <line x1="12" y1="1" x2="12" y2="3"></line>
            <line x1="12" y1="21" x2="12" y2="23"></line>
            <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
            <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
            <line x1="1" y1="12" x2="3" y2="12"></line>
            <line x1="21" y1="12" x2="23" y2="12"></line>
            <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
            <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
          </svg>
        </div>
        <div class="message-content">
          <div class="message-text">Chat cleared! I'm ready to help you with your Excel data analysis and questions.</div>
          <div class="message-time">${new Date().toLocaleTimeString()}</div>
        </div>
      </div>
    `;
    this.chatHistory = [];
    this.updateStatus("Chat cleared", "success");
  }

  updateStatus(text, type = "ready") {
    const statusIndicator = document.getElementById("statusIndicator");
    if (!statusIndicator) return;
    const statusText = statusIndicator.querySelector(".status-text");
    if (!statusText) return;
    statusText.textContent = text;
    statusIndicator.classList.remove("processing", "error", "success");
    if (type !== "ready") statusIndicator.classList.add(type);
    if (type !== "error" && type !== "processing") {
      setTimeout(() => {
        if (statusText.textContent === text) {
          statusText.textContent = "Ready";
          statusIndicator.classList.remove("processing", "error", "success");
        }
      }, 3000);
    }
  }

  showLoading(show) {
    const overlay = document.getElementById("loadingOverlay");
    const sendBtn = document.getElementById("sendBtn");
    if (overlay) overlay.classList.toggle("hidden", !show);
    if (sendBtn) sendBtn.disabled = show;
  }

  autoResizeTextarea(textarea) {
    if (!textarea) return;
    textarea.style.height = "auto";
    const newHeight = Math.min(textarea.scrollHeight, 120);
    textarea.style.height = newHeight + "px";
  }

  setupMtoListeners() {
    const syncBtn = document.getElementById("mtoSyncNowBtn");
    const liveSyncChk = document.getElementById("mtoLiveSync");

    if (syncBtn) {
      syncBtn.addEventListener("click", () => {
        this.syncMtoData(true);
      });
    }

    if (liveSyncChk) {
      liveSyncChk.addEventListener("change", (e) => {
        if (e.target.checked) {
          // Start polling
          this.syncMtoData(true); // Initial sync with status
          this.mtoTimer = setInterval(() => this.syncMtoData(false), 2000); // Silent background sync
          this.updateMtoStatus("Live Sync Enabled", "success");
        } else {
          if (this.mtoTimer) {
            clearInterval(this.mtoTimer);
            this.mtoTimer = null;
          }
          this.updateMtoStatus("Live Sync Disabled");
        }
      });

      // Auto-enable Live Sync by default for better UX
      liveSyncChk.checked = true;
      liveSyncChk.dispatchEvent(new Event("change"));
    }
  }


  async syncMtoData(manual = false) {
    if (this.isMtoSyncing) return; // Prevent overlapping syncs

    const urlInput = document.getElementById("mtoServerUrl");
    const url = urlInput ? urlInput.value.replace(/\/$/, "") : "http://localhost:8010";

    if (manual) {
      this.updateMtoStatus("Syncing...", "processing");
    }

    this.isMtoSyncing = true;

    try {
      const res = await fetch(`${url}/api/store`);
      if (!res.ok) throw new Error("Connection failed");

      const data = await res.json();
      // data should contain: annotations, measurements, scale, image_width, image_height
      if (!data) {
        if (manual) this.updateMtoStatus("No data found on server");
        return;
      }

      const annCount = data.annotations ? data.annotations.length : 0;
      const measCount = data.measurements ? data.measurements.length : 0;
      const totalCount = annCount + measCount;

      // Robust check: compare content (signature), not just count
      // This ensures that if a measurement is resized (but count is same), it still syncs.
      const signature = JSON.stringify(data.annotations || []) +
        JSON.stringify(data.measurements || []) +
        JSON.stringify(data.scale || {}) +
        JSON.stringify(data.drawings || []);

      // Only write if manual or if content changed 
      if (manual || this.lastMtoSignature !== signature) {
        await this.writeMtoToExcel(data);
        this.lastMtoSignature = signature;
        this.updateMtoStatus(`Synced ${totalCount} items`, "success");
      }

    } catch (e) {
      console.error("MTO Sync Error", e);
      if (manual) this.updateMtoStatus("Sync Failed: " + e.message, "error");
    } finally {
      this.isMtoSyncing = false;
    }
  }

  calculateLength(points, width, height) {
    if (!points || points.length < 2) return 0;
    let total = 0;
    for (let i = 1; i < points.length; i++) {
      const p1 = points[i - 1];
      const p2 = points[i];
      const dx = (p2.x - p1.x) * width;
      const dy = (p2.y - p1.y) * height;
      total += Math.sqrt(dx * dx + dy * dy);
    }
    return total;
  }

  calculateArea(points, width, height) {
    if (!points || points.length < 3) return 0;
    let area = 0;
    for (let i = 0; i < points.length; i++) {
      const p1 = points[i];
      const p2 = points[(i + 1) % points.length];
      area += (p1.x * width * p2.y * height);
      area -= (p2.x * width * p1.y * height);
    }
    return Math.abs(area) / 2.0;
  }



  async writeMtoToExcel(data) {
    if (typeof Excel === "undefined") return;

    await Excel.run(async (context) => {
      // Check/Create "MTO Data" sheet
      let mtoSheet;
      const sheets = context.workbook.worksheets;
      try {
        mtoSheet = sheets.getItem("MTO Data");
        mtoSheet.load("name");
        await context.sync();
      } catch (e) {
        mtoSheet = sheets.add("MTO Data");
      }

      mtoSheet.activate();

      // Prepare data
      // New Schema: Drawing Number, Label, Type, Value, Unit, Confidence, Color
      const headers = ["Drawing Number", "Label", "Type", "Value", "Unit", "Confidence", "Color"];
      const rows = [];

      // Helper to process a single drawing's data
      const processDrawing = (drawing, drawingNumber) => {
        const imageWidth = drawing.image_width || 1000;
        const imageHeight = drawing.image_height || 1000;
        const scaleFactor = (drawing.scale && drawing.scale.unitsPerPixel) ? drawing.scale.unitsPerPixel : 1.0;
        const scaleUnit = (drawing.scale && drawing.scale.unit) ? drawing.scale.unit : "px";

        // 1. Annotations (Counts) - Aggregated
        if (Array.isArray(drawing.annotations)) {
          const groups = {};
          drawing.annotations.forEach(ann => {
            const label = ann.label || "Count Item";
            const color = ann.color || "";
            const key = label + "||" + color;

            if (!groups[key]) {
              groups[key] = { count: 0, confidences: [], label, color };
            }
            groups[key].count++;
            if (ann.confidence != null) {
              groups[key].confidences.push(ann.confidence);
            }
          });

          // Convert groups to rows
          Object.values(groups).forEach(g => {
            let avgConf = "";
            if (g.confidences.length > 0) {
              const sum = g.confidences.reduce((a, b) => a + b, 0);
              avgConf = (sum / g.confidences.length).toFixed(2);
            }

            rows.push([
              drawingNumber,
              g.label,
              "Count",
              g.count,
              "count",
              avgConf,
              g.color
            ]);
          });
        }

        // 2. Measurements - Aggregated
        if (Array.isArray(drawing.measurements)) {
          const mGroups = {};

          drawing.measurements.forEach(m => {
            try {
              const type = (m.type || "").toLowerCase();
              let val = 0;
              let unit = "";
              let typeName = "";

              if (type === "length" || type === "line") {
                const pxVal = calculateLengthHelper(m.points, imageWidth, imageHeight);
                val = pxVal * scaleFactor;
                unit = scaleUnit;
                typeName = "Length";
              } else if (type === "area" || type === "polygon") {
                const pxVal = calculateAreaHelper(m.points, imageWidth, imageHeight);
                val = pxVal * (scaleFactor * scaleFactor);
                unit = "sq " + scaleUnit;
                typeName = "Area";
              } else {
                return; // Skip unknown types
              }

              const label = m.label || (typeName + " Item");
              const color = m.color || "";

              // Group Key
              const key = label + "||" + typeName + "||" + unit + "||" + color;

              if (!mGroups[key]) {
                mGroups[key] = {
                  value: 0.0,
                  label,
                  typeName,
                  unit,
                  color,
                  count: 0
                };
              }
              mGroups[key].value += val;
              mGroups[key].count++;

            } catch (err) {
              console.error("Error processing measurement item:", m, err);
            }
          });

          // Output Aggregated Rows
          Object.values(mGroups).forEach(g => {
            rows.push([
              drawingNumber,
              g.label,
              g.typeName,
              Number(g.value.toFixed(2)),
              g.unit,
              "", // Confidence not applicable to measurements usually
              g.color
            ]);
          });
        }
      };

      if (Array.isArray(data.drawings) && data.drawings.length > 0) {
        // Process all drawings
        data.drawings.forEach(d => {
          processDrawing(d, d.drawing_number || "Unknown");
        });
      } else {
        // Fallback to legacy single page structure (annotations at root)
        // We don't have a specific drawing number in legacy, assume "Current"
        if (data.annotations || data.measurements) {
          processDrawing(data, "Current Drawing");
        }
      }

      // Clear previous
      mtoSheet.getUsedRange().clear();

      // Write new
      if (rows.length > 0) {
        const range = mtoSheet.getRange("A1").getResizedRange(rows.length, headers.length - 1);
        range.values = [headers, ...rows];

        // Format header
        const headerRange = mtoSheet.getRange("A1").getResizedRange(0, headers.length - 1);
        headerRange.format.fill.color = "#444444";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;

        mtoSheet.getUsedRange().format.autofitColumns();
      } else {
        // Just headers
        const range = mtoSheet.getRange("A1").getResizedRange(0, headers.length - 1);
        range.values = [headers];
      }

      await context.sync();
    });
  }

  updateMtoStatus(msg, type = "") {
    const el = document.getElementById("mtoStatus");
    if (el) {
      el.querySelector(".status-text").textContent = msg;
      el.className = "status-indicator " + type;
    }
  }

  async runReferenceSearch() {
    const inputEl = document.getElementById("referenceFolderPath");
    const statusEl = document.getElementById("referenceStatus");
    const setStatus = (msg, type) => {
      if (statusEl) {
        const textEl = statusEl.querySelector(".status-text");
        if (textEl) textEl.textContent = msg;
        statusEl.className = "status-indicator " + (type || "ready");
      }
    };

    if (!inputEl) return;
    let folderPath = inputEl.value.trim();
    // Strip surrounding quotes if the user pasted them from Windows Explorer
    folderPath = folderPath.replace(/^["']|["']$/g, '');

    if (!folderPath) {
      setStatus("Please enter a folder path", "error");
      return;
    }

    if (typeof Excel === "undefined") {
      setStatus("Excel APIs not available", "error");
      return;
    }

    setStatus("Reading selected cells...", "processing");
    this.showLoading(true);

    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
        await context.sync();

        const searchTerms = [];
        const termData = [];

        for (let r = 0; r < range.rowCount; r++) {
          for (let c = 0; c < range.columnCount; c++) {
            const val = range.values[r][c];
            if (val && String(val).trim() !== "") {
              const term = String(val).trim();
              if (!searchTerms.includes(term)) {
                searchTerms.push(term);
              }
              termData.push({ term, row: r, col: c });
            }
          }
        }

        if (searchTerms.length === 0) {
          setStatus("No text found in selection", "error");
          this.showLoading(false);
          return;
        }

        setStatus(`Searching for ${searchTerms.length} terms...`, "processing");

        const response = await fetch("/api/reference-search", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ folderPath, searchTerms })
        });

        if (!response.ok) {
          throw new Error("Search request failed");
        }

        const data = await response.json();
        const results = data.results || {};
        let linkCount = 0;

        for (const item of termData) {
          const matchedFiles = results[item.term];
          if (matchedFiles && matchedFiles.length > 0) {
            // Take the first match
            const filePath = matchedFiles[0];
            const cell = range.getCell(item.row, item.col);

            // Excel URI format for local files
            const uri = "file:///" + filePath.replace(/\\/g, "/");

            // To set a hyperlink, we need it to look like a formula =HYPERLINK("uri", "friendly_name")
            // Or use the range.hyperlink property if available, but formula is often safer/easier
            cell.formulas = [[`=HYPERLINK("${uri}", "${item.term.replace(/"/g, '""')}")`]];
            linkCount++;
          }
        }

        await context.sync();

        if (linkCount > 0) {
          setStatus(`Linked ${linkCount} cells!`, "success");
        } else {
          setStatus("No matching files found.", "ready");
        }
      });
    } catch (e) {
      console.error("Reference search failed", e);
      setStatus("Error during search", "error");
    } finally {
      this.showLoading(false);
    }
  }

}

// Initialize when DOM is ready (and Office if available)
document.addEventListener("DOMContentLoaded", () => {
  try {
    window.excelAI = new ExcelAIAssistant();
  } catch (error) {
    console.error("Failed to initialize Excel AI Assistant:", error);
  }
});

// Error logging
window.addEventListener("error", (e) => console.error("Uncaught error:", e.error));
window.addEventListener("unhandledrejection", (e) =>
  console.error("Unhandled promise rejection:", e.reason)
);
