(() => {
  const state = {
    licenseCatalog: null,
    itCatalog: null,
    licenseOffer: null,
    itOffer: null,
    savedLicenseId: null,
    savedItId: null,
  };

  const licenseForm = document.getElementById("licenseForm");
  const itForm = document.getElementById("itForm");
  const archiveList = document.getElementById("archiveList");

  const money = (value, currency = "EUR") =>
    new Intl.NumberFormat("de-CH", {
      style: "currency",
      currency,
      maximumFractionDigits: 2,
    }).format(Number(value || 0));

  async function api(path, options = {}) {
    const res = await fetch(path, {
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
      ...options,
    });
    if (!res.ok) {
      let detail = "Anfrage fehlgeschlagen";
      try {
        const data = await res.json();
        detail = data.detail || detail;
      } catch (_) {}
      throw new Error(typeof detail === "string" ? detail : JSON.stringify(detail));
    }
    const type = res.headers.get("content-type") || "";
    if (type.includes("application/json")) return res.json();
    return res;
  }

  function switchView(name) {
    document.querySelectorAll(".view").forEach((el) => el.classList.remove("active"));
    document.querySelectorAll(".nav-link").forEach((el) => {
      el.classList.toggle("active", el.dataset.view === name);
    });
    document.getElementById(`view-${name}`).classList.add("active");
    if (name === "archive") loadArchive();
  }

  // -------------------- LICENSE --------------------
  function selectedInstanceId() {
    return licenseForm.querySelector('input[name="instanceId"]:checked')?.value
      || state.licenseCatalog.instances[0].id;
  }

  function currentInstance() {
    return state.licenseCatalog.instances.find((i) => i.id === selectedInstanceId());
  }

  function collectLicensePayload() {
    const data = new FormData(licenseForm);
    return {
      customer: {
        company: String(data.get("company") || "").trim(),
        contact: String(data.get("contact") || "").trim(),
        email: String(data.get("email") || "").trim(),
        phone: "",
        address: String(data.get("address") || "").trim(),
        projectName: String(data.get("projectName") || "").trim(),
      },
      instanceId: selectedInstanceId(),
      instanceCount: Number(data.get("instanceCount") || 1),
      selectedAddons: [...licenseForm.querySelectorAll('input[name="addon"]:checked')].map((el) => el.value),
      extraOpeningClients: Number(data.get("extraOpeningClients") || 0),
      extraAdminClients: Number(data.get("extraAdminClients") || 0),
      mobileTerminalClients: Number(data.get("mobileTerminalClients") || 0),
      thirdPartyVlmTypes: Number(data.get("thirdPartyVlmTypes") || 0),
      testInstances: Number(data.get("testInstances") || 0),
      upgradeYears: Number(data.get("upgradeYears") || 0),
      notes: String(data.get("notes") || "").trim(),
      preparedBy: String(data.get("preparedBy") || "").trim(),
    };
  }

  function renderInstances() {
    const root = document.getElementById("instanceOptions");
    root.innerHTML = "";
    state.licenseCatalog.instances.forEach((inst, index) => {
      const label = document.createElement("label");
      label.className = `option${index === 0 ? " selected" : ""}`;
      label.innerHTML = `
        <input type="radio" name="instanceId" value="${inst.id}" ${index === 0 ? "checked" : ""} />
        <div>
          <h3>${inst.name}</h3>
          <p>${inst.description}</p>
          <p class="muted">${inst.functionalSummary || ""}</p>
        </div>
        <div class="price">${money(inst.price, "EUR")}</div>`;
      label.querySelector("input").addEventListener("change", () => {
        root.querySelectorAll(".option").forEach((el) => el.classList.remove("selected"));
        label.classList.add("selected");
        renderIncluded();
        renderAddons();
        updateClientHints();
        recalcLicense();
      });
      root.appendChild(label);
    });
    renderIncluded();
  }

  function renderIncluded() {
    const catalog = state.licenseCatalog.functionCatalog || {};
    const ids = currentInstance()?.includedFunctionIds || [];
    document.getElementById("includedFunctions").innerHTML = ids
      .map((id) => catalog[id])
      .filter(Boolean)
      .map((f) => `<li><strong>${f.name}</strong><span>${f.description}</span></li>`)
      .join("");
  }

  function renderAddons() {
    const root = document.getElementById("addonOptions");
    const instId = selectedInstanceId();
    const prev = new Set([...root.querySelectorAll('input[name="addon"]:checked')].map((el) => el.value));
    root.innerHTML = "";
    state.licenseCatalog.addons.forEach((addon) => {
      const available = addon.availableFor.includes(instId);
      const el = document.createElement("label");
      el.className = `feature${available ? "" : " locked"}`;
      el.innerHTML = `
        <input type="checkbox" name="addon" value="${addon.id}"
          ${available && prev.has(addon.id) ? "checked" : ""} ${available ? "" : "disabled"} />
        <div>
          <strong>${addon.name}</strong>
          <span>${addon.functionalDescription || addon.description || ""}</span>
          <span>${available ? money(addon.price, "EUR") : "Nur Advanced"}</span>
        </div>`;
      if (available) el.querySelector("input").addEventListener("change", () => recalcLicense());
      root.appendChild(el);
    });
  }

  function updateClientHints() {
    const advanced = selectedInstanceId() === "advanced";
    licenseForm.mobileTerminalClients.disabled = !advanced;
    if (!advanced) licenseForm.mobileTerminalClients.value = 0;
    document.getElementById("clientHints").innerHTML = (state.licenseCatalog.clientLicenses || [])
      .map((c) => {
        const locked = !c.availableFor.includes(selectedInstanceId());
        return `<div class="hint-card${locked ? " locked" : ""}">
          <strong>${c.name}</strong>
          <span>${c.functionalDescription || c.description}</span>
          <span>${money(c.price, "EUR")}${locked ? " · nur Advanced" : ""}</span>
        </div>`;
      })
      .join("");
  }

  function renderLicensePreview(offer) {
    state.licenseOffer = offer;
    const c = offer.totals.currency || "EUR";
    document.getElementById("licenseOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("licenseGross").textContent = money(offer.totals.net, c);
    document.getElementById("sllBadge").textContent =
      `SLL: ${offer.totals.sllCount} · Rabatt ${offer.totals.discountPercent}%`;
    document.getElementById("licenseMeta").innerHTML = `
      <div><strong>${offer.customer.company}</strong>${offer.customer.projectName ? ` · ${offer.customer.projectName}` : ""}</div>
      <div>${offer.configuration.instanceCount}× ${offer.configuration.instanceName}</div>
      <div>${offer.meta.priceBasis || ""}</div>`;
    document.getElementById("licenseLines").innerHTML = offer.lines.map((line) => `
      <tr>
        <td>${line.name}<div class="muted">${line.description || ""}</div></td>
        <td>${line.qty}</td>
        <td>${money(line.total, c)}</td>
      </tr>`).join("");
    document.getElementById("licenseTotals").innerHTML = `
      <div class="row"><span>Zwischensumme</span><span>${money(offer.totals.subtotal, c)}</span></div>
      <div class="row"><span>Mengenrabatt</span><span>− ${money(offer.totals.discountAmount, c)}</span></div>
      <div class="row strong"><span>IC Total</span><span>${money(offer.totals.net, c)}</span></div>`;
    document.getElementById("licenseScope").innerHTML =
      (offer.scopeOfSupply || []).map((s) => `<li>${s}</li>`).join("");
    document.getElementById("licenseDisclaimer").textContent = offer.product.disclaimer;
    document.getElementById("btnLicenseExcel").disabled = !state.savedLicenseId;
    document.getElementById("btnLicensePrint").disabled = !offer;
  }

  async function recalcLicense() {
    const payload = collectLicensePayload();
    if (!payload.customer.company) {
      payload.customer.company = "Entwurf";
    }
    try {
      const offer = await api("/api/offers/calculate", { method: "POST", body: JSON.stringify(payload) });
      state.savedLicenseId = null;
      renderLicensePreview(offer);
    } catch (err) {
      alert(err.message);
    }
  }

  // -------------------- IT --------------------
  function renderItOptions() {
    const root = document.getElementById("itOptions");
    root.innerHTML = "";
    state.itCatalog.options.forEach((opt) => {
      const el = document.createElement("label");
      el.className = "feature";
      el.innerHTML = `
        <input type="checkbox" name="itOption" value="${opt.id}" />
        <div>
          <strong>${opt.name}</strong>
          <span>${opt.description}</span>
        </div>`;
      el.querySelector("input").addEventListener("change", () => recalcIt());
      root.appendChild(el);
    });
  }

  function renderItExtensions() {
    const root = document.getElementById("itExtensions");
    root.innerHTML = "";
    for (let i = 1; i <= 5; i += 1) {
      root.innerHTML += `
        <label>Erweiterung ${i} Beschreibung
          <input name="extDesc${i}" placeholder="Beschreibung …" />
        </label>
        <label>Erweiterung ${i} Stunden
          <input name="extHours${i}" type="number" min="0" max="1000" step="0.5" value="0" />
        </label>`;
    }
    root.querySelectorAll("input").forEach((el) => el.addEventListener("change", () => recalcIt()));
  }

  function collectItPayload() {
    const data = new FormData(itForm);
    const options = {};
    state.itCatalog.options.forEach((opt) => { options[opt.id] = false; });
    itForm.querySelectorAll('input[name="itOption"]:checked').forEach((el) => {
      options[el.value] = true;
    });
    const customExtensions = [];
    for (let i = 1; i <= 5; i += 1) {
      customExtensions.push({
        description: String(data.get(`extDesc${i}`) || "").trim(),
        hours: Number(data.get(`extHours${i}`) || 0),
      });
    }
    return {
      customer: {
        company: String(data.get("company") || "").trim(),
        contact: "",
        email: "",
        phone: "",
        address: "",
        projectName: String(data.get("projectName") || "").trim(),
      },
      realizationPeriod: String(data.get("realizationPeriod") || "").trim(),
      deviceCount: Number(data.get("deviceCount") || 1),
      zoneCount: Number(data.get("zoneCount") || 1),
      openingCount: Number(data.get("openingCount") || 1),
      options,
      customExtensions,
      trips: Number(data.get("trips") || 0),
      travelHoursPerTrip: Number(data.get("travelHoursPerTrip") || 0),
      kmPerTrip: Number(data.get("kmPerTrip") || 0),
      overnightCount: Number(data.get("overnightCount") || 0),
      mealCount: Number(data.get("mealCount") || 0),
      notes: String(data.get("notes") || "").trim(),
      preparedBy: String(data.get("preparedBy") || "").trim(),
    };
  }

  function renderItPreview(offer) {
    state.itOffer = offer;
    const c = offer.totals.currency || "CHF";
    document.getElementById("itOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("itGross").textContent = money(offer.totals.totalAmount, c);
    document.getElementById("itMeta").innerHTML = `
      <div><strong>${offer.customer.company}</strong>${offer.customer.projectName ? ` · ${offer.customer.projectName}` : ""}</div>
      <div>${offer.configuration.deviceCount} Geräte · ${offer.configuration.zoneCount} Zone(n) · ${offer.configuration.openingCount} Öffnung(en)</div>
      <div>Stundensatz ${money(offer.configuration.hourlyRate, c)} · ${offer.configuration.realizationPeriod || "–"}</div>`;

    const visibleLines = offer.lines.filter((l) => {
      if (l.category === "option") return Boolean(l.selected) && (Number(l.hours) > 0 || Number(l.amount) > 0 || l.note);
      if (l.category === "travel") return Number(l.amount) > 0 || Number(l.hours) > 0;
      if (l.category === "custom") return Number(l.hours) > 0 || Number(l.amount) > 0;
      return Number(l.hours) > 0 || Number(l.amount) > 0 || ["IT-DEVICES", "IT-ZONES", "IT-OPENINGS"].includes(l.sku);
    });

    document.getElementById("itLines").innerHTML = visibleLines.map((line) => `
        <tr>
          <td>${line.name}<div class="muted">${line.description || ""}${line.note ? ` · ${line.note}` : ""}</div></td>
          <td>${line.hours || 0}</td>
          <td>${money(line.amount, c)}</td>
        </tr>`).join("");
    document.getElementById("itTotals").innerHTML = `
      <div class="row"><span>IT-Aufwand</span><span>${offer.totals.workHours} h · ${money(offer.totals.workAmount, c)}</span></div>
      <div class="row"><span>Reisekosten</span><span>${money(offer.totals.travelAmount, c)}</span></div>
      <div class="row strong"><span>Total exkl. MwSt</span><span>${money(offer.totals.totalAmount, c)}</span></div>`;
    document.getElementById("itScope").innerHTML = (offer.offerSections || [])
      .filter((s) => Number(s.amount) > 0 || (s.bullets || []).length)
      .map((s) => `<li><strong>${s.title}</strong> · ${money(s.amount, c)}<div class="muted">${(s.bullets || []).join(" · ")}</div></li>`)
      .join("");
    document.getElementById("itDisclaimer").textContent = offer.product.disclaimer;
    document.getElementById("btnItExcel").disabled = !state.savedItId;
    document.getElementById("btnItPrint").disabled = !offer;
  }

  async function recalcIt() {
    const payload = collectItPayload();
    // Live-Vorschau auch ohne Firmenname; Speichern verlangt weiterhin Firma.
    if (!payload.customer.company) {
      payload.customer.company = "Entwurf";
    }
    try {
      const offer = await api("/api/it/calculate", { method: "POST", body: JSON.stringify(payload) });
      state.savedItId = null;
      renderItPreview(offer);
    } catch (err) {
      alert(err.message);
    }
  }

  // -------------------- ARCHIVE --------------------
  async function loadArchive() {
    const data = await api("/api/offers");
    const offers = data.offers || [];
    if (!offers.length) {
      archiveList.innerHTML = '<p class="empty-state">Noch keine Kalkulationen gespeichert.</p>';
      return;
    }
    archiveList.innerHTML = offers.map((o) => `
      <article class="archive-item" data-id="${o.id}">
        <div>
          <h3>${o.offerNumber} <span class="muted">(${o.kind === "it" ? "IT" : "Lizenz"})</span></h3>
          <p>${o.company || "—"} · ${o.summary || ""}</p>
          <p>${money(o.amount, o.currency)} · ${o.createdAt || ""}</p>
        </div>
        <div class="archive-actions">
          <button type="button" class="btn" data-action="excel">Excel</button>
          <button type="button" class="btn danger" data-action="delete">Löschen</button>
        </div>
      </article>`).join("");
  }

  function bindEvents() {
    document.querySelectorAll(".nav-link").forEach((btn) => {
      btn.addEventListener("click", () => switchView(btn.dataset.view));
    });
    document.getElementById("btnScrollLicense").addEventListener("click", () => {
      document.getElementById("licensePanel").scrollIntoView({ behavior: "smooth" });
    });
    document.getElementById("btnScrollIt").addEventListener("click", () => {
      document.getElementById("itPanel").scrollIntoView({ behavior: "smooth" });
    });

    document.getElementById("btnLicenseRecalc").addEventListener("click", recalcLicense);
    document.getElementById("btnLicensePrint").addEventListener("click", () => window.print());
    document.getElementById("btnLicenseExcel").addEventListener("click", () => {
      if (state.savedLicenseId) window.location.href = `/api/offers/${encodeURIComponent(state.savedLicenseId)}/excel`;
    });
    licenseForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const payload = collectLicensePayload();
      if (!payload.customer.company) {
        alert("Bitte Firma angeben.");
        return;
      }
      try {
        const offer = await api("/api/offers", { method: "POST", body: JSON.stringify(payload) });
        state.savedLicenseId = offer.id;
        renderLicensePreview(offer);
        alert(`Gespeichert: ${offer.meta.offerNumber}`);
      } catch (err) { alert(err.message); }
    });
    licenseForm.addEventListener("input", () => recalcLicense());
    licenseForm.addEventListener("change", () => recalcLicense());

    document.getElementById("btnItRecalc").addEventListener("click", recalcIt);
    document.getElementById("btnItPrint").addEventListener("click", () => window.print());
    document.getElementById("btnItExcel").addEventListener("click", () => {
      if (state.savedItId) window.location.href = `/api/offers/${encodeURIComponent(state.savedItId)}/excel`;
    });
    itForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const payload = collectItPayload();
      if (!payload.customer.company) {
        alert("Bitte Firma angeben.");
        return;
      }
      try {
        const offer = await api("/api/it/offers", { method: "POST", body: JSON.stringify(payload) });
        state.savedItId = offer.id;
        renderItPreview(offer);
        alert(`Gespeichert: ${offer.meta.offerNumber}`);
      } catch (err) { alert(err.message); }
    });

    // Live-Update bei jeder Eingabe/Änderung (nicht erst beim Verlassen des Felds)
    itForm.addEventListener("input", (event) => {
      const t = event.target;
      if (!t || !t.name) return;
      if (t.name.startsWith("ext") || [
        "deviceCount", "zoneCount", "openingCount", "trips", "travelHoursPerTrip",
        "kmPerTrip", "overnightCount", "mealCount", "company", "projectName",
        "realizationPeriod", "preparedBy", "notes",
      ].includes(t.name)) {
        recalcIt();
      }
    });
    itForm.addEventListener("change", (event) => {
      const t = event.target;
      if (!t) return;
      if (t.name === "itOption" || t.type === "checkbox" || t.tagName === "SELECT") {
        recalcIt();
      }
    });

    archiveList.addEventListener("click", async (event) => {
      const btn = event.target.closest("button[data-action]");
      if (!btn) return;
      const id = btn.closest(".archive-item")?.dataset.id;
      if (!id) return;
      if (btn.dataset.action === "excel") {
        window.location.href = `/api/offers/${encodeURIComponent(id)}/excel`;
      }
      if (btn.dataset.action === "delete") {
        if (!confirm("Eintrag löschen?")) return;
        await api(`/api/offers/${encodeURIComponent(id)}`, { method: "DELETE" });
        await loadArchive();
      }
    });
  }

  async function init() {
    state.licenseCatalog = await api("/api/catalog");
    state.itCatalog = await api("/api/it/catalog");
    renderInstances();
    renderAddons();
    updateClientHints();
    document.getElementById("sllHint").textContent = state.licenseCatalog.sllDefinition || "";
    document.getElementById("licenseDisclaimer").textContent = state.licenseCatalog.product.disclaimer;
    renderItOptions();
    renderItExtensions();
    const r = state.itCatalog.rates;
    document.getElementById("itRatesHint").textContent =
      `Stammdaten: Stundensatz ${money(r.hourlyRate, "CHF")} · km ${money(r.kmRate, "CHF")} · Verpflegung ${money(r.mealRate, "CHF")} · Übernachtung ${money(r.overnightRate, "CHF")}`;
    document.getElementById("itDisclaimer").textContent = state.itCatalog.meta.disclaimer;
    bindEvents();
    // Initiale Live-Vorschau (auch ohne Firmenname)
    recalcLicense();
    recalcIt();
  }

  init().catch((err) => {
    console.error(err);
    alert("Kataloge konnten nicht geladen werden.");
  });
})();
