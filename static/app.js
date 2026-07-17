(() => {
  const state = {
    catalog: null,
    currentOffer: null,
    savedOfferId: null,
  };

  const form = document.getElementById("offerForm");
  const instanceOptions = document.getElementById("instanceOptions");
  const addonOptions = document.getElementById("addonOptions");
  const archiveList = document.getElementById("archiveList");
  const includedFunctions = document.getElementById("includedFunctions");

  const money = (value, currency = "EUR") =>
    new Intl.NumberFormat("de-DE", {
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
    if (res.status === 204) return null;
    const type = res.headers.get("content-type") || "";
    if (type.includes("application/json")) return res.json();
    return res;
  }

  function selectedInstanceId() {
    const checked = form.querySelector('input[name="instanceId"]:checked');
    return checked ? checked.value : state.catalog.instances[0].id;
  }

  function currentInstance() {
    return state.catalog.instances.find((i) => i.id === selectedInstanceId());
  }

  function collectPayload() {
    const data = new FormData(form);
    return {
      customer: {
        company: String(data.get("company") || "").trim(),
        contact: String(data.get("contact") || "").trim(),
        email: String(data.get("email") || "").trim(),
        phone: String(data.get("phone") || "").trim(),
        address: String(data.get("address") || "").trim(),
        projectName: String(data.get("projectName") || "").trim(),
      },
      instanceId: selectedInstanceId(),
      instanceCount: Number(data.get("instanceCount") || 1),
      selectedAddons: [...form.querySelectorAll('input[name="addon"]:checked')].map((el) => el.value),
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

  function functionDetails(ids = []) {
    const catalog = state.catalog.functionCatalog || {};
    return ids.map((id) => catalog[id]).filter(Boolean);
  }

  function renderInstances() {
    instanceOptions.innerHTML = "";
    state.catalog.instances.forEach((inst, index) => {
      const label = document.createElement("label");
      label.className = `option${index === 0 ? " selected" : ""}`;
      label.innerHTML = `
        <input type="radio" name="instanceId" value="${inst.id}" ${index === 0 ? "checked" : ""} />
        <div>
          <h3>${inst.name}</h3>
          <p>${inst.description}</p>
          <p class="muted">${inst.functionalSummary || ""}</p>
          <p class="muted">inkl. ${inst.includedOpeningClients} Opening + ${inst.includedAdminClients} Admin Client</p>
        </div>
        <div class="price">${money(inst.price)}</div>
      `;
      label.querySelector("input").addEventListener("change", () => {
        instanceOptions.querySelectorAll(".option").forEach((el) => el.classList.remove("selected"));
        label.classList.add("selected");
        renderIncluded();
        renderAddons();
        updateClientAvailability();
        recalculate();
      });
      instanceOptions.appendChild(label);
    });
    renderIncluded();
  }

  function renderIncluded() {
    const inst = currentInstance();
    const details = functionDetails(inst?.includedFunctionIds || []);
    includedFunctions.innerHTML = details
      .map(
        (f) => `
        <li>
          <strong>${f.name}</strong>
          <span>${f.description}</span>
        </li>`
      )
      .join("");
  }

  function renderAddons() {
    const instId = selectedInstanceId();
    const previous = new Set(
      [...addonOptions.querySelectorAll('input[name="addon"]:checked')].map((el) => el.value)
    );
    addonOptions.innerHTML = "";
    state.catalog.addons.forEach((addon) => {
      const available = addon.availableFor.includes(instId);
      const el = document.createElement("label");
      el.className = `feature${available ? "" : " locked"}`;
      el.innerHTML = `
        <input type="checkbox" name="addon" value="${addon.id}"
          ${available && previous.has(addon.id) ? "checked" : ""}
          ${available ? "" : "disabled"} />
        <div>
          <strong>${addon.name}</strong>
          <span>${addon.functionalDescription || addon.description || ""}</span>
          <span>${available ? money(addon.price) : "Nur für Advanced Instance"}</span>
        </div>
      `;
      if (available) {
        el.querySelector("input").addEventListener("change", () => recalculate());
      }
      addonOptions.appendChild(el);
    });
  }

  function updateClientAvailability() {
    const advanced = selectedInstanceId() === "advanced";
    form.mobileTerminalClients.disabled = !advanced;
    if (!advanced) form.mobileTerminalClients.value = 0;
    const hints = document.getElementById("clientHints");
    if (!hints) return;
    hints.innerHTML = (state.catalog.clientLicenses || [])
      .map((c) => {
        const locked = c.availableFor && !c.availableFor.includes(selectedInstanceId());
        return `
          <div class="hint-card${locked ? " locked" : ""}">
            <strong>${c.name}</strong>
            <span>${c.functionalDescription || c.description}</span>
            <span>${money(c.price)}${locked ? " · nur Advanced" : ""}</span>
          </div>`;
      })
      .join("");
  }

  function renderSllHint() {
    document.getElementById("sllHint").textContent = state.catalog.sllDefinition || "";
  }

  function renderPreview(offer) {
    state.currentOffer = offer;
    const currency = offer.totals.currency || "EUR";
    document.getElementById("previewOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("previewGross").textContent = money(offer.totals.net, currency);
    document.getElementById("sllBadge").textContent =
      `SLL: ${offer.totals.sllCount} · Rabatt ${offer.totals.discountPercent}%`;

    const customer = offer.customer;
    const version = offer.product.version || "";
    document.getElementById("previewMeta").innerHTML = `
      <div><strong>${customer.company || "—"}</strong>${customer.projectName ? ` · ${customer.projectName}` : ""}</div>
      <div>${offer.configuration.instanceCount}× ${offer.configuration.instanceName} · LS ${version}</div>
      <div>${offer.meta.priceBasis || "IC Prices"}</div>
      <div>Gültig bis ${offer.meta.validUntil}${offer.configuration.preparedBy ? ` · ${offer.configuration.preparedBy}` : ""}</div>
    `;

    document.getElementById("previewLines").innerHTML = offer.lines
      .map(
        (line) => `
        <tr>
          <td>
            ${line.name}
            <div class="muted">${line.description || ""}</div>
            <div class="muted">${money(line.unitPrice, currency)} · SLL ${line.sllUnits || 0}</div>
          </td>
          <td>${line.qty}</td>
          <td>${money(line.total, currency)}</td>
        </tr>`
      )
      .join("");

    document.getElementById("previewTotals").innerHTML = `
      <div class="row"><span>Zwischensumme</span><span>${money(offer.totals.subtotal, currency)}</span></div>
      <div class="row"><span>Mengenrabatt (${offer.totals.discountPercent}%)</span><span>− ${money(offer.totals.discountAmount, currency)}</span></div>
      <div class="row strong"><span>IC Total</span><span>${money(offer.totals.net, currency)}</span></div>
    `;

    const scope = offer.scopeOfSupply || [];
    document.getElementById("previewScope").innerHTML = scope.map((line) => `<li>${line}</li>`).join("");
    document.getElementById("previewDisclaimer").textContent = offer.product.disclaimer;

    document.getElementById("btnExportExcel").disabled = !state.savedOfferId;
    document.getElementById("btnPrint").disabled = !offer;
  }

  async function recalculate() {
    const payload = collectPayload();
    if (!payload.customer.company) return;
    try {
      const offer = await api("/api/offers/calculate", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      state.savedOfferId = null;
      renderPreview(offer);
    } catch (err) {
      console.error(err);
      alert(err.message);
    }
  }

  async function saveOffer(event) {
    event.preventDefault();
    const payload = collectPayload();
    if (!payload.customer.company) {
      alert("Bitte Firma angeben.");
      return;
    }
    try {
      const offer = await api("/api/offers", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      state.savedOfferId = offer.id;
      renderPreview(offer);
      await loadArchive();
      alert(`Kalkulation ${offer.meta.offerNumber} gespeichert.`);
    } catch (err) {
      alert(err.message);
    }
  }

  async function loadArchive() {
    const data = await api("/api/offers");
    const offers = data.offers || [];
    if (!offers.length) {
      archiveList.innerHTML = '<p class="empty-state">Noch keine Kalkulationen gespeichert.</p>';
      return;
    }
    archiveList.innerHTML = offers
      .map(
        (o) => `
        <article class="archive-item" data-id="${o.id}">
          <div>
            <h3>${o.offerNumber}</h3>
            <p>${o.company || "—"}${o.projectName ? ` · ${o.projectName}` : ""} · ${o.instanceName || ""}</p>
            <p>SLL ${o.sllCount ?? "–"} · ${money(o.gross, o.currency)} · ${o.createdAt || ""}</p>
          </div>
          <div class="archive-actions">
            <button type="button" class="btn ghost" data-action="open">Öffnen</button>
            <button type="button" class="btn ghost" data-action="excel">Excel</button>
            <button type="button" class="btn danger" data-action="delete">Löschen</button>
          </div>
        </article>`
      )
      .join("");
  }

  async function openOffer(id) {
    const offer = await api(`/api/offers/${encodeURIComponent(id)}`);
    state.savedOfferId = offer.id;
    const cfg = offer.configuration;
    form.company.value = offer.customer.company || "";
    form.contact.value = offer.customer.contact || "";
    form.email.value = offer.customer.email || "";
    form.phone.value = offer.customer.phone || "";
    form.address.value = offer.customer.address || "";
    form.projectName.value = offer.customer.projectName || "";
    form.preparedBy.value = cfg.preparedBy || "";
    form.notes.value = cfg.notes || "";
    form.instanceCount.value = cfg.instanceCount || 1;
    form.extraOpeningClients.value = cfg.extraOpeningClients || 0;
    form.extraAdminClients.value = cfg.extraAdminClients || 0;
    form.mobileTerminalClients.value = cfg.mobileTerminalClients || 0;
    form.thirdPartyVlmTypes.value = cfg.thirdPartyVlmTypes || 0;
    form.testInstances.value = cfg.testInstances || 0;
    form.upgradeYears.value = cfg.upgradeYears || 0;

    const input = form.querySelector(`input[name="instanceId"][value="${cfg.instanceId}"]`);
    if (input) {
      input.checked = true;
      instanceOptions.querySelectorAll(".option").forEach((el) => el.classList.remove("selected"));
      input.closest(".option")?.classList.add("selected");
    }
    renderIncluded();
    renderAddons();
    updateClientAvailability();
    const selected = new Set(cfg.selectedAddons || []);
    addonOptions.querySelectorAll('input[name="addon"]').forEach((el) => {
      if (!el.disabled) el.checked = selected.has(el.value);
    });
    renderPreview(offer);
    switchView("configurator");
  }

  function switchView(name) {
    document.querySelectorAll(".view").forEach((el) => el.classList.remove("active"));
    document.querySelectorAll(".nav-link").forEach((el) => {
      el.classList.toggle("active", el.dataset.view === name);
    });
    document.getElementById(`view-${name}`).classList.add("active");
    if (name === "archive") loadArchive();
  }

  function bindEvents() {
    document.querySelectorAll(".nav-link").forEach((btn) => {
      btn.addEventListener("click", () => switchView(btn.dataset.view));
    });
    document.getElementById("btnScrollConfig").addEventListener("click", () => {
      document.getElementById("configPanel").scrollIntoView({ behavior: "smooth" });
    });
    document.getElementById("btnRecalc").addEventListener("click", recalculate);
    document.getElementById("btnPrintPreview").addEventListener("click", () => window.print());
    document.getElementById("btnPrint").addEventListener("click", () => window.print());
    form.addEventListener("submit", saveOffer);

    [
      "instanceCount",
      "extraOpeningClients",
      "extraAdminClients",
      "mobileTerminalClients",
      "thirdPartyVlmTypes",
      "testInstances",
      "upgradeYears",
    ].forEach((name) => form[name].addEventListener("change", recalculate));
    form.company.addEventListener("change", recalculate);

    document.getElementById("btnExportExcel").addEventListener("click", () => {
      if (!state.savedOfferId) return;
      window.location.href = `/api/offers/${encodeURIComponent(state.savedOfferId)}/excel`;
    });

    archiveList.addEventListener("click", async (event) => {
      const btn = event.target.closest("button[data-action]");
      if (!btn) return;
      const item = btn.closest(".archive-item");
      const id = item?.dataset.id;
      if (!id) return;
      try {
        if (btn.dataset.action === "open") await openOffer(id);
        if (btn.dataset.action === "excel") {
          window.location.href = `/api/offers/${encodeURIComponent(id)}/excel`;
        }
        if (btn.dataset.action === "delete") {
          if (!confirm("Kalkulation wirklich löschen?")) return;
          await api(`/api/offers/${encodeURIComponent(id)}`, { method: "DELETE" });
          await loadArchive();
        }
      } catch (err) {
        alert(err.message);
      }
    });
  }

  async function init() {
    state.catalog = await api("/api/catalog");
    renderInstances();
    renderAddons();
    updateClientAvailability();
    renderSllHint();
    bindEvents();
    document.getElementById("previewDisclaimer").textContent = state.catalog.product.disclaimer;
    document.getElementById("previewOfferNo").textContent = "Entwurf";
  }

  init().catch((err) => {
    console.error(err);
    alert("Katalog konnte nicht geladen werden.");
  });
})();
