(() => {
  const state = {
    catalog: null,
    currentOffer: null,
    savedOfferId: null,
  };

  const form = document.getElementById("offerForm");
  const packageOptions = document.getElementById("packageOptions");
  const featureOptions = document.getElementById("featureOptions");
  const discountSelect = document.getElementById("discountSelect");
  const archiveList = document.getElementById("archiveList");

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
      throw new Error(detail);
    }
    if (res.status === 204) return null;
    const type = res.headers.get("content-type") || "";
    if (type.includes("application/json")) return res.json();
    return res;
  }

  function selectedPackageId() {
    const checked = form.querySelector('input[name="packageId"]:checked');
    return checked ? checked.value : state.catalog.basePackages[0].id;
  }

  function packageFeatureSet(packageId) {
    const pkg = state.catalog.basePackages.find((p) => p.id === packageId);
    return new Set(pkg?.features || []);
  }

  function collectPayload() {
    const data = new FormData(form);
    const packageId = selectedPackageId();
    const included = packageFeatureSet(packageId);
    const selectedFeatures = [...form.querySelectorAll('input[name="feature"]:checked')]
      .map((el) => el.value)
      .filter((id) => !included.has(id));

    return {
      customer: {
        company: String(data.get("company") || "").trim(),
        contact: String(data.get("contact") || "").trim(),
        email: String(data.get("email") || "").trim(),
        phone: String(data.get("phone") || "").trim(),
        address: String(data.get("address") || "").trim(),
        projectName: String(data.get("projectName") || "").trim(),
      },
      packageId,
      liftCount: Number(data.get("liftCount") || 1),
      selectedFeatures,
      trainingDays: Number(data.get("trainingDays") || 0),
      includeImplementation: form.includeImplementation.checked,
      includeMaintenance: form.includeMaintenance.checked,
      discountId: String(data.get("discountId") || "none"),
      notes: String(data.get("notes") || "").trim(),
      preparedBy: String(data.get("preparedBy") || "").trim(),
    };
  }

  function renderPackages() {
    packageOptions.innerHTML = "";
    state.catalog.basePackages.forEach((pkg, index) => {
      const label = document.createElement("label");
      label.className = `option${index === 1 ? " selected" : ""}`;
      label.innerHTML = `
        <input type="radio" name="packageId" value="${pkg.id}" ${index === 1 ? "checked" : ""} />
        <div>
          <h3>${pkg.name}</h3>
          <p>${pkg.description}</p>
          <p class="muted">${pkg.includedLiftLicenses} Lift-Lizenz(en) inklusive</p>
        </div>
        <div class="price">${money(pkg.basePrice)}</div>
      `;
      const input = label.querySelector("input");
      input.addEventListener("change", () => {
        packageOptions.querySelectorAll(".option").forEach((el) => el.classList.remove("selected"));
        label.classList.add("selected");
        renderFeatures();
        recalculate();
      });
      packageOptions.appendChild(label);
    });
  }

  function renderFeatures() {
    const included = packageFeatureSet(selectedPackageId());
    const previous = new Set(
      [...featureOptions.querySelectorAll('input[name="feature"]:checked')].map((el) => el.value)
    );

    featureOptions.innerHTML = "";
    state.catalog.features.forEach((feat) => {
      const locked = included.has(feat.id);
      const checked = locked || previous.has(feat.id);
      const el = document.createElement("label");
      el.className = `feature${locked ? " locked" : ""}`;
      el.innerHTML = `
        <input type="checkbox" name="feature" value="${feat.id}" ${checked ? "checked" : ""} ${locked ? "disabled" : ""} />
        <div>
          <strong>${feat.name}</strong>
          <span>${feat.description}</span>
          <span>${locked ? "Im Paket enthalten" : money(feat.price)}</span>
        </div>
      `;
      if (!locked) {
        el.querySelector("input").addEventListener("change", () => recalculate());
      }
      featureOptions.appendChild(el);
    });
  }

  function renderDiscounts() {
    discountSelect.innerHTML = state.catalog.discounts
      .map((d) => `<option value="${d.id}">${d.name}${d.percent ? ` (−${d.percent}%)` : ""}</option>`)
      .join("");
  }

  function renderPreview(offer) {
    state.currentOffer = offer;
    const currency = offer.totals.currency || "EUR";
    document.getElementById("previewOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("previewGross").textContent = money(offer.totals.gross, currency);

    const customer = offer.customer;
    const version = offer.product.version || offer.meta.productVersion || "";
    document.getElementById("previewMeta").innerHTML = `
      <div><strong>${customer.company || "—"}</strong>${customer.projectName ? ` · ${customer.projectName}` : ""}</div>
      <div>WAMAS Lift &amp; Store ${version} · Paket ${offer.configuration.packageName} · ${offer.configuration.liftCount} Lift(e)</div>
      <div>Gültig bis ${offer.meta.validUntil}${offer.configuration.preparedBy ? ` · ${offer.configuration.preparedBy}` : ""}</div>
    `;

    document.getElementById("previewLines").innerHTML = offer.lines
      .map(
        (line) => `
        <tr>
          <td>
            ${line.name}
            ${line.includedInPackage ? '<div class="muted">Im Paket</div>' : ""}
          </td>
          <td>${line.qty}</td>
          <td>${money(line.total, currency)}</td>
        </tr>`
      )
      .join("");

    document.getElementById("previewTotals").innerHTML = `
      <div class="row"><span>Zwischensumme</span><span>${money(offer.totals.subtotal, currency)}</span></div>
      <div class="row"><span>Rabatt (${offer.totals.discountPercent}%)</span><span>− ${money(offer.totals.discountAmount, currency)}</span></div>
      <div class="row"><span>Netto</span><span>${money(offer.totals.net, currency)}</span></div>
      <div class="row"><span>MwSt. (${Math.round(offer.totals.vatRate * 100)}%)</span><span>${money(offer.totals.vat, currency)}</span></div>
      <div class="row strong"><span>Brutto</span><span>${money(offer.totals.gross, currency)}</span></div>
    `;

    const scope = offer.scopeOfSupply || [];
    document.getElementById("previewScope").innerHTML = scope.length
      ? scope.map((line) => `<li>${line}</li>`).join("")
      : "<li>Kein zusätzlicher Leistungstext.</li>";

    document.getElementById("previewDisclaimer").textContent = offer.product.disclaimer;

    const canExport = Boolean(state.savedOfferId);
    document.getElementById("btnExportExcel").disabled = !canExport;
    document.getElementById("btnPrint").disabled = !offer;
  }

  async function recalculate() {
    const payload = collectPayload();
    if (!payload.customer.company) {
      return;
    }
    try {
      const offer = await api("/api/offers/calculate", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      // Keep provisional number until saved
      if (!state.savedOfferId) {
        renderPreview(offer);
      } else {
        // After edits, treat as draft again until re-saved
        state.savedOfferId = null;
        renderPreview(offer);
      }
    } catch (err) {
      console.error(err);
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
      alert(`Angebot ${offer.meta.offerNumber} gespeichert.`);
    } catch (err) {
      alert(err.message);
    }
  }

  async function loadArchive() {
    const data = await api("/api/offers");
    const offers = data.offers || [];
    if (!offers.length) {
      archiveList.innerHTML = '<p class="empty-state">Noch keine Angebote gespeichert.</p>';
      return;
    }
    archiveList.innerHTML = offers
      .map(
        (o) => `
        <article class="archive-item" data-id="${o.id}">
          <div>
            <h3>${o.offerNumber}</h3>
            <p>${o.company || "—"}${o.projectName ? ` · ${o.projectName}` : ""} · ${o.packageName || ""}</p>
            <p>${o.createdAt || ""} · ${money(o.gross, o.currency)}</p>
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
    form.company.value = offer.customer.company || "";
    form.contact.value = offer.customer.contact || "";
    form.email.value = offer.customer.email || "";
    form.phone.value = offer.customer.phone || "";
    form.address.value = offer.customer.address || "";
    form.projectName.value = offer.customer.projectName || "";
    form.preparedBy.value = offer.configuration.preparedBy || "";
    form.notes.value = offer.configuration.notes || "";
    form.liftCount.value = offer.configuration.liftCount || 1;
    form.trainingDays.value = offer.configuration.trainingDays || 0;
    form.includeImplementation.checked = !!offer.configuration.includeImplementation;
    form.includeMaintenance.checked = !!offer.configuration.includeMaintenance;
    form.discountId.value = offer.configuration.discountId || "none";

    const pkgInput = form.querySelector(`input[name="packageId"][value="${offer.configuration.packageId}"]`);
    if (pkgInput) {
      pkgInput.checked = true;
      packageOptions.querySelectorAll(".option").forEach((el) => el.classList.remove("selected"));
      pkgInput.closest(".option")?.classList.add("selected");
    }

    const selected = new Set(offer.configuration.selectedFeatures || []);
    renderFeatures();
    featureOptions.querySelectorAll('input[name="feature"]').forEach((input) => {
      if (!input.disabled) input.checked = selected.has(input.value);
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

    ["liftCount", "trainingDays", "discountId", "includeImplementation", "includeMaintenance"].forEach((name) => {
      form[name].addEventListener("change", recalculate);
    });
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
      const action = btn.dataset.action;
      try {
        if (action === "open") await openOffer(id);
        if (action === "excel") window.location.href = `/api/offers/${encodeURIComponent(id)}/excel`;
        if (action === "delete") {
          if (!confirm("Angebot wirklich löschen?")) return;
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
    renderPackages();
    renderFeatures();
    renderDiscounts();
    bindEvents();
    document.getElementById("previewDisclaimer").textContent = state.catalog.product.disclaimer;
    document.getElementById("previewOfferNo").textContent = "Entwurf";
    document.getElementById("previewGross").textContent = "–";
  }

  init().catch((err) => {
    console.error(err);
    alert("Katalog konnte nicht geladen werden. Bitte Server prüfen.");
  });
})();
