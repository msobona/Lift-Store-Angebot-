(() => {
  const state = {
    licenseCatalog: null,
    itCatalog: null,
    licenseOffer: null,
    itOffer: null,
    offerDocument: null,
    savedOfferId: null,
    savedLicenseId: null,
    savedItId: null,
    editingFromOfferId: null,
  };

  const licenseForm = document.getElementById("licenseForm");
  const itForm = document.getElementById("itForm");
  const archiveList = document.getElementById("archiveList");

  // Lizenz-Add-ons → IT-Optionen (Excel-Kalkulation)
  const LICENSE_TO_IT_OPTIONS = {
    external_storage: "externalStorage",
    rfid_login: "rfid",
    printing_support: "pickLabel",
    advanced_security: "advancedSecurity",
  };

  const money = (value, currency = "CHF") =>
    new Intl.NumberFormat("de-CH", {
      style: "currency",
      currency,
      maximumFractionDigits: 2,
    }).format(Number(value || 0));

  function licensePricing() {
    const p = state.licenseCatalog?.product || {};
    const formMargin = licenseForm?.licenseMarginPercent
      ? Number(licenseForm.licenseMarginPercent.value)
      : NaN;
    const formRate = licenseForm?.eurToChfRate
      ? Number(licenseForm.eurToChfRate.value)
      : NaN;
    return {
      marginPercent: Number.isFinite(formMargin)
        ? formMargin
        : Number(p.licenseMarginPercent ?? 28),
      eurToChfRate: Number.isFinite(formRate) && formRate > 0
        ? formRate
        : Number(p.eurToChfRate ?? 0.93),
      sellCurrency: p.offerCurrency || "CHF",
      icCurrency: p.currency || "EUR",
    };
  }

  function icEurToSellChf(eurAmount) {
    const { marginPercent, eurToChfRate } = licensePricing();
    return Math.round(Number(eurAmount || 0) * (1 + marginPercent / 100) * eurToChfRate * 100) / 100;
  }

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

  function haversineKm(lat1, lon1, lat2, lon2) {
    const toRad = (d) => (d * Math.PI) / 180;
    const r = 6371;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    return 2 * r * Math.asin(Math.sqrt(a));
  }

  function initAddressAutocomplete(input) {
    if (!input) return;
    const list = document.getElementById("addressSuggestions");
    if (!list) return;

    let timer = null;
    let items = [];
    let activeIndex = -1;
    let userGeo = null;
    let requestSeq = 0;

    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => {
          userGeo = { lat: pos.coords.latitude, lon: pos.coords.longitude };
        },
        () => {},
        { enableHighAccuracy: false, timeout: 4000, maximumAge: 600000 },
      );
    }

    function closeList() {
      list.hidden = true;
      list.innerHTML = "";
      input.setAttribute("aria-expanded", "false");
      activeIndex = -1;
    }

    function openList() {
      list.hidden = false;
      input.setAttribute("aria-expanded", "true");
    }

    function applyActive() {
      [...list.querySelectorAll("li")].forEach((li, i) => {
        li.classList.toggle("active", i === activeIndex);
      });
    }

    function selectItem(item) {
      if (!item) return;
      input.value = item.label;
      closeList();
      input.dispatchEvent(new Event("change", { bubbles: true }));
    }

    function render(suggestions) {
      items = suggestions;
      activeIndex = suggestions.length ? 0 : -1;
      if (!suggestions.length) {
        closeList();
        return;
      }
      list.innerHTML = suggestions
        .map((s, i) => {
          const dist =
            s.distanceKm != null && Number.isFinite(s.distanceKm)
              ? `<span class="suggest-meta">${s.distanceKm < 10 ? s.distanceKm.toFixed(1) : Math.round(s.distanceKm)} km</span>`
              : "";
          return `<li role="option" data-index="${i}"${i === 0 ? ' class="active"' : ""}>${escapeHtml(s.label)}${dist}</li>`;
        })
        .join("");
      openList();
    }

    async function search(q) {
      const seq = ++requestSeq;
      try {
        const data = await api(
          `/api/geo/address-suggest?q=${encodeURIComponent(q)}&limit=10`,
        );
        if (seq !== requestSeq) return;
        let suggestions = data.suggestions || [];
        if (userGeo) {
          suggestions = suggestions
            .map((s) => {
              const lat = Number(s.lat);
              const lon = Number(s.lon);
              const distanceKm =
                Number.isFinite(lat) && Number.isFinite(lon)
                  ? haversineKm(userGeo.lat, userGeo.lon, lat, lon)
                  : null;
              return { ...s, distanceKm };
            })
            .sort((a, b) => {
              if (a.distanceKm == null && b.distanceKm == null) return 0;
              if (a.distanceKm == null) return 1;
              if (b.distanceKm == null) return -1;
              return a.distanceKm - b.distanceKm;
            });
        }
        render(suggestions);
      } catch (_) {
        if (seq === requestSeq) closeList();
      }
    }

    input.addEventListener("input", () => {
      const q = input.value.trim();
      clearTimeout(timer);
      if (q.length < 2) {
        closeList();
        return;
      }
      timer = setTimeout(() => search(q), 220);
    });

    input.addEventListener("keydown", (ev) => {
      if (list.hidden || !items.length) return;
      if (ev.key === "ArrowDown") {
        ev.preventDefault();
        activeIndex = (activeIndex + 1) % items.length;
        applyActive();
      } else if (ev.key === "ArrowUp") {
        ev.preventDefault();
        activeIndex = (activeIndex - 1 + items.length) % items.length;
        applyActive();
      } else if (ev.key === "Enter" && activeIndex >= 0) {
        ev.preventDefault();
        selectItem(items[activeIndex]);
      } else if (ev.key === "Escape") {
        closeList();
      }
    });

    list.addEventListener("mousedown", (ev) => {
      const li = ev.target.closest("li[data-index]");
      if (!li) return;
      ev.preventDefault();
      selectItem(items[Number(li.dataset.index)]);
    });

    input.addEventListener("blur", () => {
      setTimeout(closeList, 120);
    });
  }

  function syncCustomerProject(fromForm, toForm) {
    ["company", "projectName", "preparedBy"].forEach((field) => {
      if (fromForm[field] && toForm[field]) {
        toForm[field].value = fromForm[field].value;
      }
    });
  }

  function applyLicenseSelectionToIt() {
    // Kunde / Projekt / Ersteller
    syncCustomerProject(licenseForm, itForm);

    // Mengen aus Lizenz ableiten
    const instances = Math.max(1, Number(licenseForm.instanceCount?.value || 1));
    const extraOpenings = Math.max(0, Number(licenseForm.extraOpeningClients?.value || 0));
    itForm.deviceCount.value = instances;
    itForm.openingCount.value = Math.max(instances, instances + extraOpenings);
    if (!Number(itForm.zoneCount.value)) itForm.zoneCount.value = 1;

    // Add-ons → IT-Optionen
    const selectedAddons = new Set(
      [...licenseForm.querySelectorAll('input[name="addon"]:checked')].map((el) => el.value)
    );
    Object.entries(LICENSE_TO_IT_OPTIONS).forEach(([licenseId, itId]) => {
      const checkbox = itForm.querySelector(`input[name="itOption"][value="${itId}"]`);
      if (checkbox) checkbox.checked = selectedAddons.has(licenseId);
    });

    // Advanced Instance enthält Order Handling → IT Order Handling vorsehen
    const orderHandling = itForm.querySelector('input[name="itOption"][value="orderHandling"]');
    if (orderHandling && selectedInstanceId() === "advanced") {
      orderHandling.checked = true;
    }

    // Mobile Terminal Clients > 0 oft zusammen mit Externen Lagerplätzen
    if (Number(licenseForm.mobileTerminalClients?.value || 0) > 0) {
      const ext = itForm.querySelector('input[name="itOption"][value="externalStorage"]');
      if (ext) ext.checked = true;
    }
  }

  function switchView(name) {
    if (name === "it") {
      applyLicenseSelectionToIt();
      recalcIt();
    }
    if (name === "license") {
      syncCustomerProject(itForm, licenseForm);
    }
    document.querySelectorAll(".view").forEach((el) => el.classList.remove("active"));
    document.querySelectorAll(".nav-link").forEach((el) => {
      el.classList.toggle("active", el.dataset.view === name);
    });
    document.getElementById(`view-${name}`).classList.add("active");
    if (name === "archive") loadArchive();
    if (name === "offer") {
      updateInternalContributionBox();
      if (state.licenseOffer || state.itOffer) {
        // still require explicit button, but keep print enabled if doc exists
        document.getElementById("btnPrintOffer").disabled = !document.querySelector(".offer-sheet");
      }
    }
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function formatDateDe(value) {
    if (!value) return "—";
    const raw = String(value).slice(0, 10);
    const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[3]}.${m[2]}.${m[1]}`;
    return String(value);
  }

  function updateInternalContributionBox() {
    const box = document.getElementById("internalContributionBox");
    if (!box) return;
    const lic = state.licenseOffer?.totals || null;
    const it = state.itOffer?.totals || null;
    const licSell = lic?.sellNetChf != null ? Number(lic.sellNetChf) : null;
    const itSell = it?.totalAmount != null ? Number(it.totalAmount) : null;
    const licDb = lic?.contributionMarginChf != null
      ? Number(lic.contributionMarginChf)
      : (licSell != null && lic?.costChf != null
        ? licSell - Number(lic.costChf)
        : null);
    const itDb = it?.contributionMarginChf != null
      ? Number(it.contributionMarginChf)
      : (it?.marginAmount != null ? Number(it.marginAmount) : null);
    if (licDb == null && itDb == null) {
      box.innerHTML = `<span class="muted">DB (intern): Lizenz- und/oder IT-Kalkulation ausfüllen.</span>`;
      return;
    }
    const parts = [];
    if (licDb != null && licSell != null) {
      const p = licSell > 0 ? Math.round((licDb / licSell) * 1000) / 10 : 0;
      parts.push(`Lizenzen ${money(licDb, "CHF")} (${p}%)`);
    }
    if (itDb != null && itSell != null) {
      const p = itSell > 0 ? Math.round((itDb / itSell) * 1000) / 10 : 0;
      parts.push(`IT ${money(itDb, "CHF")} (${p}%)`);
    }
    const totalDb = (licDb || 0) + (itDb || 0);
    const totalSell = (licSell || 0) + (itSell || 0);
    const totalPct = totalSell > 0 ? Math.round((totalDb / totalSell) * 1000) / 10 : 0;
    box.innerHTML = `
      <div><strong>DB intern: ${money(totalDb, "CHF")} · ${totalPct}%</strong></div>
      <span class="muted">${parts.join(" · ")} · vom Verkauf · erscheint nicht im Kundenangebot</span>`;
  }

  async function ensureCurrentCalcs() {
    // frische Live-Daten aus den Formularen holen
    applyLicenseSelectionToIt();
    await Promise.all([recalcLicense(), recalcIt()]);
  }

  async function composeOfferDocument({ save = true, notify = false } = {}) {
    await ensureCurrentCalcs();
    if (!state.licenseOffer && !state.itOffer) {
      alert("Bitte zuerst Lizenz- und/oder IT-Kalkulation ausfüllen.");
      return;
    }
    const body = {
      license: state.licenseOffer || null,
      it: state.itOffer || null,
      basedOnOfferNumber: state.editingFromOfferId || null,
      discountPercent: Number(document.getElementById("offerDiscountPercent")?.value || 0) || 0,
      discountAmountChf: Number(document.getElementById("offerDiscountAmount")?.value || 0) || 0,
      roundTo: Number(document.getElementById("offerRoundTo")?.value || 0) || 0,
    };
    try {
      const endpoint = save ? "/api/offer/compose/save" : "/api/offer/compose";
      const doc = await api(endpoint, {
        method: "POST",
        body: JSON.stringify(body),
      });
      state.offerDocument = doc;
      state.savedOfferId = doc.id || doc.meta.offerNumber;
      const previousId = state.editingFromOfferId;
      state.editingFromOfferId = null;
      renderOfferDocument(doc);
      document.getElementById("btnPrintOffer").disabled = false;
      document.getElementById("btnSaveOfferDoc").disabled = false;
      document.getElementById("btnExcelOfferDoc").disabled = !state.savedOfferId;
      document.getElementById("btnWordOfferDoc").disabled = !state.savedOfferId;
      switchView("offer");
      if (notify || previousId) {
        const rev = previousId ? `\n(Neu erzeugt aus ${previousId})` : "";
        alert(`Gespeichert: ${doc.meta.offerNumber}${rev}`);
      }
    } catch (err) {
      alert(err.message);
    }
  }

  function setFormValue(form, name, value) {
    const el = form.elements.namedItem(name);
    if (!el) return;
    if (el instanceof RadioNodeList) {
      el.value = value == null ? "" : String(value);
      return;
    }
    el.value = value == null ? "" : String(value);
  }

  function fillLicenseFormFromOffer(offer) {
    if (!offer) return;
    const c = offer.customer || {};
    const cfg = offer.configuration || {};
    setFormValue(licenseForm, "company", c.company || "");
    setFormValue(licenseForm, "projectName", c.projectName || "");
    setFormValue(licenseForm, "contact", c.contact || "");
    setFormValue(licenseForm, "email", c.email || "");
    setFormValue(licenseForm, "address", c.address || "");
    setFormValue(licenseForm, "preparedBy", cfg.preparedBy || "");
    setFormValue(licenseForm, "notes", cfg.notes || "");
    setFormValue(licenseForm, "instanceCount", cfg.instanceCount ?? 1);
    setFormValue(licenseForm, "extraOpeningClients", cfg.extraOpeningClients ?? 0);
    setFormValue(licenseForm, "extraAdminClients", cfg.extraAdminClients ?? 0);
    setFormValue(licenseForm, "mobileTerminalClients", cfg.mobileTerminalClients ?? 0);
    setFormValue(licenseForm, "thirdPartyVlmTypes", cfg.thirdPartyVlmTypes ?? 0);
    setFormValue(licenseForm, "testInstances", cfg.testInstances ?? 0);
    setFormValue(licenseForm, "upgradeYears", cfg.upgradeYears ?? 0);
    const defaults = state.licenseCatalog?.product || {};
    const totals = offer.totals || {};
    setFormValue(
      licenseForm,
      "licenseMarginPercent",
      cfg.licenseMarginPercent ?? totals.marginPercent ?? defaults.licenseMarginPercent ?? 28
    );
    setFormValue(
      licenseForm,
      "eurToChfRate",
      cfg.eurToChfRate ?? totals.eurToChfRate ?? defaults.eurToChfRate ?? 0.93
    );

    const instanceId = cfg.instanceId || "basic";
    const radio = licenseForm.querySelector(`input[name="instanceId"][value="${instanceId}"]`);
    if (radio) {
      radio.checked = true;
      licenseForm.querySelectorAll("#instanceOptions .option").forEach((el) => el.classList.remove("selected"));
      radio.closest(".option")?.classList.add("selected");
    }
    renderIncluded();
    renderAddons();
    updateClientHints();

    const selected = new Set(cfg.selectedAddons || []);
    licenseForm.querySelectorAll('input[name="addon"]').forEach((el) => {
      el.checked = selected.has(el.value) && !el.disabled;
    });
  }

  function fillItFormFromOffer(offer) {
    if (!offer) return;
    const c = offer.customer || {};
    const cfg = offer.configuration || {};
    setFormValue(itForm, "company", c.company || "");
    setFormValue(itForm, "projectName", c.projectName || "");
    setFormValue(itForm, "preparedBy", cfg.preparedBy || "");
    setFormValue(itForm, "notes", cfg.notes || "");
    setFormValue(itForm, "realizationPeriod", cfg.realizationPeriod || "");
    setFormValue(itForm, "deviceCount", cfg.deviceCount ?? 1);
    setFormValue(itForm, "zoneCount", cfg.zoneCount ?? 1);
    setFormValue(itForm, "openingCount", cfg.openingCount ?? 1);
    setFormValue(itForm, "trips", cfg.trips ?? 0);
    setFormValue(itForm, "travelHoursPerTrip", cfg.travelHoursPerTrip ?? 0);
    setFormValue(itForm, "kmPerTrip", cfg.kmPerTrip ?? 0);
    setFormValue(itForm, "overnightCount", cfg.overnightCount ?? 0);
    setFormValue(itForm, "mealCount", cfg.mealCount ?? 0);
    const totals = offer.totals || {};
    const defaultMargin = state.itCatalog?.rates?.marginPercent ?? 28;
    setFormValue(
      itForm,
      "itMarginPercent",
      cfg.itMarginPercent ?? totals.marginPercent ?? defaultMargin
    );

    const opts = cfg.options || {};
    itForm.querySelectorAll('input[name="itOption"]').forEach((el) => {
      el.checked = Boolean(opts[el.value]);
    });

    const exts = cfg.customExtensions || [];
    for (let i = 1; i <= 5; i += 1) {
      const ext = exts[i - 1] || {};
      setFormValue(itForm, `extDesc${i}`, ext.description || "");
      setFormValue(itForm, `extHours${i}`, ext.hours ?? 0);
    }
  }

  async function openArchiveForEdit(id) {
    const offer = await api(`/api/offers/${encodeURIComponent(id)}`);
    state.editingFromOfferId = offer.meta?.offerNumber || offer.id || id;

    if (offer.kind === "license") {
      fillLicenseFormFromOffer(offer);
      applyLicenseSelectionToIt();
      await Promise.all([recalcLicense(), recalcIt()]);
      switchView("license");
      alert(`Lizenzkalkulation geladen: ${state.editingFromOfferId}\nAnpassen und danach unter „Angebot“ neu erzeugen.`);
      return;
    }

    if (offer.kind === "it") {
      fillItFormFromOffer(offer);
      await recalcIt();
      switchView("it");
      alert(`IT-Kalkulation geladen: ${state.editingFromOfferId}\nAnpassen und danach unter „Angebot“ neu erzeugen.`);
      return;
    }

    if (offer.kind === "offer_document") {
      const license = offer.editable?.license || null;
      const it = offer.editable?.it || null;
      // Fallback: Quellen aus Archiv nachladen
      let lic = license;
      let itOffer = it;
      if (!lic && offer.sources?.licenseOfferNumber) {
        try { lic = await api(`/api/offers/${encodeURIComponent(offer.sources.licenseOfferNumber)}`); }
        catch (_) { /* ignore */ }
      }
      if (!itOffer && offer.sources?.itOfferNumber) {
        try { itOffer = await api(`/api/offers/${encodeURIComponent(offer.sources.itOfferNumber)}`); }
        catch (_) { /* ignore */ }
      }
      if (!lic && !itOffer) {
        // Mindest-Fallback aus Konfigurationszusammenfassung
        const cfg = offer.content?.configurationSummary || {};
        const cust = offer.customer || {};
        fillLicenseFormFromOffer({
          customer: cust,
          configuration: {
            instanceId: (cfg.instanceName || "").toLowerCase().includes("advanced") ? "advanced" : "basic",
            instanceCount: cfg.instanceCount || 1,
            preparedBy: offer.meta?.preparedBy || "",
            selectedAddons: [],
            extraOpeningClients: Math.max(0, (cfg.openingCount || 1) - (cfg.deviceCount || cfg.instanceCount || 1)),
          },
        });
        fillItFormFromOffer({
          customer: cust,
          configuration: {
            deviceCount: cfg.deviceCount || 1,
            zoneCount: cfg.zoneCount || 1,
            openingCount: cfg.openingCount || 1,
            preparedBy: offer.meta?.preparedBy || "",
            options: { orderHandling: Boolean(cfg.hasOrderHandling) },
            customExtensions: [],
          },
        });
      } else {
        if (lic) fillLicenseFormFromOffer(lic);
        if (itOffer) fillItFormFromOffer(itOffer);
        if (lic && !itOffer) applyLicenseSelectionToIt();
      }
      await Promise.all([recalcLicense(), recalcIt()]);
      switchView("license");
      alert(
        `Gesamtangebot geladen: ${state.editingFromOfferId}\n` +
        `Jetzt bearbeiten (Lizenz/IT), danach „Angebot erzeugen“ für eine neue Version.`
      );
      return;
    }

    alert("Dieser Archiv-Eintrag kann nicht geladen werden.");
  }

  async function openArchivePreview(id) {
    const offer = await api(`/api/offers/${encodeURIComponent(id)}`);
    if (offer.kind === "offer_document") {
      state.offerDocument = offer;
      state.savedOfferId = offer.id || offer.meta?.offerNumber;
      renderOfferDocument(offer);
      document.getElementById("btnPrintOffer").disabled = false;
      document.getElementById("btnSaveOfferDoc").disabled = false;
      document.getElementById("btnExcelOfferDoc").disabled = !!state.savedOfferId;
      document.getElementById("btnWordOfferDoc").disabled = !!state.savedOfferId;
      switchView("offer");
      return;
    }
    // Einzelkalkulation → in Formular laden und Vorschau zeigen
    await openArchiveForEdit(id);
  }

  function downloadOfferDocx() {
    const id = state.savedOfferId || state.offerDocument?.id || state.offerDocument?.meta?.offerNumber;
    if (!id) {
      alert("Bitte das Angebot zuerst erzeugen/speichern.");
      return;
    }
    window.location.href = `/api/offers/${encodeURIComponent(id)}/docx`;
  }

  function renderOfferDocument(doc) {
    const root = document.getElementById("offerDocument");
    const c = doc.customer || {};
    const cfg = doc.content.configurationSummary || {};
    const summary = doc.priceSummary || {};
    const licSum = summary.license;
    const itSum = summary.it;
    const arch = doc.content.architecture || {};
    const req = doc.content.requirements || {};
    const scopeGroups = doc.scopeGroups || doc.content?.scopeGroups || [];

    const scopeHtml = scopeGroups.map((group) => {
      const items = (group.items || [])
        .map((item) => `
          <div class="offer-fn">
            <div class="offer-fn-mark">•</div>
            <div>
              <h4>${escapeHtml(item.name || "")}</h4>
              ${item.description ? `<p>${escapeHtml(item.description)}</p>` : ""}
            </div>
          </div>`)
        .join("");
      return `
        <div class="offer-scope-group">
          <h3>${escapeHtml(group.title || "")}</h3>
          <div class="offer-fn-list">${items || "<p class=\"offer-empty-note\">Keine Positionen.</p>"}</div>
          ${group.total != null ? `
            <div class="offer-scope-total">
              <span>Total ${escapeHtml(group.title || "")}</span>
              <strong>${money(group.total, group.currency || "CHF")}</strong>
            </div>` : ""}
        </div>`;
    }).join("");

    const grand = summary.grandTotalChf;
    const customerNote = summary.note || "Alle Preise in CHF, exkl. MwSt.";

    const priceCards = `
      <div class="offer-price-summary">
        ${licSum ? `
          <div class="offer-price-card">
            <span>Softwarelizenzen</span>
            <strong>${money(licSum.total, licSum.currency || "CHF")}</strong>
            <small>Gesamttotal Lizenzen · exkl. MwSt.</small>
          </div>` : ""}
        ${itSum ? `
          <div class="offer-price-card">
            <span>IT-Aufwand / Installation</span>
            <strong>${money(itSum.total, itSum.currency || "CHF")}</strong>
            <small>Gesamttotal IT inkl. Reise · exkl. MwSt.</small>
          </div>` : ""}
        <div class="offer-price-card">
          <span>Gesamttotal CHF</span>
          <strong>${grand != null ? money(grand, "CHF") : "—"}</strong>
          <small>${escapeHtml(customerNote)}</small>
        </div>
      </div>`;

    const standardHtml = (doc.content.standardFunctions || [])
      .map((fn, idx) => `
        <div class="offer-fn">
          <div class="offer-fn-num">1.1.${idx + 1}</div>
          <div>
            <h4>${escapeHtml(fn.title)}</h4>
            <p>${escapeHtml(fn.text)}</p>
          </div>
        </div>`)
      .join("");

    const selectedOpts = doc.content.selectedOptions || [];
    const optionsHtml = selectedOpts.length
      ? selectedOpts.map((opt) => `
          <div class="offer-fn">
            <div class="offer-fn-mark">✓</div>
            <div>
              <h4>${escapeHtml(opt.title)}</h4>
              <p>${escapeHtml(opt.text)}</p>
            </div>
          </div>`).join("")
      : `<p class="offer-empty-note">Keine optionalen Softwaremodule gewählt.</p>`;

    const hardwareHtml = (doc.content.hardwareOptions || [])
      .map((opt) => `
        <div class="offer-hw-item">
          <strong>${escapeHtml(opt.title)}</strong>
          <span>${escapeHtml(opt.text)}</span>
        </div>`)
      .join("");

    const respHtml = (doc.content.responsibilities || [])
      .map((r) => `
        <tr>
          <td>${escapeHtml(r.task)}</td>
          <td class="center">${r.ssi ? "X" : ""}</td>
          <td class="center">${r.customer ? "X" : ""}</td>
        </tr>`)
      .join("");

    const listItems = (items) =>
      (items || []).map((x) => `<li>${escapeHtml(x)}</li>`).join("");

    const cfgBits = [
      cfg.instanceName
        ? `${escapeHtml(cfg.instanceName)}${cfg.instanceCount ? ` (${cfg.instanceCount}×)` : ""}`
        : "",
      cfg.deviceCount ? `${cfg.deviceCount} Geräte` : "",
      cfg.zoneCount ? `${cfg.zoneCount} Zone(n)` : "",
      cfg.openingCount ? `${cfg.openingCount} Öffnung(en)` : "",
      cfg.hasOrderHandling ? "Order Handling" : "Standalone",
    ].filter(Boolean).join(" · ");

    root.innerHTML = `
      <div class="offer-sheet">
        <header class="offer-cover">
          <div class="offer-letterhead">
            <div class="offer-letterhead-brand">
              <div class="offer-wamas-mark">WAMAS<span>Lift &amp; Store</span></div>
              <p>Softwarelösung für Vertical Lift Modules</p>
            </div>
            <img class="offer-logo-ssi" src="/static/assets/ssi-schaefer.png" alt="SSI SCHÄFER" />
          </div>
          <div class="offer-cover-band">
            <p class="offer-doc-label">${escapeHtml(doc.content.documentLabel || "Angebot / Preisliste")}</p>
            <h1>${escapeHtml(doc.content.title)}</h1>
            <p class="offer-cover-sub">${escapeHtml(doc.content.subtitle)}</p>
          </div>
          <div class="offer-cover-meta">
            <div>
              <span>Version</span>
              <strong>${escapeHtml(doc.branding.version || "2.8")}</strong>
            </div>
            <div>
              <span>Datum</span>
              <strong>${escapeHtml(doc.meta.documentDate || formatDateDe(doc.meta.createdAt))}</strong>
            </div>
            <div>
              <span>Angebotsnummer</span>
              <strong>${escapeHtml(doc.meta.offerNumber)}</strong>
            </div>
            <div>
              <span>Gültig bis</span>
              <strong>${escapeHtml(formatDateDe(doc.meta.validUntil))}</strong>
            </div>
            ${doc.meta.revisionOf || doc.sources?.basedOnOfferNumber ? `
            <div>
              <span>Revision von</span>
              <strong>${escapeHtml(doc.meta.revisionOf || doc.sources.basedOnOfferNumber)}</strong>
            </div>` : ""}
          </div>
          <div class="offer-party-grid">
            <div class="offer-party">
              <h3>Kunde / Projekt</h3>
              <p><strong>${escapeHtml(c.company || "—")}</strong></p>
              <p>${escapeHtml(c.projectName || "—")}</p>
              <p>${escapeHtml(c.contact || "")}${c.email ? ` · ${escapeHtml(c.email)}` : ""}</p>
              <p class="muted">${c.address ? escapeHtml(c.address) : ""}</p>
            </div>
            <div class="offer-party">
              <h3>SSI SCHÄFER</h3>
              <p><strong>Erstellt von</strong><br>${escapeHtml(doc.meta.preparedBy || "—")}</p>
              <p><strong>Konfiguration</strong><br>${cfgBits || "—"}</p>
              <p class="muted">Quellen: ${escapeHtml(doc.sources?.licenseOfferNumber || "Lizenz")} · ${escapeHtml(doc.sources?.itOfferNumber || "IT")}</p>
            </div>
          </div>
          ${priceCards}
        </header>

        <section class="offer-section offer-section-prices" id="sec-commercial">
          <h2>1. Leistungsumfang &amp; Preise</h2>
          <p>Auflistung der Leistungen ohne Einzelpreise. Ausgewiesen werden die Bereichstotals und das Gesamttotal.</p>
          ${scopeHtml || "<p class=\"offer-empty-note\">Kein Leistungsumfang kalkuliert.</p>"}
          <div class="offer-totals-box">
            ${licSum ? `<div><span>Softwarelizenzen</span><strong>${money(licSum.total, licSum.currency || "CHF")}</strong></div>` : ""}
            ${itSum ? `<div><span>IT-Aufwand inkl. Reise</span><strong>${money(itSum.total, itSum.currency || "CHF")}</strong></div>` : ""}
            ${summary.subtotalChf != null && (summary.commercialDiscountChf || summary.roundingAmountChf) ? `<div><span>Zwischentotal</span><strong>${money(summary.subtotalChf, "CHF")}</strong></div>` : ""}
            ${summary.commercialDiscountChf ? `<div><span>Projektrabatt</span><strong>− ${money(summary.commercialDiscountChf, "CHF")}</strong></div>` : ""}
            ${summary.roundingAmountChf ? `<div><span>Abrundung</span><strong>− ${money(summary.roundingAmountChf, "CHF")}</strong></div>` : ""}
            ${grand != null ? `<div class="offer-totals-grand"><span>Gesamttotal exkl. MwSt</span><strong>${money(grand, "CHF")}</strong></div>` : ""}
          </div>
          <p class="offer-note">${escapeHtml(customerNote)}</p>
        </section>

        <div class="offer-annex-start">
          <p class="offer-doc-label">${escapeHtml(doc.content.annexLabel || "Anhang zur Software")}</p>
          <h2>Leistungsbeschreibung WAMAS® Lift &amp; Store</h2>
        </div>

        <nav class="offer-toc" aria-label="Inhaltsverzeichnis Anhang">
          <h2>Inhaltsverzeichnis Anhang</h2>
          <ol>
            <li><a href="#sec-scope">Umfang WAMAS Lift &amp; Store</a>
              <ol>
                <li><a href="#sec-functions">Standard-Funktionen / Prozesse</a></li>
                <li><a href="#sec-options">Mögliche Optionen</a></li>
                <li><a href="#sec-clients">Bedienoberflächen</a></li>
              </ol>
            </li>
            <li><a href="#sec-architecture">Standard-Systemarchitektur</a></li>
            <li><a href="#sec-requirements">Anforderungen</a></li>
            <li><a href="#sec-responsibilities">Zuständigkeiten</a></li>
            <li><a href="#sec-documents">Begleitende Dokumente</a></li>
            <li><a href="#sec-terms">Kaufmännische Bedingungen (Schweiz)</a></li>
          </ol>
        </nav>

        <section class="offer-section" id="sec-scope">
          <h2>A1. Umfang WAMAS Lift &amp; Store</h2>
          <p>${escapeHtml(doc.content.intro)}</p>
          ${doc.content.introVariant ? `<p><strong>${escapeHtml(doc.content.introVariant)}</strong></p>` : ""}
        </section>

        <section class="offer-section" id="sec-functions">
          <h2>A1.1 Standard-Funktionen / Prozesse</h2>
          <p>${escapeHtml(doc.content.standardLead || "")}</p>
          <p>${escapeHtml(doc.content.recommendation)}</p>
          <div class="offer-footnotes">
            ${(doc.content.footnotes || []).map((f) => `<div>${escapeHtml(f)}</div>`).join("")}
          </div>
          <div class="offer-fn-list">${standardHtml}</div>
        </section>

        <section class="offer-section" id="sec-options">
          <h2>A1.2 Mögliche Optionen für WAMAS® Lift &amp; Store</h2>
          <p>${escapeHtml(doc.content.machineOptionsLead || "")}</p>
          <h3>Gewählte Software-Optionen</h3>
          <div class="offer-fn-list">${optionsHtml}</div>
          <h3>A1.2.1 Hardware-Optionen vom WAMAS® Lift &amp; Store</h3>
          <div class="offer-hw-grid">${hardwareHtml}</div>
        </section>

        <section class="offer-section" id="sec-clients">
          <h2>A1.3 Bedienoberfläche Bediener und Admin Client</h2>
          <div class="offer-client-grid">
            <div class="offer-client">
              <h4>Touch Client</h4>
              <p>${escapeHtml(doc.content.clients.touch)}</p>
            </div>
            <div class="offer-client">
              <h4>Admin Client</h4>
              <p>${escapeHtml(doc.content.clients.admin)}</p>
            </div>
            <div class="offer-client">
              <h4>Mobile Terminal</h4>
              <p>${escapeHtml(doc.content.clients.mobile)}</p>
            </div>
          </div>
        </section>

        <section class="offer-section" id="sec-architecture">
          <h2>A2. ${escapeHtml(arch.title || "Standard-Systemarchitektur")}</h2>
          <p>${escapeHtml(arch.text || "")}</p>
          ${arch.image ? `<figure class="offer-figure"><img src="${escapeHtml(arch.image)}" alt="Systemarchitektur WAMAS Lift & Store" /></figure>` : ""}
          <ul class="offer-legend">
            ${(arch.legend || []).map((x) => `<li>${escapeHtml(x)}</li>`).join("")}
          </ul>
        </section>

        <section class="offer-section" id="sec-requirements">
          <h2>A3. ${escapeHtml(req.title || "Anforderungen")}</h2>
          <p>${escapeHtml(req.note || "")}</p>
          <div class="offer-req-grid">
            <div>
              <h3>Server</h3>
              <ul>${listItems(req.server)}</ul>
            </div>
            <div>
              <h3>Desktop / Admin Client</h3>
              <ul>${listItems(req.desktop)}</ul>
            </div>
            <div>
              <h3>Touch Client / IPC</h3>
              <ul>${listItems(req.touch)}</ul>
            </div>
          </div>
          ${req.networkHighlight ? `<p class="offer-note"><strong>Netzwerk:</strong> ${escapeHtml(req.networkHighlight)}</p>` : ""}
        </section>

        <section class="offer-section" id="sec-responsibilities">
          <h2>A4. Zuständigkeiten</h2>
          <h3>Endabnahme</h3>
          <p>${escapeHtml(doc.content.acceptance || "")}</p>
          <h3>Zuständigkeitsmatrix</h3>
          <table class="offer-resp">
            <thead>
              <tr><th>Aufgabe</th><th>SSI</th><th>Kunde</th></tr>
            </thead>
            <tbody>${respHtml}</tbody>
          </table>
        </section>

        <section class="offer-section" id="sec-documents">
          <h2>A5. Begleitende Dokumente</h2>
          <p>${escapeHtml(doc.content.documentsLead || "")}</p>
          <ul class="offer-doc-list">
            ${(doc.content.documents || []).map((d) => `<li>${escapeHtml(d)}</li>`).join("")}
          </ul>
          <p>${escapeHtml(doc.content.closing)}</p>
        </section>

        ${renderCommercialTermsHtml(doc.content.commercialTerms)}

        <div class="offer-footer">
          <div>
            <div class="offer-wamas-mark offer-wamas-mark-sm">WAMAS<span>Lift &amp; Store</span></div>
            <div>${escapeHtml(doc.meta.offerNumber)} · ${escapeHtml(doc.meta.documentDate || formatDateDe(doc.meta.createdAt))} · gültig bis ${escapeHtml(formatDateDe(doc.meta.validUntil))}</div>
          </div>
          <img class="offer-logo-ssi offer-logo-ssi-sm" src="/static/assets/ssi-schaefer.png" alt="SSI SCHÄFER" />
        </div>
      </div>`;
  }

  function linkifyUrls(text) {
    return escapeHtml(text).replace(
      /(https?:\/\/[^\s]+)/g,
      '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>'
    );
  }

  function renderCommercialTermsHtml(terms) {
    if (!terms || !(terms.sections || []).length) return "";
    const sectionsHtml = (terms.sections || []).map((sec) => {
      const paras = (sec.paragraphs || [])
        .map((p) => `<p>${linkifyUrls(p)}</p>`)
        .join("");
      const bullets = (sec.bullets || []).length
        ? `<ul>${sec.bullets.map((b) => `<li>${escapeHtml(b)}</li>`).join("")}</ul>`
        : "";
      const after = (sec.paragraphsAfter || [])
        .map((p) => `<p>${linkifyUrls(p)}</p>`)
        .join("");
      const subs = (sec.subsections || [])
        .map((sub) => `
          <div class="offer-terms-sub">
            <h4>${escapeHtml(sub.title)}</h4>
            ${(sub.paragraphs || []).map((p) => `<p>${linkifyUrls(p)}</p>`).join("")}
          </div>`)
        .join("");
      const table = sec.table
        ? `<table class="offer-terms-table">
            <thead><tr>${sec.table.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
            <tbody>
              ${sec.table.rows.map((row) => `<tr>${row.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("")}
            </tbody>
          </table>`
        : "";
      return `
        <article class="offer-terms-section" id="term-${escapeHtml(sec.id)}">
          <h3>${escapeHtml(sec.id)} ${escapeHtml(sec.title)}</h3>
          ${paras}${bullets}${after}${subs}${table}
        </article>`;
    }).join("");

    const closing = terms.closing || {};
    const signs = (closing.signatories || [])
      .map((s) => `
        <div class="offer-signatory">
          <div class="offer-sign-line"></div>
          <strong>${escapeHtml(s.name)}</strong>
          <span>${escapeHtml(s.title || "")}</span>
          ${s.role ? `<span>${escapeHtml(s.role)}</span>` : ""}
        </div>`)
      .join("");

    return `
      <section class="offer-section offer-terms" id="sec-terms">
        <div class="offer-terms-head">
          <p class="offer-doc-label">Rechtliche Bedingungen</p>
          <h2>${escapeHtml(terms.title || "Kaufmännische Bedingungen")}</h2>
          <p>Version ${escapeHtml(terms.version || "")} · Angebotsgültigkeit ${escapeHtml(String(terms.validityDays || 14))} Tage</p>
        </div>
        ${sectionsHtml}
        <div class="offer-terms-closing">
          <p>${escapeHtml(closing.text || "")}</p>
          <p><strong>${escapeHtml(closing.greeting || "Freundliche Grüsse")}</strong><br>${escapeHtml(closing.company || "SSI SCHÄFER AG")}</p>
          <div class="offer-signatories">${signs}</div>
          <div class="offer-customer-accept">
            <div>
              <p>Ort / Datum</p>
              <div class="offer-sign-line"></div>
            </div>
            <div>
              <p>Unterschrift / Stempel Kunde</p>
              <div class="offer-sign-line"></div>
            </div>
          </div>
        </div>
      </section>`;
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
    const pricing = licensePricing();
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
      licenseMarginPercent: pricing.marginPercent,
      eurToChfRate: pricing.eurToChfRate,
    };
  }

  function renderInstances() {
    const root = document.getElementById("instanceOptions");
    const previousId = selectedInstanceId();
    root.innerHTML = "";
    state.licenseCatalog.instances.forEach((inst, index) => {
      const selected = previousId
        ? inst.id === previousId
        : index === 0;
      const label = document.createElement("label");
      label.className = `option${selected ? " selected" : ""}`;
      label.innerHTML = `
        <input type="radio" name="instanceId" value="${inst.id}" ${selected ? "checked" : ""} />
        <div>
          <h3>${inst.name}</h3>
          <p>${inst.description}</p>
          <p class="muted">${inst.functionalSummary || ""}</p>
        </div>
        <div class="price">${money(icEurToSellChf(inst.price), "CHF")}<small class="muted">Einkauf ${money(inst.price, "EUR")}</small></div>`;
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
      el.className = `option-row${available ? "" : " locked"}`;
      el.innerHTML = `
        <input type="checkbox" name="addon" value="${addon.id}"
          ${available && prev.has(addon.id) ? "checked" : ""} ${available ? "" : "disabled"} />
        <div class="row-main">
          <strong>${addon.name}</strong>
          <span>${addon.functionalDescription || addon.description || ""}</span>
          ${available ? "" : "<span>Nur für Advanced</span>"}
        </div>
        <div class="row-price">
          ${available
            ? `${money(icEurToSellChf(addon.price), "CHF")}<small>Einkauf ${money(addon.price, "EUR")}</small>`
            : "—"}
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
        return `<div class="option-row info-only${locked ? " locked" : ""}">
          <div class="row-main">
            <strong>${c.name}</strong>
            <span>${c.functionalDescription || c.description}</span>
            ${locked ? "<span>Nur für Advanced</span>" : ""}
          </div>
          <div class="row-price">
            ${money(icEurToSellChf(c.price), "CHF")}
            <small>Einkauf ${money(c.price, "EUR")}</small>
          </div>
        </div>`;
      })
      .join("");
  }

  function renderLicensePreview(offer) {
    state.licenseOffer = offer;
    const t = offer.totals || {};
    const pricing = licensePricing();
    const ic = t.icCurrency || pricing.icCurrency;
    const sell = "CHF";
    const marginPercent = t.marginPercent ?? pricing.marginPercent;
    const fx = t.eurToChfRate ?? pricing.eurToChfRate;
    const sellChf = t.sellNetChf != null
      ? Number(t.sellNetChf)
      : icEurToSellChf(t.net);
    const sellEur = t.sellNetEur != null
      ? Number(t.sellNetEur)
      : Math.round(Number(t.net || 0) * (1 + marginPercent / 100) * 100) / 100;
    const marginEur = t.marginAmountEur != null
      ? Number(t.marginAmountEur)
      : Math.round((sellEur - Number(t.net || 0)) * 100) / 100;

    document.getElementById("licenseOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("licenseGross").textContent = money(sellChf, sell);
    document.getElementById("sllBadge").textContent =
      `SLL: ${t.sllCount} · Rabatt ${t.discountPercent}% · Marge ${marginPercent}% · Kurs ${fx}`;
    document.getElementById("licenseMeta").innerHTML = `
      <div><strong>${offer.customer.company}</strong>${offer.customer.projectName ? ` · ${offer.customer.projectName}` : ""}</div>
      <div>${offer.configuration.instanceCount}× ${offer.configuration.instanceName}</div>
      <div>Verkaufspreise CHF (Einkauf + ${marginPercent}% Marge, Kurs ${fx})</div>`;
    document.getElementById("licenseLines").innerHTML = offer.lines.map((line) => {
      const icTotal = line.totalIcEur != null ? Number(line.totalIcEur) : null;
      const chfTotal = line.currency === "CHF" && line.totalIcEur != null
        ? Number(line.total)
        : icEurToSellChf(icTotal != null ? icTotal : line.total);
      return `
      <tr>
        <td>${line.name}<div class="muted">${line.description || ""}${icTotal != null ? ` · Einkauf ${money(icTotal, ic)}` : ""}</div></td>
        <td>${line.qty}</td>
        <td>${money(chfTotal, sell)}</td>
      </tr>`;
    }).join("");
    const costChf = t.costChf != null ? Number(t.costChf) : Number(t.net || 0) * fx;
    const dbChf = t.contributionMarginChf != null
      ? Number(t.contributionMarginChf)
      : sellChf - costChf;
    const dbPercent = t.contributionMarginPercent != null
      ? Number(t.contributionMarginPercent)
      : (sellChf > 0 ? Math.round((dbChf / sellChf) * 1000) / 10 : 0);
    document.getElementById("licenseTotals").innerHTML = `
      <div class="row"><span>Einkauf Total</span><span>${money(t.net, ic)}</span></div>
      <div class="row"><span>Marge ${marginPercent}%</span><span>+ ${money(marginEur, ic)}</span></div>
      <div class="row"><span>Verkauf EUR</span><span>${money(sellEur, ic)}</span></div>
      <div class="row"><span>Einkauf CHF (× ${fx})</span><span>${money(costChf, sell)}</span></div>
      <div class="row strong"><span>Verkauf CHF</span><span>${money(sellChf, sell)}</span></div>
      <div class="row db"><span>DB (intern)</span><span>${money(dbChf, sell)} · ${dbPercent}%</span></div>`;
    document.getElementById("licenseScope").innerHTML =
      (offer.scopeOfSupply || []).map((s) => `<li>${s}</li>`).join("");
    document.getElementById("licenseDisclaimer").textContent = offer.product.disclaimer;
    document.getElementById("btnLicenseExcel").disabled = !state.savedLicenseId;
    document.getElementById("btnLicensePrint").disabled = !offer;
    updateInternalContributionBox();
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
  function itOptionHoursHint(opt) {
    const effort = state.itCatalog.effort || {};
    const direct = effort[`${opt.id}Hours`];
    if (direct != null) return `${direct} h`;
    if (opt.id === "rfid" && effort.rfidBaseHours != null) {
      return `ab ${effort.rfidBaseHours} h`;
    }
    if (opt.id === "scanner" && effort.scannerBaseHours != null) {
      return `ab ${effort.scannerBaseHours} h`;
    }
    return "";
  }

  function renderItOptions() {
    const root = document.getElementById("itOptions");
    const prev = new Set([...root.querySelectorAll('input[name="itOption"]:checked')].map((el) => el.value));
    root.innerHTML = "";
    state.itCatalog.options.forEach((opt) => {
      const hours = itOptionHoursHint(opt);
      const el = document.createElement("label");
      el.className = hours ? "option-row" : "option-row no-price";
      el.innerHTML = `
        <input type="checkbox" name="itOption" value="${opt.id}" ${prev.has(opt.id) ? "checked" : ""} />
        <div class="row-main">
          <strong>${opt.name}</strong>
          <span>${opt.description}</span>
        </div>
        ${hours ? `<div class="row-price">${hours}<small>Richtwert</small></div>` : ""}`;
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
      itMarginPercent: (() => {
        const v = Number(data.get("itMarginPercent"));
        return Number.isFinite(v) ? v : (state.itCatalog?.rates?.marginPercent ?? 28);
      })(),
    };
  }

  function renderItPreview(offer) {
    state.itOffer = offer;
    const c = offer.totals.currency || "CHF";
    const t = offer.totals || {};
    const marginPercent = t.marginPercent ?? offer.configuration?.itMarginPercent ?? 28;
    document.getElementById("itOfferNo").textContent = offer.meta.offerNumber;
    document.getElementById("itGross").textContent = money(t.totalAmount, c);
    document.getElementById("itMeta").innerHTML = `
      <div><strong>${offer.customer.company}</strong>${offer.customer.projectName ? ` · ${offer.customer.projectName}` : ""}</div>
      <div>${offer.configuration.deviceCount} Geräte · ${offer.configuration.zoneCount} Zone(n) · ${offer.configuration.openingCount} Öffnung(en)</div>
      <div>Stundensatz Einkauf ${money(offer.configuration.hourlyRate, c)} · Verkauf ${money(offer.configuration.hourlyRateSell ?? offer.configuration.hourlyRate, c)} · Marge ${marginPercent}%</div>
      <div>${offer.configuration.realizationPeriod || "–"} · CHF intern (kein Kurs)</div>`;

    const visibleLines = offer.lines.filter((l) => {
      if (l.category === "option") return Boolean(l.selected) && (Number(l.hours) > 0 || Number(l.amount) > 0 || l.note);
      if (l.category === "travel") return Number(l.amount) > 0 || Number(l.hours) > 0;
      if (l.category === "custom") return Number(l.hours) > 0 || Number(l.amount) > 0;
      return Number(l.hours) > 0 || Number(l.amount) > 0 || ["IT-DEVICES", "IT-ZONES", "IT-OPENINGS"].includes(l.sku);
    });

    document.getElementById("itLines").innerHTML = visibleLines.map((line) => `
        <tr>
          <td>${line.name}<div class="muted">${line.description || ""}${line.note ? ` · ${line.note}` : ""}${line.amountCost != null ? ` · Einkauf ${money(line.amountCost, c)}` : ""}</div></td>
          <td>${line.hours || 0}</td>
          <td>${money(line.amount, c)}</td>
        </tr>`).join("");
    const itCost = Number(t.totalAmountCost ?? 0);
    const itSell = Number(t.totalAmount || 0);
    const itDb = t.contributionMarginChf != null
      ? Number(t.contributionMarginChf)
      : Number(t.marginAmount || 0);
    const itDbPercent = t.contributionMarginPercent != null
      ? Number(t.contributionMarginPercent)
      : (itSell > 0 ? Math.round((itDb / itSell) * 1000) / 10 : 0);
    document.getElementById("itTotals").innerHTML = `
      <div class="row"><span>IT-Aufwand Einkauf</span><span>${t.workHours} h · ${money(t.workAmountCost ?? t.workAmount, c)}</span></div>
      <div class="row"><span>Reisekosten Einkauf</span><span>${money(t.travelAmountCost ?? t.travelAmount, c)}</span></div>
      <div class="row"><span>Einkauf Total</span><span>${money(itCost, c)}</span></div>
      <div class="row"><span>Marge ${marginPercent}%</span><span>+ ${money(t.marginAmount || 0, c)}</span></div>
      <div class="row strong"><span>Verkauf Total exkl. MwSt</span><span>${money(itSell, c)}</span></div>
      <div class="row db"><span>DB (intern)</span><span>${money(itDb, c)} · ${itDbPercent}%</span></div>`;
    document.getElementById("itScope").innerHTML = (offer.offerSections || [])
      .filter((s) => Number(s.amount) > 0 || (s.bullets || []).length)
      .map((s) => `<li><strong>${s.title}</strong> · ${money(s.amount, c)}<div class="muted">${(s.bullets || []).join(" · ")}</div></li>`)
      .join("");
    document.getElementById("itDisclaimer").textContent = offer.product.disclaimer;
    document.getElementById("btnItExcel").disabled = !state.savedItId;
    document.getElementById("btnItPrint").disabled = !offer;
    updateInternalContributionBox();
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
    archiveList.innerHTML = offers.map((o) => {
      const kindLabel = o.kind === "it" ? "IT" : o.kind === "offer_document" ? "Gesamtangebot" : "Lizenz";
      const amountText = o.kind === "offer_document"
        ? (o.amount || "—")
        : money(o.amount, o.currency || "CHF");
      return `
      <article class="archive-item" data-id="${escapeHtml(o.id)}" data-kind="${escapeHtml(o.kind || "")}">
        <div>
          <h3>${escapeHtml(o.offerNumber)} <span class="muted">(${kindLabel})</span></h3>
          <p>${escapeHtml(o.company || "—")} · ${escapeHtml(o.summary || "")}</p>
          <p>${escapeHtml(String(amountText))} · ${escapeHtml(o.createdAt || "")}</p>
        </div>
        <div class="archive-actions">
          <button type="button" class="btn primary" data-action="edit">Bearbeiten</button>
          ${o.kind === "offer_document" ? '<button type="button" class="btn" data-action="view">Anzeigen</button>' : ""}
          ${o.kind === "offer_document" ? '<button type="button" class="btn" data-action="docx">Word</button>' : ""}
          <button type="button" class="btn" data-action="excel">Excel</button>
          <button type="button" class="btn danger" data-action="delete">Löschen</button>
        </div>
      </article>`;
    }).join("");
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
    licenseForm.addEventListener("input", (event) => {
      syncCustomerProject(licenseForm, itForm);
      const name = event.target?.name;
      if (name === "licenseMarginPercent" || name === "eurToChfRate") {
        renderInstances();
        renderAddons();
        updateClientHints();
      }
      recalcLicense();
    });
    licenseForm.addEventListener("change", (event) => {
      // Bei Lizenzänderungen IT-Vorschläge mitziehen (inkl. External Storage → IT)
      applyLicenseSelectionToIt();
      syncCustomerProject(licenseForm, itForm);
      const name = event.target?.name;
      if (name === "licenseMarginPercent" || name === "eurToChfRate") {
        renderInstances();
        renderAddons();
        updateClientHints();
      }
      recalcLicense();
      // Falls IT-Tab aktiv ist, sofort neu rechnen
      if (document.getElementById("view-it").classList.contains("active")) {
        recalcIt();
      }
    });

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
      if (["company", "projectName", "preparedBy"].includes(t.name)) {
        syncCustomerProject(itForm, licenseForm);
      }
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
      try {
        if (btn.dataset.action === "edit") {
          await openArchiveForEdit(id);
        }
        if (btn.dataset.action === "view") {
          await openArchivePreview(id);
        }
        if (btn.dataset.action === "docx") {
          window.location.href = `/api/offers/${encodeURIComponent(id)}/docx`;
        }
        if (btn.dataset.action === "excel") {
          window.location.href = `/api/offers/${encodeURIComponent(id)}/excel`;
        }
        if (btn.dataset.action === "delete") {
          if (!confirm("Eintrag löschen?")) return;
          await api(`/api/offers/${encodeURIComponent(id)}`, { method: "DELETE" });
          await loadArchive();
        }
      } catch (err) {
        alert(err.message);
      }
    });

    document.getElementById("btnComposeOffer").addEventListener("click", () => composeOfferDocument({ save: true }));
    document.getElementById("btnSaveOfferDoc").addEventListener("click", () => composeOfferDocument({ save: true, notify: true }));
    document.getElementById("btnWordOfferDoc").addEventListener("click", downloadOfferDocx);
    document.getElementById("btnExcelOfferDoc").addEventListener("click", () => {
      const id = state.savedOfferId || state.offerDocument?.id || state.offerDocument?.meta?.offerNumber;
      if (!id) {
        alert("Bitte das Angebot zuerst speichern.");
        return;
      }
      window.location.href = `/api/offers/${encodeURIComponent(id)}/excel`;
    });
    document.getElementById("btnPrintOffer").addEventListener("click", () => {
      switchView("offer");
      window.print();
    });
  }

  async function init() {
    state.licenseCatalog = await api("/api/catalog");
    state.itCatalog = await api("/api/it/catalog");
    renderInstances();
    renderAddons();
    updateClientHints();
    const product = state.licenseCatalog.product || {};
    setFormValue(licenseForm, "licenseMarginPercent", product.licenseMarginPercent ?? 28);
    setFormValue(licenseForm, "eurToChfRate", product.eurToChfRate ?? 0.93);
    document.getElementById("sllHint").textContent = state.licenseCatalog.sllDefinition || "";
    document.getElementById("licenseDisclaimer").textContent = state.licenseCatalog.product.disclaimer;
    renderItOptions();
    renderItExtensions();
    const r = state.itCatalog.rates;
    setFormValue(itForm, "itMarginPercent", r.marginPercent ?? 28);
    document.getElementById("itRatesHint").textContent =
      `Stammdaten: Stundensatz ${money(r.hourlyRate, "CHF")} · km ${money(r.kmRate, "CHF")} · Verpflegung ${money(r.mealRate, "CHF")} · Übernachtung ${money(r.overnightRate, "CHF")} · Default-Marge ${r.marginPercent ?? 28}%`;
    document.getElementById("itDisclaimer").textContent = state.itCatalog.meta.disclaimer;
    initAddressAutocomplete(document.getElementById("customerAddress"));
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
