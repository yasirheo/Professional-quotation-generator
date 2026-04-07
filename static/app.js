const seed = window.SEED_DATA;
const STORAGE_KEYS = {
    library: "reliance_quote_library_v2",
    quote: "reliance_quote_draft_v2",
    theme: "reliance_quote_theme_v2",
    logo: "reliance_quote_logo_v2"
};

const state = { library: null, quote: null, ui: { activeLibraryCategoryId: "" } };
const els = {};

function clone(value) { return JSON.parse(JSON.stringify(value)); }
function id(prefix) { return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`; }
function slugify(value) { return (value || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "item"; }
function num(value) { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function money(value) { return num(value).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
function setLibraryMarkup(markup) {
    els.libraryTree.innerHTML = String(markup || "")
        .replaceAll("â€¢", "|");
}
function sortByLabel(values, pick = value => value?.name || "") {
    return [...(values || [])].sort((left, right) =>
        pick(left).localeCompare(pick(right), undefined, { numeric: true, sensitivity: "base" })
    );
}
function currentDateString() {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, "0");
    const month = now.toLocaleString("en-US", { month: "long" });
    const year = now.getFullYear();
    return `${day}-${month}-${year}`;
}

function toast(target, message, type = "info") {
    target.textContent = message;
    target.className = `status show ${type}`;
    clearTimeout(target._timer);
    target._timer = setTimeout(() => { target.className = "status"; }, 3200);
}

function getCategories() {
    return clone(state.library.categories).sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0) || a.name.localeCompare(b.name));
}

function getCategory(categoryId) {
    return state.library.categories.find(category => category.id === categoryId) || null;
}

function getSubcategory(categoryId, subcategoryId) {
    const category = getCategory(categoryId);
    return category?.subcategories.find(subcategory => subcategory.id === subcategoryId) || null;
}

function getItem(categoryId, subcategoryId, itemId) {
    return getSubcategory(categoryId, subcategoryId)?.items.find(item => item.id === itemId) || null;
}

function saveLibrary() {
    localStorage.setItem(STORAGE_KEYS.library, JSON.stringify(state.library));
}

function mergeLibrary(savedLibrary, seedLibrary) {
    if (!savedLibrary) return clone(seedLibrary);

    const merged = clone(savedLibrary);
    merged.company = { ...seedLibrary.company, ...(savedLibrary.company || {}) };

    const specSet = new Set([...(savedLibrary.specOptions || []), ...(seedLibrary.specOptions || [])]);
    merged.specOptions = Array.from(specSet);

    const categoryMap = new Map((merged.categories || []).map(category => [category.id, category]));
    (seedLibrary.categories || []).forEach(seedCategory => {
        let targetCategory = categoryMap.get(seedCategory.id);
        if (!targetCategory) {
            targetCategory = clone(seedCategory);
            merged.categories.push(targetCategory);
            categoryMap.set(seedCategory.id, targetCategory);
            return;
        }

        targetCategory.name = targetCategory.name || seedCategory.name;
        targetCategory.color = targetCategory.color || seedCategory.color;
        targetCategory.sortOrder = targetCategory.sortOrder || seedCategory.sortOrder;
        const subcategoryMap = new Map((targetCategory.subcategories || []).map(subcategory => [subcategory.id, subcategory]));

        (seedCategory.subcategories || []).forEach(seedSubcategory => {
            let targetSubcategory = subcategoryMap.get(seedSubcategory.id);
            if (!targetSubcategory) {
                targetCategory.subcategories.push(clone(seedSubcategory));
                return;
            }

            const existingKeys = new Set((targetSubcategory.items || []).map(item => `${(item.name || "").toLowerCase()}|${(item.description || "").toLowerCase()}|${item.defaultSpec || ""}`));
            (seedSubcategory.items || []).forEach(seedItem => {
                const key = `${(seedItem.name || "").toLowerCase()}|${(seedItem.description || "").toLowerCase()}|${seedItem.defaultSpec || ""}`;
                if (!existingKeys.has(key)) {
                    targetSubcategory.items.push(clone(seedItem));
                }
            });
        });
    });

    return normalizeLibrary(merged);
}

function dedupeItems(items) {
    const seen = new Set();
    return (items || []).filter(item => {
        const key = `${(item.name || "").toLowerCase()}|${(item.description || "").toLowerCase()}|${item.defaultSpec || ""}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function ensureSubcategory(category, subcategoryId, subcategoryName) {
    let subcategory = category.subcategories.find(entry => entry.id === subcategoryId);
    if (!subcategory) {
        subcategory = { id: subcategoryId, name: subcategoryName, items: [] };
        category.subcategories.push(subcategory);
    }
    return subcategory;
}

function resolveStructureSubcategory(item) {
    const text = `${item.name || ""} ${item.description || ""}`.toLowerCase();
    if (/epoxy|oxide|smoke|paint|checker/.test(text)) return "finishing-materials";
    if (/labor|civil|earthing|lightning|curve|concrete|installation/.test(text)) return "mechanical-civil";
    if (/plate|bolt|clamp|tz|nut/.test(text)) return "plates-anchors-fasteners";
    return "fabricated-structure-steel";
}

function resolveSolarSubcategory(item) {
    const text = `${item.name || ""} ${item.description || ""}`.toLowerCase();
    if (/transport|installation|earthing|earthling/.test(text)) return "earthing-transport-installation";
    if (/pipe|duct|tape|tie|flexible|sleeve|socket|saddle|screw/.test(text)) return "fittings-accessories";
    if (/breaker|db|distribution|change over|changeover|mc4/.test(text)) return "protection-switching";
    return "electrical-wiring";
}

function normalizeLibrary(library) {
    const elevated = (library.categories || []).find(category => category.id === "elevated-structure");
    if (elevated) {
        const legacy = elevated.subcategories.find(subcategory => subcategory.id === "fabricated-structure");
        if (legacy) {
            legacy.items.forEach(item => {
                const targetId = resolveStructureSubcategory(item);
                const target = ensureSubcategory(
                    elevated,
                    targetId,
                    {
                        "fabricated-structure-steel": "Fabricated Structure Steel",
                        "plates-anchors-fasteners": "Plates, Anchors & Fasteners",
                        "finishing-materials": "Finishing Materials",
                        "mechanical-civil": "Mechanical Installation & Civil",
                    }[targetId]
                );
                target.items.push(item);
            });
            elevated.subcategories = elevated.subcategories.filter(subcategory => subcategory.id !== "fabricated-structure");
        }
    }

    const solar = (library.categories || []).find(category => category.id === "solar-system");
    if (solar) {
        ["electrical", "transport-installation"].forEach(legacyId => {
            const legacy = solar.subcategories.find(subcategory => subcategory.id === legacyId);
            if (!legacy) return;
            legacy.items.forEach(item => {
                const targetId = resolveSolarSubcategory(item);
                const target = ensureSubcategory(
                    solar,
                    targetId,
                    {
                        "electrical-wiring": "Electrical Wiring",
                        "protection-switching": "Protection & Switching",
                        "fittings-accessories": "Fittings & Accessories",
                        "earthing-transport-installation": "Earthing, Transportation & Installation",
                    }[targetId]
                );
                target.items.push(item);
            });
            solar.subcategories = solar.subcategories.filter(subcategory => subcategory.id !== legacyId);
        });
    }

    (library.categories || []).forEach(category => {
        category.subcategories = (category.subcategories || []).map(subcategory => ({
            ...subcategory,
            items: dedupeItems(subcategory.items)
        }));
    });

    return library;
}

function readQuoteForm() {
    return {
        preparedBy: els.preparedBy.value.trim(),
        preparedPhone: els.preparedPhone.value.trim(),
        preparedEmail: els.preparedEmail.value.trim(),
        to: els.quoteTo.value.trim(),
        proposalFor: els.proposalFor.value.trim(),
        city: els.quoteCity.value.trim(),
        quoteDate: els.quoteDate.value.trim(),
        discountValue: num(els.discountInput.value),
        discountType: els.discountType.value,
        taxValue: num(els.taxInput.value),
        taxType: els.taxType.value,
        lines: state.quote.lines,
        selectedLineId: state.quote.selectedLineId
    };
}

function saveQuote() {
    state.quote = readQuoteForm();
    localStorage.setItem(STORAGE_KEYS.quote, JSON.stringify(state.quote));
}

function setTheme(theme) {
    document.body.dataset.theme = theme;
    localStorage.setItem(STORAGE_KEYS.theme, theme);
}

function loadState() {
    const savedLibrary = JSON.parse(localStorage.getItem(STORAGE_KEYS.library) || "null");
    const seedLibrary = {
        company: seed.company,
        specOptions: seed.specOptions,
        categories: seed.categories
    };
    state.library = mergeLibrary(savedLibrary, seedLibrary);

    state.quote = JSON.parse(localStorage.getItem(STORAGE_KEYS.quote) || "null") || {
        preparedBy: seed.defaultQuote.preparedBy,
        preparedPhone: seed.defaultQuote.preparedPhone,
        preparedEmail: seed.defaultQuote.preparedEmail,
        to: seed.defaultQuote.to,
        proposalFor: seed.defaultQuote.proposalFor,
        city: seed.defaultQuote.city,
        quoteDate: seed.defaultQuote.quoteDate,
        discountValue: 0,
        discountType: "amount",
        taxValue: 0,
        taxType: "amount",
        lines: [],
        selectedLineId: null
    };

    if (
        state.quote.to === "Rizwan Bhai" &&
        state.quote.city === "Karachi" &&
        state.quote.proposalFor === "Proposal for 10 kw hybrid solar system structure and 6 kw inverter"
    ) {
        state.quote.to = "";
        state.quote.city = "";
        state.quote.proposalFor = "";
    }

    setTheme(localStorage.getItem(STORAGE_KEYS.theme) || "light");
}

function cache() {
    [
        "libraryTree", "libraryTabs", "libraryStatus", "builderStatus", "librarySearch", "categorySelect", "subcategorySelect", "itemSelect",
        "itemSearchInput", "itemQuickList", "descriptionInput", "specInput", "specOptionsList", "quantityInput", "unitPriceInput", "saveToLibraryCheckbox", "quoteTableBody",
        "subtotalCard", "discountCard", "taxCard", "grandTotalCard", "preparedBy", "preparedPhone", "preparedEmail",
        "quoteDate", "quoteTo", "quoteCity", "proposalFor", "discountInput", "discountType", "taxInput", "taxType",
        "heroQuoteTo", "heroQuoteProposal", "heroQuoteDate", "printBodyInfo", "printDateBox", "printTableBody",
        "printSubtotal", "printDiscount", "printTax", "printGrandTotal", "printTerms", "printClosing",
        "printCompanyName", "printCompanyLine1", "printCompanyLine2", "printCompanyLine3", "printFooterLine1",
        "printFooterLine2", "printLogo", "logoFallback", "subcategoryCategorySelect", "newCategoryName", "newCategoryColor",
        "newSubcategoryName", "newSpecValue", "updateLineButton", "duplicateLineButton", "deleteLineButton", "logoInput",
        "importFileInput", "itemDialogCategorySelect", "itemDialogSubcategorySelect", "itemDialogName", "itemDialogDescription",
        "itemDialogSpec", "itemDialogPrice", "editingItemId", "itemDialogTitle"
    ].forEach(name => { els[name] = document.getElementById(name); });
}

function hydrateQuoteMeta() {
    els.preparedBy.value = state.quote.preparedBy || "";
    els.preparedPhone.value = state.quote.preparedPhone || "";
    els.preparedEmail.value = state.quote.preparedEmail || "";
    els.quoteDate.value = currentDateString();
    els.quoteTo.value = state.quote.to || "";
    els.quoteCity.value = state.quote.city || "";
    els.proposalFor.value = state.quote.proposalFor || "";
    els.discountInput.value = state.quote.discountValue || 0;
    els.discountType.value = state.quote.discountType || "amount";
    els.taxInput.value = state.quote.taxValue || 0;
    els.taxType.value = state.quote.taxType || "amount";
    state.quote.quoteDate = els.quoteDate.value;
}

function ensureActiveLibraryCategory() {
    const categories = getCategories();
    if (!categories.length) {
        state.ui.activeLibraryCategoryId = "";
        return;
    }
    if (!state.ui.activeLibraryCategoryId || !categories.some(category => category.id === state.ui.activeLibraryCategoryId)) {
        state.ui.activeLibraryCategoryId = categories[0].id;
    }
}

function renderLibraryTabs() {
    ensureActiveLibraryCategory();
    const categories = getCategories();
    els.libraryTabs.innerHTML = categories.map(category => `
        <button type="button" class="library-tab ${state.ui.activeLibraryCategoryId === category.id ? "active" : ""}" data-action="filter-library" data-category-id="${category.id}">
            ${category.name}
        </button>
    `).join("") || `<span class="muted">No categories yet</span>`;
}

function renderCategoryOptions() {
    const categories = getCategories();
    const selected = els.categorySelect.value || categories[0]?.id || "";
    const options = categories.map(category => `<option value="${category.id}">${category.name}</option>`).join("");
    els.categorySelect.innerHTML = options || `<option value="">No categories</option>`;
    els.subcategoryCategorySelect.innerHTML = options || `<option value="">No categories</option>`;
    if (categories.some(category => category.id === selected)) {
        els.categorySelect.value = selected;
        els.subcategoryCategorySelect.value = selected;
    }
    ensureActiveLibraryCategory();
    renderLibraryTabs();
    renderSubcategoryOptions();
}

function renderSubcategoryOptions() {
    const subcategories = getCategory(els.categorySelect.value)?.subcategories || [];
    const previous = els.subcategorySelect.value;
    els.subcategorySelect.innerHTML = subcategories.length
        ? subcategories.map(subcategory => `<option value="${subcategory.id}">${subcategory.name}</option>`).join("")
        : `<option value="">No subcategories</option>`;
    if (subcategories.some(subcategory => subcategory.id === previous)) {
        els.subcategorySelect.value = previous;
    }
    renderItemOptions();
}

function renderItemOptions() {
    const items = getSubcategory(els.categorySelect.value, els.subcategorySelect.value)?.items || [];
    const previous = els.itemSelect.value;
    els.itemSelect.innerHTML = `<option value="">Custom description</option>` + items.map(item => `<option value="${item.id}">${item.name}</option>`).join("");
    els.itemSelect.value = items.some(item => item.id === previous) ? previous : "";
    if (!els.itemSelect.value) els.itemSearchInput.value = "";
    renderItemQuickList();
}

function renderSpecOptions() {
    const current = els.specInput.value;
    const specs = clone(state.library.specOptions).sort((a, b) => a.localeCompare(b));
    els.specOptionsList.innerHTML = specs.map(spec => `<option value="${spec}"></option>`).join("");
    if (current) els.specInput.value = current;
}

function renderItemQuickList() {
    const items = getSubcategory(els.categorySelect.value, els.subcategorySelect.value)?.items || [];
    const query = els.itemSearchInput.value.trim().toLowerCase();
    const filtered = items.filter(item => {
        const haystack = `${item.name} ${item.description || ""} ${item.defaultSpec || ""}`.toLowerCase();
        return !query || haystack.includes(query);
    }).slice(0, 10);

    els.itemQuickList.innerHTML = filtered.length
        ? filtered.map(item => `<button type="button" class="picker-item ${els.itemSelect.value === item.id ? "active" : ""}" data-action="pick-item" data-item-id="${item.id}">${item.name}${item.defaultSpec ? ` | ${item.defaultSpec}` : ""}</button>`).join("")
        : `<span class="muted">No saved items for this subcategory.</span>`;
}

function pickItem(itemId) {
    els.itemSelect.value = itemId;
    const item = getItem(els.categorySelect.value, els.subcategorySelect.value, itemId);
    els.itemSearchInput.value = item?.name || "";
    loadItemToForm();
    renderItemQuickList();
}

function summary() {
    const subtotal = state.quote.lines.reduce((sum, line) => sum + num(line.lineTotal), 0);
    const discountValue = num(els.discountInput.value);
    const discountAmount = els.discountType.value === "percent" ? subtotal * discountValue / 100 : discountValue;
    const base = Math.max(subtotal - discountAmount, 0);
    const taxValue = num(els.taxInput.value);
    const taxAmount = els.taxType.value === "percent" ? base * taxValue / 100 : taxValue;
    return { subtotal, discountAmount, taxAmount, grandTotal: base + taxAmount };
}

function groupLines() {
    const groups = new Map();
    state.quote.lines.forEach(line => {
        if (!groups.has(line.category)) groups.set(line.category, []);
        groups.get(line.category).push(line);
    });
    return Array.from(groups.entries());
}

function normalizeLinesByCategoryBlocks() {
    const grouped = new Map();
    const order = [];

    state.quote.lines.forEach(line => {
        const key = line.categoryId || line.category || "uncategorized";
        if (!grouped.has(key)) {
            grouped.set(key, []);
            order.push(key);
        }
        grouped.get(key).push(line);
    });

    state.quote.lines = order.flatMap(key => grouped.get(key));
}

function syncSerials() {
    normalizeLinesByCategoryBlocks();
    state.quote.lines = state.quote.lines.map((line, index) => ({ ...line, serialNo: index + 1 }));
}

function renderTable() {
    const selectedId = state.quote.selectedLineId;
    const rows = [];
    groupLines().forEach(([categoryName, items]) => {
        rows.push(`<tr class="group-row"><td colspan="9">${categoryName}</td></tr>`);
        items.forEach(line => {
            rows.push(`
                <tr data-line-id="${line.id}" ${selectedId === line.id ? 'style="outline:2px solid rgba(138,100,36,.35);"' : ""}>
                    <td class="center">${line.serialNo}</td>
                    <td>${line.category}</td>
                    <td>${line.subcategory}</td>
                    <td class="desc">${line.description}</td>
                    <td class="center">${line.spec}</td>
                    <td class="center">${line.quantity}</td>
                    <td class="num">${money(line.unitPrice)}</td>
                    <td class="num">${money(line.lineTotal)}</td>
                    <td><div class="row-tools">
                        <button class="btn-link" data-action="edit" data-line-id="${line.id}">Edit</button>
                        <button class="btn-link" data-action="copy" data-line-id="${line.id}">Copy</button>
                        <button class="btn-link" data-action="delete" data-line-id="${line.id}">Delete</button>
                    </div></td>
                </tr>
            `);
        });
        rows.push(`<tr class="subtotal-row"><td colspan="7" class="num">${categoryName} Subtotal</td><td class="num">${money(items.reduce((sum, item) => sum + num(item.lineTotal), 0))}</td><td></td></tr>`);
    });
    els.quoteTableBody.innerHTML = rows.join("") || `<tr><td colspan="9" class="center muted">No line items yet. Add the first row from the builder form.</td></tr>`;

    const calc = summary();
    els.subtotalCard.textContent = money(calc.subtotal);
    els.discountCard.textContent = money(calc.discountAmount);
    els.taxCard.textContent = money(calc.taxAmount);
    els.grandTotalCard.textContent = money(calc.grandTotal);
    els.updateLineButton.disabled = !state.quote.selectedLineId;
    els.duplicateLineButton.disabled = !state.quote.selectedLineId;
    els.deleteLineButton.disabled = !state.quote.selectedLineId;
}

function renderPreview() {
    const company = state.library.company;
    const calc = summary();
    els.heroQuoteTo.textContent = `To: ${els.quoteTo.value.trim() || "Customer Name"}`;
    els.heroQuoteProposal.textContent = els.proposalFor.value.trim() || "Proposal details will appear here";
    els.heroQuoteDate.textContent = els.quoteDate.value.trim() || currentDateString();
    els.printCompanyName.textContent = company.name;
    els.printCompanyLine1.textContent = company.tagline;
    els.printCompanyLine2.textContent = els.preparedBy.value.trim() || company.contact_person;
    els.printCompanyLine3.textContent = `${els.preparedPhone.value.trim() || company.phone} | ${els.preparedEmail.value.trim() || company.email}`;
    els.printFooterLine1.textContent = company.footer_line_1;
    els.printFooterLine2.textContent = company.footer_line_2;

    els.printBodyInfo.innerHTML = `
        <div>${company.tagline}</div>
        <div>${els.preparedBy.value.trim() || company.contact_person}</div>
        <div>${els.preparedPhone.value.trim() || company.phone}</div>
        <div>${els.preparedEmail.value.trim() || company.email}</div>
        <br>
        <div><strong>TO</strong>       ${els.quoteTo.value.trim() || "Customer Name"}</div>
        <div>${els.proposalFor.value.trim() || "Proposal details"}</div>
        <div>${els.quoteCity.value.trim() || ""}</div>
    `;
    els.printDateBox.innerHTML = `<strong>${els.quoteDate.value.trim() || currentDateString()}</strong>`;

    const previewRows = [];
    groupLines().forEach(([categoryName, items]) => {
        previewRows.push(`<tr class="print-group"><td colspan="8">${categoryName}</td></tr>`);
        items.forEach(line => {
            previewRows.push(`
                <tr>
                    <td class="center">${line.serialNo}</td>
                    <td>${line.category}</td>
                    <td>${line.subcategory}</td>
                    <td class="desc">${line.description}</td>
                    <td class="center">${line.spec}</td>
                    <td class="center">${line.quantity}</td>
                    <td class="num">${money(line.unitPrice)}</td>
                    <td class="num">${money(line.lineTotal)}</td>
                </tr>
            `);
        });
        previewRows.push(`<tr class="print-subtotal"><td colspan="7" class="num">${categoryName} Subtotal</td><td class="num">${money(items.reduce((sum, item) => sum + num(item.lineTotal), 0))}</td></tr>`);
    });
    els.printTableBody.innerHTML = previewRows.join("") || `<tr><td colspan="8" class="center">No items added yet.</td></tr>`;
    els.printSubtotal.textContent = money(calc.subtotal);
    els.printDiscount.textContent = money(calc.discountAmount);
    els.printTax.textContent = money(calc.taxAmount);
    els.printGrandTotal.textContent = money(calc.grandTotal);
    els.printTerms.innerHTML = company.terms.map((term, index) => `<div>${index + 1}- ${term}</div>`).join("");
    els.printClosing.innerHTML = company.closing.map(line => `<div>${line}</div>`).join("");

    const savedLogo = localStorage.getItem(STORAGE_KEYS.logo);
    if (savedLogo) {
        els.printLogo.src = savedLogo;
        els.printLogo.classList.remove("hidden");
        els.logoFallback.classList.add("hidden");
    } else {
        els.printLogo.classList.add("hidden");
        els.logoFallback.classList.remove("hidden");
    }
}

function clearEntry() {
    const keepCategory = els.categorySelect.value;
    const keepSubcategory = els.subcategorySelect.value;
    const keepSpec = els.specInput.value;
    els.itemSelect.value = "";
    els.itemSearchInput.value = "";
    els.descriptionInput.value = "";
    els.quantityInput.value = "1";
    els.unitPriceInput.value = "";
    els.saveToLibraryCheckbox.value = "no";
    state.quote.selectedLineId = null;
    renderCategoryOptions();
    els.categorySelect.value = keepCategory;
    renderSubcategoryOptions();
    if (keepSubcategory) els.subcategorySelect.value = keepSubcategory;
    renderItemOptions();
    renderSpecOptions();
    if (keepSpec) els.specInput.value = keepSpec;
    renderTable();
    saveQuote();
}

function saveCurrentItem({ silent = false } = {}) {
    const category = getCategory(els.categorySelect.value);
    const subcategory = getSubcategory(els.categorySelect.value, els.subcategorySelect.value);
    const description = els.descriptionInput.value.trim();
    if (!category || !subcategory || !description) {
        if (!silent) toast(els.builderStatus, "Choose category, subcategory, and description before saving a custom item.", "error");
        return false;
    }
    const spec = els.specInput.value.trim();
    const unitPrice = num(els.unitPriceInput.value);
    const exists = subcategory.items.some(item => item.description === description && (item.defaultSpec || "") === spec);
    if (exists) {
        if (!silent) toast(els.builderStatus, "This custom item already exists in the saved library.", "info");
        return true;
    }
    subcategory.items.push({
        id: id("item"),
        name: description.length > 46 ? `${description.slice(0, 46)}...` : description,
        description,
        defaultSpec: spec,
        defaultUnitPrice: unitPrice || ""
    });
    saveLibrary();
    renderItemOptions();
    renderLibraryTree();
    if (!silent) toast(els.builderStatus, "Custom item saved for future use.", "success");
    return true;
}

function lineFromForm() {
    const category = getCategory(els.categorySelect.value);
    const subcategory = getSubcategory(els.categorySelect.value, els.subcategorySelect.value);
    const description = els.descriptionInput.value.trim();
    const quantity = num(els.quantityInput.value);
    const unitPrice = num(els.unitPriceInput.value);
    if (!category) throw new Error("Select a main category.");
    if (!subcategory) throw new Error("Select a subcategory.");
    if (!description) throw new Error("Enter an item description.");
    if (quantity <= 0) throw new Error("Quantity must be greater than 0.");
    if (unitPrice < 0) throw new Error("Unit price must be numeric.");
    return {
        id: state.quote.selectedLineId || id("line"),
        serialNo: 0,
        categoryId: category.id,
        category: category.name,
        subcategoryId: subcategory.id,
        subcategory: subcategory.name,
        itemId: els.itemSelect.value || "",
        description,
        spec: els.specInput.value.trim(),
        quantity,
        unitPrice,
        lineTotal: quantity * unitPrice
    };
}

function addOrUpdateLine(mode) {
    try {
        const line = lineFromForm();
        if (els.saveToLibraryCheckbox.value === "yes") saveCurrentItem({ silent: true });
        if (mode === "update" && state.quote.selectedLineId) {
            const index = state.quote.lines.findIndex(item => item.id === state.quote.selectedLineId);
            if (index === -1) throw new Error("Selected row no longer exists.");
            state.quote.lines[index] = line;
            toast(els.builderStatus, "Line item updated.", "success");
        } else {
            state.quote.lines.push(line);
            toast(els.builderStatus, "Line item added.", "success");
        }
        syncSerials();
        clearEntry();
        renderPreview();
    } catch (error) {
        toast(els.builderStatus, error.message, "error");
    }
}

function loadItemToForm() {
    const item = getItem(els.categorySelect.value, els.subcategorySelect.value, els.itemSelect.value);
    if (!item) return;
    els.itemSearchInput.value = item.name || "";
    els.descriptionInput.value = item.description || item.name || "";
    if (item.defaultSpec) {
        els.specInput.value = item.defaultSpec;
    }
    if (item.defaultUnitPrice !== "" && item.defaultUnitPrice !== null && item.defaultUnitPrice !== undefined) {
        els.unitPriceInput.value = item.defaultUnitPrice;
    }
    renderItemQuickList();
}

function editLine(lineId) {
    const line = state.quote.lines.find(item => item.id === lineId);
    if (!line) return;
    state.quote.selectedLineId = line.id;
    els.categorySelect.value = line.categoryId;
    renderSubcategoryOptions();
    els.subcategorySelect.value = line.subcategoryId;
    renderItemOptions();
    els.itemSelect.value = line.itemId || "";
    const loadedItem = getItem(line.categoryId, line.subcategoryId, line.itemId || "");
    els.itemSearchInput.value = loadedItem?.name || "";
    els.descriptionInput.value = line.description;
    if (!state.library.specOptions.includes(line.spec)) {
        state.library.specOptions.push(line.spec);
        saveLibrary();
        renderSpecOptions();
    }
    els.specInput.value = line.spec;
    els.quantityInput.value = line.quantity;
    els.unitPriceInput.value = line.unitPrice;
    renderItemQuickList();
    renderTable();
    toast(els.builderStatus, `Editing row ${line.serialNo}. Update it when ready.`, "info");
}

function deleteLine(lineId) {
    state.quote.lines = state.quote.lines.filter(item => item.id !== lineId);
    if (state.quote.selectedLineId === lineId) state.quote.selectedLineId = null;
    syncSerials();
    renderTable();
    renderPreview();
    saveQuote();
    toast(els.builderStatus, "Line item deleted.", "success");
}

function duplicateLine(lineId) {
    const source = state.quote.lines.find(item => item.id === lineId);
    if (!source) return;
    state.quote.lines.push({ ...source, id: id("line"), serialNo: 0 });
    syncSerials();
    renderTable();
    renderPreview();
    saveQuote();
    toast(els.builderStatus, "Line item duplicated.", "success");
}

function renderLibraryTree() {
    const query = els.librarySearch.value.trim().toLowerCase();
    ensureActiveLibraryCategory();
    const category = getCategories().find(entry => entry.id === state.ui.activeLibraryCategoryId);

    if (!category) {
        setLibraryMarkup(`<div class="tree-card"><div class="library-empty">No categories are available yet. Add a main category to start building your library.</div></div>`);
        return;
    }

    const categoryMatches = !query || category.name.toLowerCase().includes(query);
    const visibleSubcategories = sortByLabel(category.subcategories).map(subcategory => {
        const allItems = sortByLabel(subcategory.items, item => `${item.description || item.name || ""} ${item.defaultSpec || ""}`);
        const subcategoryMatches = !query || subcategory.name.toLowerCase().includes(query);
        const matchedItems = allItems.filter(item => `${item.name || ""} ${item.description || ""} ${item.defaultSpec || ""} ${item.defaultUnitPrice || ""}`.toLowerCase().includes(query));
        const visibleItems = (!query || categoryMatches || subcategoryMatches) ? allItems : matchedItems;

        if (query && !categoryMatches && !subcategoryMatches && !matchedItems.length) return null;

        return {
            ...subcategory,
            allItems,
            visibleItems
        };
    }).filter(Boolean);

    const totalItems = category.subcategories.reduce((sum, subcategory) => sum + (subcategory.items?.length || 0), 0);
    const shownItems = visibleSubcategories.reduce((sum, subcategory) => sum + subcategory.visibleItems.length, 0);

    if (!visibleSubcategories.length) {
        setLibraryMarkup(`
            <div class="tree-card">
                <div class="library-category-head">
                    <div class="library-title-block">
                        <strong>${escapeHtml(category.name)}</strong>
                        <div class="tree-meta">${category.subcategories.length} subcategories | ${totalItems} items</div>
                    </div>
                    <div class="library-section-actions">
                        <button type="button" class="btn-link" data-action="rename-category" data-category-id="${category.id}">Rename</button>
                        <button type="button" class="btn-link" data-action="delete-category" data-category-id="${category.id}">Delete</button>
                    </div>
                </div>
                <div class="library-empty">No matching items were found in this category. Try a different search or clear the search box.</div>
            </div>
        `);
        return;
    }

    const sections = visibleSubcategories.map(subcategory => {
        const countLabel = query && subcategory.visibleItems.length !== subcategory.allItems.length
            ? `${subcategory.visibleItems.length} of ${subcategory.allItems.length} items shown`
            : `${subcategory.allItems.length} items`;

        const itemsMarkup = subcategory.visibleItems.length
            ? subcategory.visibleItems.map(item => {
                const displayTitle = item.description || item.name || "Saved item";
                const helperLabel = item.name && item.description && item.name !== item.description
                    ? `<div class="library-entry-note">Saved name: ${escapeHtml(item.name)}</div>`
                    : "";
                const specLabel = item.defaultSpec
                    ? `<span class="library-meta-pill">Spec / Unit: ${escapeHtml(item.defaultSpec)}</span>`
                    : `<span class="library-meta-pill">Spec / Unit: custom when used</span>`;
                const priceLabel = item.defaultUnitPrice !== "" && item.defaultUnitPrice !== null && item.defaultUnitPrice !== undefined
                    ? `<span class="library-meta-pill">Default price: ${escapeHtml(money(item.defaultUnitPrice))}</span>`
                    : "";

                return `
                    <article class="library-entry">
                        <strong>${escapeHtml(displayTitle)}</strong>
                        ${helperLabel}
                        <div class="library-entry-meta">
                            ${specLabel}
                            ${priceLabel}
                        </div>
                        <div class="library-entry-actions">
                            <button type="button" class="btn-secondary" data-action="use-item" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}" data-item-id="${item.id}">Use</button>
                            <button type="button" class="btn-secondary" data-action="rename-item" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}" data-item-id="${item.id}">Edit</button>
                            <button type="button" class="btn-danger" data-action="delete-item" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}" data-item-id="${item.id}">Delete</button>
                        </div>
                    </article>
                `;
            }).join("")
            : `<div class="library-empty">No saved items are in this subcategory yet. Use Add Item to save one here.</div>`;

        return `
            <section class="library-section">
                <div class="library-section-head">
                    <div class="library-title-block">
                        <strong>${escapeHtml(subcategory.name)}</strong>
                        <div class="tree-meta">${countLabel}</div>
                    </div>
                    <div class="library-section-actions">
                        <button type="button" class="btn-link" data-action="add-item-here" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}">Add Item</button>
                        <button type="button" class="btn-link" data-action="rename-subcategory" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}">Rename</button>
                        <button type="button" class="btn-link" data-action="delete-subcategory" data-category-id="${category.id}" data-subcategory-id="${subcategory.id}">Delete</button>
                    </div>
                </div>
                <div class="library-entry-list">
                    ${itemsMarkup}
                </div>
            </section>
        `;
    }).join("");

    setLibraryMarkup(`
        <div class="tree-card">
            <div class="library-category-head">
                <div class="library-title-block">
                    <strong>${escapeHtml(category.name)}</strong>
                    <div class="tree-meta">${visibleSubcategories.length} subcategories | ${query ? `${shownItems} matching items` : `${totalItems} items`}</div>
                </div>
                <div class="library-section-actions">
                    <button type="button" class="btn-link" data-action="rename-category" data-category-id="${category.id}">Rename</button>
                    <button type="button" class="btn-link" data-action="delete-category" data-category-id="${category.id}">Delete</button>
                </div>
            </div>
            <div class="library-body">${sections}</div>
        </div>
    `);
}

function addCategory(name, color) {
    const trimmed = name.trim();
    if (!trimmed) return toast(els.libraryStatus, "Category name is required.", "error");
    if (state.library.categories.some(category => category.name.toLowerCase() === trimmed.toLowerCase())) {
        return toast(els.libraryStatus, "This category already exists.", "error");
    }
    state.library.categories.push({
        id: slugify(trimmed),
        name: trimmed,
        color: color.trim() || "#5b6770",
        sortOrder: state.library.categories.length + 1,
        subcategories: []
    });
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    toast(els.libraryStatus, "Main category added.", "success");
}

function addSubcategory(categoryId, name) {
    const category = getCategory(categoryId);
    const trimmed = name.trim();
    if (!category || !trimmed) return toast(els.libraryStatus, "Choose a category and enter a subcategory name.", "error");
    if (category.subcategories.some(subcategory => subcategory.name.toLowerCase() === trimmed.toLowerCase())) {
        return toast(els.libraryStatus, "This subcategory already exists.", "error");
    }
    category.subcategories.push({ id: slugify(trimmed), name: trimmed, items: [] });
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    toast(els.libraryStatus, "Subcategory added.", "success");
}

function addSpec(value) {
    const trimmed = value.trim();
    if (!trimmed) return toast(els.libraryStatus, "Spec / unit value is required.", "error");
    if (state.library.specOptions.includes(trimmed)) return toast(els.libraryStatus, "That spec already exists.", "info");
    state.library.specOptions.push(trimmed);
    saveLibrary();
    renderSpecOptions();
    toast(els.libraryStatus, "Spec option added.", "success");
}

function renameCategory(categoryId) {
    const category = getCategory(categoryId);
    if (!category) return;
    const next = window.prompt("Enter the new main category name:", category.name);
    if (!next || !next.trim()) return;
    category.name = next.trim();
    state.quote.lines = state.quote.lines.map(line => line.categoryId === categoryId ? { ...line, category: category.name } : line);
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    renderTable();
    renderPreview();
    saveQuote();
}

function renameSubcategory(categoryId, subcategoryId) {
    const subcategory = getSubcategory(categoryId, subcategoryId);
    if (!subcategory) return;
    const next = window.prompt("Enter the new subcategory name:", subcategory.name);
    if (!next || !next.trim()) return;
    subcategory.name = next.trim();
    state.quote.lines = state.quote.lines.map(line => line.subcategoryId === subcategoryId ? { ...line, subcategory: subcategory.name } : line);
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    renderTable();
    renderPreview();
    saveQuote();
}

function removeCategory(categoryId) {
    if (state.quote.lines.some(line => line.categoryId === categoryId)) {
        return toast(els.libraryStatus, "Delete related quote rows first, then remove this category.", "error");
    }
    state.library.categories = state.library.categories.filter(category => category.id !== categoryId);
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    toast(els.libraryStatus, "Category deleted.", "success");
}

function removeSubcategory(categoryId, subcategoryId) {
    if (state.quote.lines.some(line => line.categoryId === categoryId && line.subcategoryId === subcategoryId)) {
        return toast(els.libraryStatus, "Delete related quote rows first, then remove this subcategory.", "error");
    }
    const category = getCategory(categoryId);
    if (!category) return;
    category.subcategories = category.subcategories.filter(subcategory => subcategory.id !== subcategoryId);
    saveLibrary();
    renderCategoryOptions();
    renderLibraryTree();
    toast(els.libraryStatus, "Subcategory deleted.", "success");
}

function removeItem(categoryId, subcategoryId, itemId) {
    if (state.quote.lines.some(line => line.categoryId === categoryId && line.subcategoryId === subcategoryId && line.itemId === itemId)) {
        return toast(els.libraryStatus, "Delete related quote rows first, then remove this item.", "error");
    }
    const subcategory = getSubcategory(categoryId, subcategoryId);
    if (!subcategory) return;
    subcategory.items = subcategory.items.filter(item => item.id !== itemId);
    saveLibrary();
    renderItemOptions();
    renderLibraryTree();
    toast(els.libraryStatus, "Saved item deleted.", "success");
}

function renameItem(categoryId, subcategoryId, itemId) {
    const item = getItem(categoryId, subcategoryId, itemId);
    if (!item) return;
    openItemDialog({ categoryId, subcategoryId, item });
}

function useLibraryItem(categoryId, subcategoryId, itemId) {
    state.ui.activeLibraryCategoryId = categoryId;
    renderLibraryTabs();
    els.categorySelect.value = categoryId;
    renderSubcategoryOptions();
    els.subcategorySelect.value = subcategoryId;
    renderItemOptions();
    pickItem(itemId);
    toast(els.libraryStatus, "Item loaded into the quote form.", "info");
}

function syncItemDialogSubcategories() {
    const category = getCategory(els.itemDialogCategorySelect.value);
    const subcategories = category?.subcategories || [];
    const previous = els.itemDialogSubcategorySelect.value;
    els.itemDialogSubcategorySelect.innerHTML = subcategories.length
        ? subcategories.map(subcategory => `<option value="${subcategory.id}">${subcategory.name}</option>`).join("")
        : `<option value="">No subcategories</option>`;
    if (subcategories.some(subcategory => subcategory.id === previous)) {
        els.itemDialogSubcategorySelect.value = previous;
    }
}

function openItemDialog({ categoryId = "", subcategoryId = "", item = null } = {}) {
    const categories = getCategories();
    els.itemDialogCategorySelect.innerHTML = categories.map(category => `<option value="${category.id}">${category.name}</option>`).join("");
    els.itemDialogCategorySelect.value = categoryId || els.categorySelect.value || categories[0]?.id || "";
    syncItemDialogSubcategories();
    if (subcategoryId) els.itemDialogSubcategorySelect.value = subcategoryId;

    if (item) {
        els.itemDialogTitle.textContent = "Edit Library Item";
        els.editingItemId.value = item.id;
        els.itemDialogName.value = item.name || "";
        els.itemDialogDescription.value = item.description || "";
        els.itemDialogSpec.value = item.defaultSpec || "";
        els.itemDialogPrice.value = item.defaultUnitPrice || "";
    } else {
        els.itemDialogTitle.textContent = "Add Library Item";
        els.editingItemId.value = "";
        els.itemDialogName.value = "";
        els.itemDialogDescription.value = "";
        els.itemDialogSpec.value = "";
        els.itemDialogPrice.value = "";
    }

    document.getElementById("itemDialog").showModal();
}

function saveItemFromDialog() {
    const category = getCategory(els.itemDialogCategorySelect.value);
    const subcategory = getSubcategory(els.itemDialogCategorySelect.value, els.itemDialogSubcategorySelect.value);
    const name = els.itemDialogName.value.trim();
    const description = els.itemDialogDescription.value.trim() || name;
    const defaultSpec = els.itemDialogSpec.value.trim();
    const defaultUnitPrice = els.itemDialogPrice.value.trim();

    if (!category || !subcategory || !name) {
        return toast(els.libraryStatus, "Choose category, subcategory, and item name.", "error");
    }

    const editingId = els.editingItemId.value;
    const duplicate = subcategory.items.find(item => item.id !== editingId && item.name.toLowerCase() === name.toLowerCase() && (item.defaultSpec || "") === defaultSpec);
    if (duplicate) {
        return toast(els.libraryStatus, "A similar item already exists in this subcategory.", "error");
    }

    if (defaultSpec && !state.library.specOptions.includes(defaultSpec)) {
        state.library.specOptions.push(defaultSpec);
    }

    if (editingId) {
        const existing = getItem(category.id, subcategory.id, editingId);
        if (!existing) return toast(els.libraryStatus, "The item you tried to edit was not found.", "error");
        existing.name = name;
        existing.description = description;
        existing.defaultSpec = defaultSpec;
        existing.defaultUnitPrice = defaultUnitPrice;
    } else {
        subcategory.items.push({
            id: id("item"),
            name,
            description,
            defaultSpec,
            defaultUnitPrice,
        });
    }

    saveLibrary();
    renderCategoryOptions();
    renderSpecOptions();
    renderLibraryTree();
    document.getElementById("itemDialog").close();
    toast(els.libraryStatus, editingId ? "Library item updated." : "Library item added.", "success");
}

function findCategoryByName(name) {
    return state.library.categories.find(category => category.name.toLowerCase() === (name || "").toLowerCase()) || null;
}

function findSubcategoryByName(category, name) {
    return category?.subcategories.find(subcategory => subcategory.name.toLowerCase() === (name || "").toLowerCase()) || null;
}

function ensureLibraryEntryFromLine(line) {
    const categoryName = (line.category || "GENERAL").trim() || "GENERAL";
    const subcategoryName = (line.subcategory || "General Items").trim() || "General Items";
    const description = (line.description || "").trim();
    const spec = (line.spec || "").trim();

    let category = findCategoryByName(categoryName);
    if (!category) {
        category = {
            id: slugify(categoryName),
            name: categoryName,
            color: "#5b6770",
            sortOrder: state.library.categories.length + 1,
            subcategories: []
        };
        state.library.categories.push(category);
    }

    let subcategory = findSubcategoryByName(category, subcategoryName);
    if (!subcategory) {
        subcategory = { id: slugify(subcategoryName), name: subcategoryName, items: [] };
        category.subcategories.push(subcategory);
    }

    let itemId = "";
    if (description) {
        const existing = subcategory.items.find(item => item.description === description && (item.defaultSpec || "") === spec);
        if (existing) {
            itemId = existing.id;
        } else {
            const newItem = {
                id: id("item"),
                name: description.length > 46 ? `${description.slice(0, 46)}...` : description,
                description,
                defaultSpec: spec,
                defaultUnitPrice: line.unitPrice || ""
            };
            subcategory.items.push(newItem);
            itemId = newItem.id;
        }
    }

    if (spec && !state.library.specOptions.includes(spec)) {
        state.library.specOptions.push(spec);
    }

    return { category, subcategory, itemId };
}

function applyImportedPayload(payload, sourceLabel) {
    const importedQuote = payload.quote || {};
    const importedLines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!importedLines.length) {
        return toast(els.builderStatus, `No editable line items were found in the ${sourceLabel} file.`, "error");
    }

    const normalizedLines = importedLines.map((line, index) => {
        const quantity = num(line.quantity) || 1;
        const unitPrice = num(line.unitPrice);
        const path = ensureLibraryEntryFromLine(line);
        return {
            id: id("line"),
            serialNo: index + 1,
            categoryId: path.category.id,
            category: path.category.name,
            subcategoryId: path.subcategory.id,
            subcategory: path.subcategory.name,
            itemId: path.itemId,
            description: (line.description || "").trim(),
            spec: (line.spec || "").trim(),
            quantity,
            unitPrice,
            lineTotal: quantity * unitPrice
        };
    });

    els.preparedBy.value = importedQuote.preparedBy || els.preparedBy.value;
    els.preparedPhone.value = importedQuote.preparedPhone || els.preparedPhone.value;
    els.preparedEmail.value = importedQuote.preparedEmail || els.preparedEmail.value;
    els.quoteTo.value = importedQuote.to || els.quoteTo.value;
    els.proposalFor.value = importedQuote.proposalFor || els.proposalFor.value;
    els.quoteCity.value = importedQuote.city || els.quoteCity.value;
    els.quoteDate.value = importedQuote.quoteDate || currentDateString();
    els.discountInput.value = 0;
    els.taxInput.value = 0;

    state.quote.lines = normalizedLines;
    state.quote.selectedLineId = null;
    syncSerials();
    saveLibrary();
    saveQuote();
    renderCategoryOptions();
    renderSpecOptions();
    renderLibraryTree();
    renderTable();
    renderPreview();
    toast(els.builderStatus, `${sourceLabel} file imported for re-editing.`, "success");
}

async function importQuoteFile(file) {
    if (!file) return;
    if (state.quote.lines.length && !window.confirm("Replace the current quote with the imported file?")) return;

    const formData = new FormData();
    formData.append("file", file);
    formData.append("library", JSON.stringify(state.library));

    try {
        const response = await fetch("/api/import/file", {
            method: "POST",
            body: formData
        });
        const data = await response.json();
        if (!response.ok) {
            throw new Error(data.error || "Import failed.");
        }
        applyImportedPayload(data, data.importedFrom || file.name);
    } catch (error) {
        toast(els.builderStatus, error.message, "error");
    } finally {
        els.importFileInput.value = "";
    }
}

function buildExportPayload() {
    return {
        company: state.library.company,
        logoData: localStorage.getItem(STORAGE_KEYS.logo) || "",
        quote: {
            preparedBy: els.preparedBy.value.trim(),
            preparedPhone: els.preparedPhone.value.trim(),
            preparedEmail: els.preparedEmail.value.trim(),
            to: els.quoteTo.value.trim(),
            proposalFor: els.proposalFor.value.trim(),
            city: els.quoteCity.value.trim(),
            quoteDate: els.quoteDate.value.trim()
        },
        summary: summary(),
        lines: state.quote.lines
    };
}

function exportExcel() {
    if (!state.quote.lines.length) return toast(els.builderStatus, "Add at least one line item before exporting Excel.", "error");
    const rows = [["S/N", "Main Category", "Subcategory", "Description", "Spec / Unit", "Qty", "Unit Price", "Total"]];
    groupLines().forEach(([categoryName, items]) => {
        rows.push([categoryName, "", "", "", "", "", "", ""]);
        items.forEach(line => rows.push([line.serialNo, line.category, line.subcategory, line.description.replace(/\n/g, " "), line.spec, line.quantity, line.unitPrice, line.lineTotal]));
    });
    const calc = summary();
    rows.push(["", "", "", "", "", "", "Subtotal", calc.subtotal]);
    rows.push(["", "", "", "", "", "", "Discount", calc.discountAmount]);
    rows.push(["", "", "", "", "", "", "Tax", calc.taxAmount]);
    rows.push(["", "", "", "", "", "", "Grand Total", calc.grandTotal]);
    const csv = rows.map(row => row.map(value => `"${String(value ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8;" }));
    link.download = `reliance-quote-${Date.now()}.csv`;
    link.click();
}

async function exportFile(url, extension) {
    if (!state.quote.lines.length) return toast(els.builderStatus, `Add at least one line item before exporting ${extension.toUpperCase()}.`, "error");
    try {
        saveQuote();
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(buildExportPayload())
        });
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `Failed to export ${extension}.`);
        }
        const blob = await response.blob();
        const disposition = response.headers.get("Content-Disposition") || "";
        const match = disposition.match(/filename="?([^"]+)"?/i);
        const filename = match ? match[1] : `reliance-quote.${extension}`;
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        toast(els.builderStatus, `${extension.toUpperCase()} export created.`, "success");
    } catch (error) {
        toast(els.builderStatus, error.message, "error");
    }
}

function clearAllQuote() {
    if (!window.confirm("Clear the full quote draft?")) return;
    state.quote.lines = [];
    state.quote.selectedLineId = null;
    saveQuote();
    renderTable();
    renderPreview();
    toast(els.builderStatus, "Full quote cleared.", "success");
}

function attachLogo(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
        localStorage.setItem(STORAGE_KEYS.logo, reader.result);
        renderPreview();
        toast(els.builderStatus, "Logo saved locally for future quotations.", "success");
    };
    reader.readAsDataURL(file);
}

function bind() {
    els.librarySearch.addEventListener("input", renderLibraryTree);
    els.libraryTabs.addEventListener("click", event => {
        const button = event.target.closest("[data-action='filter-library']");
        if (!button) return;
        state.ui.activeLibraryCategoryId = button.dataset.categoryId;
        renderLibraryTabs();
        renderLibraryTree();
    });
    els.categorySelect.addEventListener("change", () => {
        state.ui.activeLibraryCategoryId = els.categorySelect.value || state.ui.activeLibraryCategoryId;
        renderLibraryTabs();
        renderSubcategoryOptions();
        renderLibraryTree();
    });
    els.subcategorySelect.addEventListener("change", renderItemOptions);
    els.itemSelect.addEventListener("change", loadItemToForm);
    els.itemSearchInput.addEventListener("input", () => {
        const exact = (getSubcategory(els.categorySelect.value, els.subcategorySelect.value)?.items || [])
            .find(item => item.name.toLowerCase() === els.itemSearchInput.value.trim().toLowerCase());
        if (exact) {
            pickItem(exact.id);
        } else {
            els.itemSelect.value = "";
            renderItemQuickList();
        }
    });
    els.itemQuickList.addEventListener("click", event => {
        if (event.target.dataset.action === "pick-item") {
            pickItem(event.target.dataset.itemId);
        }
    });
    document.getElementById("addLineButton").addEventListener("click", () => addOrUpdateLine("add"));
    els.updateLineButton.addEventListener("click", () => addOrUpdateLine("update"));
    document.getElementById("clearFormButton").addEventListener("click", clearEntry);
    document.getElementById("quickSaveItemButton").addEventListener("click", () => saveCurrentItem());
    document.getElementById("duplicateLineButton").addEventListener("click", () => state.quote.selectedLineId && duplicateLine(state.quote.selectedLineId));
    document.getElementById("deleteLineButton").addEventListener("click", () => state.quote.selectedLineId && deleteLine(state.quote.selectedLineId));
    document.getElementById("clearAllButton").addEventListener("click", clearAllQuote);

    [els.preparedBy, els.preparedPhone, els.preparedEmail, els.quoteDate, els.quoteTo, els.quoteCity, els.proposalFor, els.discountInput, els.discountType, els.taxInput, els.taxType]
        .forEach(element => {
            element.addEventListener("input", () => { saveQuote(); renderTable(); renderPreview(); });
            element.addEventListener("change", () => { saveQuote(); renderTable(); renderPreview(); });
        });

    document.getElementById("themeButton").addEventListener("click", () => setTheme(document.body.dataset.theme === "dark" ? "light" : "dark"));
    document.getElementById("printButton").addEventListener("click", () => window.print());
    document.getElementById("docxButton").addEventListener("click", () => exportFile("/api/export/docx", "docx"));
    document.getElementById("jsonButton").addEventListener("click", () => exportFile("/api/export/json", "json"));
    document.getElementById("importFileButton").addEventListener("click", () => els.importFileInput.click());
    document.getElementById("uploadLogoButton").addEventListener("click", () => els.logoInput.click());
    els.importFileInput.addEventListener("change", event => importQuoteFile(event.target.files[0]));
    els.logoInput.addEventListener("change", event => attachLogo(event.target.files[0]));

    els.quoteTableBody.addEventListener("click", event => {
        const action = event.target.dataset.action;
        const lineId = event.target.dataset.lineId;
        if (action === "edit") editLine(lineId);
        if (action === "copy") duplicateLine(lineId);
        if (action === "delete") deleteLine(lineId);
    });

    els.libraryTree.addEventListener("click", event => {
        const { action, categoryId, subcategoryId } = event.target.dataset;
        if (!action) return;
        event.preventDefault();
        event.stopPropagation();
        if (action === "rename-category") renameCategory(categoryId);
        if (action === "delete-category") removeCategory(categoryId);
        if (action === "rename-subcategory") renameSubcategory(categoryId, subcategoryId);
        if (action === "delete-subcategory") removeSubcategory(categoryId, subcategoryId);
        if (action === "add-item-here") openItemDialog({ categoryId, subcategoryId });
        if (action === "use-item") useLibraryItem(categoryId, subcategoryId, event.target.dataset.itemId);
        if (action === "rename-item") renameItem(categoryId, subcategoryId, event.target.dataset.itemId);
        if (action === "delete-item") removeItem(categoryId, subcategoryId, event.target.dataset.itemId);
    });

    document.getElementById("addCategoryButton").addEventListener("click", () => document.getElementById("categoryDialog").showModal());
    document.getElementById("addSubcategoryButton").addEventListener("click", () => document.getElementById("subcategoryDialog").showModal());
    document.getElementById("addItemButton").addEventListener("click", () => openItemDialog({ categoryId: els.categorySelect.value, subcategoryId: els.subcategorySelect.value }));
    document.getElementById("addSpecButton").addEventListener("click", () => document.getElementById("specDialog").showModal());
    els.itemDialogCategorySelect.addEventListener("change", syncItemDialogSubcategories);
    document.getElementById("saveCategoryDialogButton").addEventListener("click", event => {
        event.preventDefault();
        addCategory(els.newCategoryName.value, els.newCategoryColor.value);
        els.newCategoryName.value = "";
        els.newCategoryColor.value = "";
        document.getElementById("categoryDialog").close();
    });
    document.getElementById("saveSubcategoryDialogButton").addEventListener("click", event => {
        event.preventDefault();
        addSubcategory(els.subcategoryCategorySelect.value, els.newSubcategoryName.value);
        els.newSubcategoryName.value = "";
        document.getElementById("subcategoryDialog").close();
    });
    document.getElementById("saveSpecDialogButton").addEventListener("click", event => {
        event.preventDefault();
        addSpec(els.newSpecValue.value);
        els.newSpecValue.value = "";
        document.getElementById("specDialog").close();
    });
    document.getElementById("saveItemDialogButton").addEventListener("click", event => {
        event.preventDefault();
        saveItemFromDialog();
    });

    document.addEventListener("keydown", event => {
        if (event.ctrlKey && event.key === "Enter") {
            event.preventDefault();
            addOrUpdateLine(state.quote.selectedLineId ? "update" : "add");
        }
    });
}

function autoSave() {
    saveQuote();
    toast(els.builderStatus, "Draft auto-saved locally.", "info");
}

function init() {
    loadState();
    syncSerials();
    cache();
    hydrateQuoteMeta();
    renderCategoryOptions();
    renderSpecOptions();
    renderLibraryTree();
    renderTable();
    renderPreview();
    bind();
    setInterval(autoSave, 30000);
}

init();
