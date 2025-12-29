const state = {
    mapping: {},
    dropdownValues: {},
    requiredFields: [],
    uniqueFields: [],
    templatePath: "",
    templateHeaders: [],
    assets: [],
    settings: {},
    savedSearches: {},
    editAssetId: null,
    monitorData: { stats: {}, recent_modified: [], recent_added: [] },
    monitorTimer: null
};

const searchInput = document.querySelector("#searchInput");
const fieldSelect = document.querySelector("#fieldSelect");
const limitSelect = document.querySelector("#limitSelect");
const resultSummary = document.querySelector("#resultSummary");
const kpiCount = document.querySelector("#kpiCount");
const templatePathEl = document.querySelector("#templatePath");
const assetTableBody = document.querySelector("#assetTableBody");
const toast = document.querySelector("#toast");
const drawer = document.querySelector("#drawer");
const assetForm = document.querySelector("#assetForm");
const drawerTitle = document.querySelector("#drawer h3");

const settingsPanel = document.querySelector("#settingsPanel");
const settingsButton = document.querySelector("#settingsButton");
const closeSettings = document.querySelector("#closeSettings");
const saveSettings = document.querySelector("#saveSettings");
const settingsStatus = document.querySelector("#settingsStatus");
const settingTemplatePath = document.querySelector("#settingTemplatePath");
const settingOutputDir = document.querySelector("#settingOutputDir");
const settingDatabasePath = document.querySelector("#settingDatabasePath");
const settingTheme = document.querySelector("#settingTheme");
const settingPresets = document.querySelector("#settingPresets");
const settingSavedSearches = document.querySelector("#settingSavedSearches");
const settingsFields = document.querySelector("#settingsFields");
const settingsMonitor = document.querySelector("#settingsMonitor");
const settingsReports = document.querySelector("#settingsReports");

// Tools
const importFile = document.querySelector("#importFile");
const importButton = document.querySelector("#importButton");
const exportTemplateButton = document.querySelector("#exportTemplateButton");
const exportAllButton = document.querySelector("#exportAllButton");
const bulkField = document.querySelector("#bulkField");
const bulkOp = document.querySelector("#bulkOp");
const bulkValue = document.querySelector("#bulkValue");
const bulkPreset = document.querySelector("#bulkPreset");
const runBulk = document.querySelector("#runBulk");
const savedSearchName = document.querySelector("#savedSearchName");
const savedSearchSelect = document.querySelector("#savedSearchSelect");
const saveSearchBtn = document.querySelector("#saveSearch");
const applySearchBtn = document.querySelector("#applySearch");
const deleteSearchBtn = document.querySelector("#deleteSearch");
const monitorSummary = document.querySelector("#monitorSummary");
const monitorPanel = document.querySelector("#monitorPanel");
const monitorClose = document.querySelector("#monitorClose");
const monitorToggle = document.querySelector("#monitorToggle");
const monitorStats = document.querySelector("#monitorStats");
const monitorModified = document.querySelector("#monitorModified");
const monitorAdded = document.querySelector("#monitorAdded");

const refreshButton = document.querySelector("#refreshButton");
const newAssetButton = document.querySelector("#newAssetButton");
const closeDrawer = document.querySelector("#closeDrawer");
const cancelForm = document.querySelector("#cancelForm");
const clearFilters = document.querySelector("#clearFilters");

const debounce = (fn, delay = 250) => {
    let timer;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
};

async function fetchJson(url, options) {
    const response = await fetch(url, options);
    if (!response.ok) {
        let detail;
        try {
            const data = await response.json();
            detail = data.detail || data;
        } catch (err) {
            detail = response.statusText;
        }
        throw new Error(typeof detail === "string" ? detail : JSON.stringify(detail));
    }
    return response.json();
}

function showToast(message, tone = "info") {
    toast.textContent = message;
    toast.dataset.tone = tone;
    toast.hidden = false;
    toast.classList.add("visible");
    setTimeout(() => toast.classList.remove("visible"), 3000);
}

function populateFieldSelect(mapping) {
    const preferred = ["System Name", "Serial Number", "Asset No."];
    const headers = Object.keys(mapping);
    const options = [];
    preferred.forEach((field) => {
        if (headers.includes(field)) options.push(field);
    });
    headers.forEach((h) => {
        if (!options.includes(h)) options.push(h);
    });

    fieldSelect.innerHTML = "";
    options.slice(0, 30).forEach((h) => {
        const opt = document.createElement("option");
        opt.value = h;
        opt.textContent = h;
        fieldSelect.appendChild(opt);
    });
    if (fieldSelect.options.length) {
        fieldSelect.value = options[0];
    }
}

function populateDatalists(dropdownValues) {
    const map = {
        manufacturerOptions: dropdownValues["*Manufacturer"] || dropdownValues["Manufacturer"] || [],
        modelOptions: dropdownValues["*Model"] || dropdownValues["Model"] || [],
        statusOptions: dropdownValues.Status || [],
        locationOptions: dropdownValues.Location || [],
        roomOptions: dropdownValues.Room || []
    };

    Object.entries(map).forEach(([id, values]) => {
        const list = document.getElementById(id);
        if (!list) return;
        list.innerHTML = "";
        values.slice(0, 100).forEach((value) => {
            const opt = document.createElement("option");
            opt.value = value;
            list.appendChild(opt);
        });
    });
}

function renderTable(items) {
    if (!items.length) {
        assetTableBody.innerHTML = '<tr><td colspan="8" class="muted">No assets found</td></tr>';
        return;
    }

    const rows = items.map((asset) => {
        const modified = formatDate(asset.modified_date || asset.created_date);
        return `
            <tr data-id="${asset.id}">
                <td>${clean(asset.asset_no)}</td>
                <td>${clean(asset.system_name)}</td>
                <td>${clean(asset.manufacturer)}</td>
                <td>${clean(asset.model)}</td>
                <td>${clean(asset.serial_number)}</td>
                <td>${clean(asset.location || asset.room)}</td>
                <td>${modified}</td>
                <td>
                    <button class="ghost action-edit" data-id="${asset.id}">Edit</button>
                    <button class="ghost action-delete" data-id="${asset.id}">Delete</button>
                </td>
            </tr>
        `;
    });
    assetTableBody.innerHTML = rows.join("");
}

function orderedHeaders(headers, selected) {
    const cleanSel = (selected || []).filter(Boolean);
    const remainder = headers.filter((h) => !cleanSel.includes(h));
    return [...cleanSel, ...remainder];
}

function renderChecklist(target, label, key, headers, selected) {
    const card = document.createElement("div");
    card.className = "card";
    const title = document.createElement("h3");
    title.className = "group-title";
    title.textContent = label;
    card.appendChild(title);

    const list = document.createElement("div");
    list.className = "pill-list";
    const ordered = orderedHeaders(headers, selected);
    ordered.forEach((field) => {
        const item = document.createElement("label");
        item.className = "pill-item";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.dataset.group = key;
        checkbox.value = field;
        checkbox.checked = (selected || []).includes(field);
        const text = document.createElement("span");
        text.textContent = field;
        item.appendChild(checkbox);
        item.appendChild(text);
        list.appendChild(item);
    });
    card.appendChild(list);
    target.appendChild(card);
}

function populateSettingsForm() {
    const cfg = state.settings || {};
    settingsStatus.textContent = "Loaded";
    settingTemplatePath.value = cfg.default_template_path || "";
    settingOutputDir.value = cfg.output_directory || "";
    settingDatabasePath.value = cfg.database_path || "";
    settingTheme.value = (cfg.theme || "dark").toLowerCase();
    settingPresets.value = JSON.stringify(cfg.bulk_update_presets || {}, null, 2);
    settingSavedSearches.value = JSON.stringify(cfg.saved_searches || {}, null, 2);

    const headers = state.templateHeaders || [];
    settingsFields.innerHTML = "";
    const fieldGroups = [
        { key: "dropdown_fields", label: "Dropdown Fields", values: cfg.dropdown_fields || [] },
        { key: "required_fields", label: "Required Fields", values: cfg.required_fields || [] },
        { key: "excluded_fields", label: "Excluded Fields", values: cfg.excluded_fields || [] },
        { key: "unique_fields", label: "Unique Fields", values: cfg.unique_fields || [] },
    ];
    fieldGroups.forEach((group) => renderChecklist(settingsFields, group.label, group.key, headers, group.values));

    settingsMonitor.innerHTML = "";
    const monitorGroups = [
        { key: "monitor_primary_fields", label: "Monitor Primary" },
        { key: "monitor_secondary_fields", label: "Monitor Secondary" },
        { key: "monitor_tertiary_fields", label: "Monitor Tertiary" },
    ];
    monitorGroups.forEach((group) => renderChecklist(settingsMonitor, group.label, group.key, headers, cfg[group.key] || []));

    settingsReports.innerHTML = "";
    const reportGroups = [
        { key: "label_output_fields", label: "Label Output Fields" },
        { key: "hmr_fields", label: "HMR Fields" },
        { key: "destruction_report_fields", label: "Destruction Report Fields" },
    ];
    reportGroups.forEach((group) => renderChecklist(settingsReports, group.label, group.key, headers, cfg[group.key] || []));

    populateBulkOptions();
    populateSavedSearches();
}

function collectGroup(key) {
    const nodes = Array.from(document.querySelectorAll(`input[data-group="${key}"]`));
    return nodes.filter((n) => n.checked).map((n) => n.value);
}

function findAsset(id) {
    return state.assets.find((a) => String(a.id) === String(id));
}

function currentFilters() {
    return {
        field: fieldSelect.value,
        query: searchInput.value.trim(),
        limit: Number(limitSelect.value || 100)
    };
}

function populateBulkOptions() {
    if (!bulkField) return;
    const headers = state.templateHeaders && state.templateHeaders.length ? state.templateHeaders : Object.keys(state.mapping || {});
    bulkField.innerHTML = "";
    headers.slice(0, 200).forEach((h) => {
        const opt = document.createElement("option");
        opt.value = h;
        opt.textContent = h;
        bulkField.appendChild(opt);
    });

    bulkPreset.innerHTML = "";
    const presets = state.settings.bulk_update_presets || {};
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "(no preset)";
    bulkPreset.appendChild(placeholder);
    Object.keys(presets).forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        bulkPreset.appendChild(opt);
    });
}

function populateSavedSearches() {
    if (!savedSearchSelect) return;
    const searches = state.savedSearches || {};
    savedSearchSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select saved search";
    savedSearchSelect.appendChild(placeholder);
    Object.keys(searches).forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        savedSearchSelect.appendChild(opt);
    });
}

async function openEditFromId(id) {
    const local = findAsset(id);
    if (local) {
        openDrawer(local);
        return;
    }
    try {
        const asset = await fetchJson(`/api/assets/${id}`);
        openDrawer(asset);
    } catch (err) {
        showToast(err.message || "Could not load asset", "error");
    }
}

async function deleteAsset(id) {
    if (!confirm("Delete this asset? This cannot be undone.")) return;
    try {
        await fetchJson(`/api/assets/${id}`, { method: "DELETE" });
        showToast("Asset deleted", "success");
        await loadAssets();
    } catch (err) {
        showToast(err.message || "Delete failed", "error");
    }
}

function clean(value) {
    return value ? String(value) : "—";
}

function formatDate(value) {
    if (!value) return "—";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value.slice(0, 10);
    return date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

async function loadFields() {
    const data = await fetchJson("/api/fields");
    state.mapping = data.column_mapping || {};
    state.dropdownValues = data.dropdown_values || {};
    state.requiredFields = data.required_fields || [];
    state.uniqueFields = data.unique_fields || [];
    if (data.field_metadata) {
        state.templateHeaders = Object.keys(data.field_metadata);
    }
    if (Object.keys(state.settings || {}).length) {
        populateSettingsForm();
    }

    populateFieldSelect(state.mapping);
    populateDatalists(state.dropdownValues);
}

async function loadConfig() {
    const data = await fetchJson("/api/config");
    state.templatePath = data.template_path || "";
    templatePathEl.textContent = state.templatePath ? `Template: ${state.templatePath}` : "Template: not set";
}

async function loadSettings() {
    const data = await fetchJson("/api/settings");
    state.settings = data;
    state.templateHeaders = data.template_headers || state.templateHeaders || [];
    state.savedSearches = data.saved_searches || {};
    populateSettingsForm();
}

async function loadAssets() {
    const field = fieldSelect.value;
    const query = searchInput.value.trim();
    const limit = Number(limitSelect.value || 100);

    resultSummary.textContent = "Loading…";
    assetTableBody.innerHTML = '<tr><td colspan="7" class="muted">Loading…</td></tr>';

    const params = new URLSearchParams({ limit: String(limit) });
    if (field) params.set("field", field);
    if (query) params.set("query", query);

    try {
        const data = await fetchJson(`/api/assets?${params.toString()}`);
        state.assets = data.items || [];
        renderTable(state.assets);
        const total = data.total ?? state.assets.length;
        resultSummary.textContent = `${state.assets.length} shown · total ${total}`;
        kpiCount.textContent = `${total} assets`;
    } catch (err) {
        resultSummary.textContent = "Error loading assets";
        assetTableBody.innerHTML = `<tr><td colspan="8" class="muted">${err.message}</td></tr>`;
    }
}

function setFormFromAsset(asset) {
    const reverse = Object.entries(state.mapping || {}).reduce((acc, [header, col]) => {
        acc[col] = header;
        return acc;
    }, {});
    const elements = Array.from(assetForm.elements).filter((el) => el.name);
    elements.forEach((el) => {
        const header = el.name;
        const dbCol = state.mapping[header] || header;
        const value = asset[dbCol];
        if (value !== undefined && value !== null) {
            if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
                el.value = value;
            }
        }
    });
}

function openDrawer(asset = null) {
    if (asset) {
        state.editAssetId = asset.id;
        drawerTitle.textContent = "Edit asset";
        setFormFromAsset(asset);
    } else {
        state.editAssetId = null;
        drawerTitle.textContent = "New asset";
        assetForm.reset();
    }
    drawer.hidden = false;
    drawer.classList.add("open");
}

function closeDrawerPanel() {
    drawer.classList.remove("open");
    setTimeout(() => {
        drawer.hidden = true;
        assetForm.reset();
        state.editAssetId = null;
        drawerTitle.textContent = "New asset";
    }, 180);
}

async function handleSubmit(event) {
    event.preventDefault();
    const formData = new FormData(assetForm);
    const fields = {};
    for (const [key, value] of formData.entries()) {
        if (value === "") continue;
        fields[key] = value;
    }

    try {
        const isEdit = !!state.editAssetId;
        const url = isEdit ? `/api/assets/${state.editAssetId}` : "/api/assets";
        const method = isEdit ? "PUT" : "POST";
        await fetchJson(url, {
            method,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fields })
        });
        showToast(isEdit ? "Asset updated" : "Asset saved", "success");
        await loadAssets();
        closeDrawerPanel();
    } catch (err) {
        showToast(err.message || "Could not save asset", "error");
    }
}

async function saveSettingsHandler() {
    const payload = {
        theme: settingTheme.value,
        default_template_path: settingTemplatePath.value.trim(),
        output_directory: settingOutputDir.value.trim(),
        database_path: settingDatabasePath.value.trim(),
        dropdown_fields: collectGroup("dropdown_fields"),
        required_fields: collectGroup("required_fields"),
        excluded_fields: collectGroup("excluded_fields"),
        unique_fields: collectGroup("unique_fields"),
        monitor_primary_fields: collectGroup("monitor_primary_fields"),
        monitor_secondary_fields: collectGroup("monitor_secondary_fields"),
        monitor_tertiary_fields: collectGroup("monitor_tertiary_fields"),
        label_output_fields: collectGroup("label_output_fields"),
        hmr_fields: collectGroup("hmr_fields"),
        destruction_report_fields: collectGroup("destruction_report_fields"),
    };

    try {
        payload.bulk_update_presets = JSON.parse(settingPresets.value || "{}");
    } catch (err) {
        showToast("Presets JSON is invalid", "error");
        return;
    }

    try {
        payload.saved_searches = JSON.parse(settingSavedSearches.value || "{}");
    } catch (err) {
        showToast("Saved searches JSON is invalid", "error");
        return;
    }

    try {
        await fetchJson("/api/settings", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });
        showToast("Settings saved", "success");
        await Promise.all([loadSettings(), loadFields(), loadAssets()]);
    } catch (err) {
        showToast(err.message || "Could not save settings", "error");
    }
}

async function importHandler() {
    if (!importFile.files.length) return;
    const file = importFile.files[0];
    const form = new FormData();
    form.append("file", file);
    try {
        const res = await fetchJson("/api/assets/import", { method: "POST", body: form });
        showToast(`Imported ${res.imported} assets`, "success");
        await loadAssets();
    } catch (err) {
        showToast(err.message || "Import failed", "error");
    } finally {
        importFile.value = "";
    }
}

function exportHandler(mode) {
    const q = searchInput.value.trim();
    const field = fieldSelect.value;
    const url = new URL(`/api/assets/export`, window.location.origin);
    url.searchParams.set("mode", mode);
    if (q) url.searchParams.set("query", q);
    if (field) url.searchParams.set("field", field);
    window.open(url.toString(), "_blank");
}

async function runBulkHandler() {
    let operations = [];
    const presetName = bulkPreset.value;
    if (presetName && state.settings.bulk_update_presets && state.settings.bulk_update_presets[presetName]) {
        operations = state.settings.bulk_update_presets[presetName].fields || [];
    } else {
        if (!bulkField.value || !bulkValue.value) {
            showToast("Set a field and value for bulk update", "error");
            return;
        }
        operations = [{ field: bulkField.value, operation: bulkOp.value, value: bulkValue.value }];
    }

    const filters = {};
    const q = searchInput.value.trim();
    if (q && fieldSelect.value) {
        filters[fieldSelect.value] = q;
    }

    try {
        const res = await fetchJson("/api/assets/bulk-update", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ operations, filters })
        });
        showToast(`Updated ${res.updated} assets`, "success");
        await loadAssets();
    } catch (err) {
        showToast(err.message || "Bulk update failed", "error");
    }
}

async function refreshSavedSearches() {
    try {
        const data = await fetchJson("/api/saved-searches");
        state.savedSearches = data || {};
        populateSavedSearches();
    } catch (err) {
        // ignore
    }
}

async function saveSearchHandler() {
    const name = savedSearchName.value.trim();
    if (!name) {
        showToast("Enter a name for the search", "error");
        return;
    }
    const filters = currentFilters();
    try {
        await fetchJson("/api/saved-searches", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ name, filters })
        });
        showToast("Saved search", "success");
        savedSearchName.value = "";
        await refreshSavedSearches();
    } catch (err) {
        showToast(err.message || "Could not save search", "error");
    }
}

function applySavedSearchHandler() {
    const name = savedSearchSelect.value;
    if (!name || !state.savedSearches[name]) return;
    const filters = state.savedSearches[name];
    if (filters.field) fieldSelect.value = filters.field;
    if (filters.query) searchInput.value = filters.query;
    if (filters.limit) limitSelect.value = String(filters.limit);
    loadAssets();
}

async function deleteSavedSearchHandler() {
    const name = savedSearchSelect.value;
    if (!name) return;
    try {
        await fetchJson(`/api/saved-searches/${encodeURIComponent(name)}`, { method: "DELETE" });
        showToast("Deleted saved search", "success");
        await refreshSavedSearches();
    } catch (err) {
        showToast(err.message || "Could not delete search", "error");
    }
}

async function loadMonitor() {
    try {
        const data = await fetchJson("/api/monitor");
        state.monitorData = data;
        const total = data.stats?.total_assets ?? 0;
        monitorSummary.textContent = `${total} assets • ${Math.round((data.stats?.database_size_mb || 0) * 10) / 10} MB`;
        renderMonitor();
    } catch (err) {
        monitorSummary.textContent = "Monitor unavailable";
    }
}

function renderMonitorList(target, items) {
    if (!items || !items.length) {
        target.textContent = "No entries";
        return;
    }
    target.innerHTML = items
        .slice(0, 50)
        .map((item) => {
            const when = formatDate(item.modified_date || item.created_date);
            const title = item.system_name || item.asset_no || item.serial_number || `ID ${item.id}`;
            const extra = [item.manufacturer, item.model].filter(Boolean).join(" · ");
            return `<div class="monitor-item"><strong>${clean(title)}</strong><div class="sub">${clean(extra)} • ${when}</div></div>`;
        })
        .join("");
}

function renderMonitor() {
    if (!monitorPanel || monitorPanel.hidden) return;
    const data = state.monitorData || {};
    const stats = data.stats || {};
    monitorStats.textContent = `${stats.total_assets ?? 0} assets • ${Math.round((stats.database_size_mb || 0) * 10) / 10} MB • audit log ${stats.total_audit_entries ?? 0}`;
    renderMonitorList(monitorModified, data.recent_modified || []);
    renderMonitorList(monitorAdded, data.recent_added || []);
}

function openMonitor() {
    if (!monitorPanel) return;
    monitorPanel.hidden = false;
    monitorPanel.classList.add("open");
    loadMonitor();
    if (state.monitorTimer) clearInterval(state.monitorTimer);
    state.monitorTimer = setInterval(loadMonitor, 60000);
}

function closeMonitor() {
    if (!monitorPanel) return;
    monitorPanel.classList.remove("open");
    monitorPanel.hidden = true;
    if (state.monitorTimer) {
        clearInterval(state.monitorTimer);
        state.monitorTimer = null;
    }
}

function bindEvents() {
    searchInput.addEventListener("input", debounce(loadAssets, 300));
    fieldSelect.addEventListener("change", loadAssets);
    limitSelect.addEventListener("change", loadAssets);
    refreshButton.addEventListener("click", loadAssets);
    clearFilters.addEventListener("click", () => {
        searchInput.value = "";
        loadAssets();
    });

    newAssetButton.addEventListener("click", openDrawer);
    closeDrawer.addEventListener("click", closeDrawerPanel);
    cancelForm.addEventListener("click", (e) => {
        e.preventDefault();
        closeDrawerPanel();
    });

    assetForm.addEventListener("submit", handleSubmit);

    assetTableBody.addEventListener("click", (e) => {
        const editBtn = e.target.closest(".action-edit");
        const delBtn = e.target.closest(".action-delete");
        if (editBtn) {
            const id = editBtn.dataset.id;
            openEditFromId(id);
        } else if (delBtn) {
            const id = delBtn.dataset.id;
            deleteAsset(id);
        }
    });

    importButton.addEventListener("click", () => importFile.click());
    importFile.addEventListener("change", importHandler);
    exportTemplateButton.addEventListener("click", () => exportHandler("template"));
    exportAllButton.addEventListener("click", () => exportHandler("all"));
    runBulk.addEventListener("click", runBulkHandler);

    saveSearchBtn.addEventListener("click", saveSearchHandler);
    applySearchBtn.addEventListener("click", applySavedSearchHandler);
    deleteSearchBtn.addEventListener("click", deleteSavedSearchHandler);

    monitorToggle.addEventListener("click", () => {
        if (monitorPanel.hidden) {
            openMonitor();
        } else {
            closeMonitor();
        }
    });
    monitorClose.addEventListener("click", closeMonitor);

    settingsButton.addEventListener("click", () => {
        settingsPanel.hidden = false;
        settingsPanel.scrollIntoView({ behavior: "smooth" });
    });
    closeSettings.addEventListener("click", () => {
        settingsPanel.hidden = true;
    });
    saveSettings.addEventListener("click", saveSettingsHandler);
}

async function bootstrap() {
    bindEvents();
    await Promise.all([loadSettings(), loadConfig(), loadFields(), refreshSavedSearches()]);
    await Promise.all([loadAssets(), loadMonitor()]);
}

bootstrap();
