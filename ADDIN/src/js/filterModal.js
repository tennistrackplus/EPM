/**
 * Lógica del Modal de Selección / Filtro con formato de Jerarquía y Selección por Rango
 */
const FilterModal = {
    currentData: null,
    itemsList: [],
    lastCheckedIndex: null,

    init() {
        this.bindEvents();
    },

    bindEvents() {
        const closeBtn = document.getElementById("closeModalBtn");
        const cancelBtn = document.getElementById("btnCancelModal");
        const applyBtn = document.getElementById("btnApplyModal");
        const selectAllChk = document.getElementById("chkSelectAll");
        const rangeBtn = document.getElementById("btnApplyRange");
        const searchInput = document.getElementById("modalSearchInput");

        if (closeBtn) closeBtn.addEventListener("click", () => this.close());
        if (cancelBtn) cancelBtn.addEventListener("click", () => this.close());
        if (applyBtn) applyBtn.addEventListener("click", () => this.apply());
        
        if (selectAllChk) {
            selectAllChk.addEventListener("change", (e) => {
                const checkboxes = document.querySelectorAll("#modalItemsContainer input[type='checkbox']");
                checkboxes.forEach(chk => chk.checked = e.target.checked);
            });
        }

        if (rangeBtn) {
            rangeBtn.addEventListener("click", () => this.applyRangeSelection());
        }

        if (searchInput) {
            searchInput.addEventListener("input", (e) => this.filterItems(e.target.value));
        }
    },

    async open(data) {
        this.currentData = data;
        this.lastCheckedIndex = null;
        
        const modal = document.getElementById("filterModal");
        const titleEl = document.getElementById("modalTitle");
        const container = document.getElementById("modalItemsContainer");
        const selectAllChk = document.getElementById("chkSelectAll");
        
        if (selectAllChk) selectAllChk.checked = false;
        if (titleEl) titleEl.innerText = `Filtrar: ${data.dim} -> ${data.name}`;

        if (container) {
            container.innerHTML = "<div style='padding: 10px; color: #605e5c;'>Cargando valores...</div>";
        }

        modal.style.display = "flex";

        // Cargar elementos
        this.itemsList = await window.ExcelService.readMetValuesForField(data.dim, data.name, data.isHierarchy);
        this.populateRangeDropdowns(this.itemsList);
        this.renderItems(this.itemsList, data.isHierarchy);
    },

    close() {
        const modal = document.getElementById("filterModal");
        if (modal) modal.style.display = "none";
        this.currentData = null;
    },

    populateRangeDropdowns(items) {
        const fromSelect = document.getElementById("rangeFromSelect");
        const toSelect = document.getElementById("rangeToSelect");

        if (!fromSelect || !toSelect) return;

        fromSelect.innerHTML = "";
        toSelect.innerHTML = "";

        items.forEach((item, index) => {
            const val = item.value;
            const optionFrom = document.createElement("option");
            optionFrom.value = index;
            optionFrom.textContent = item.level > 1 ? `${"  ".repeat(item.level - 1)}${val}` : val;
            fromSelect.appendChild(optionFrom);

            const optionTo = document.createElement("option");
            optionTo.value = index;
            optionTo.textContent = item.level > 1 ? `${"  ".repeat(item.level - 1)}${val}` : val;
            toSelect.appendChild(optionTo);
        });

        if (items.length > 0) {
            fromSelect.selectedIndex = 0;
            toSelect.selectedIndex = items.length - 1;
        }
    },

    applyRangeSelection() {
        const fromSelect = document.getElementById("rangeFromSelect");
        const toSelect = document.getElementById("rangeToSelect");
        
        if (!fromSelect || !toSelect) return;

        const startIndex = parseInt(fromSelect.value, 10);
        const endIndex = parseInt(toSelect.value, 10);

        const fromIdx = Math.min(startIndex, endIndex);
        const toIdx = Math.max(startIndex, endIndex);

        const container = document.getElementById("modalItemsContainer");
        const rows = container.querySelectorAll(".modal-item-row");

        rows.forEach((row, idx) => {
            const chk = row.querySelector("input[type='checkbox']");
            if (chk) {
                chk.checked = (idx >= fromIdx && idx <= toIdx);
            }
        });
    },

    renderItems(items, isHierarchy) {
        const container = document.getElementById("modalItemsContainer");
        if (!container) return;

        container.innerHTML = "";

        if (items.length === 0) {
            container.innerHTML = "<div style='padding: 10px; color: #605e5c;'>No hay elementos disponibles.</div>";
            return;
        }

        items.forEach((item, index) => {
            const row = document.createElement("div");
            row.className = "modal-item-row";
            row.dataset.index = index;

            // Formato de nivel/árbol para jerarquía
            const level = item.level || 1;
            const indentPixels = (level - 1) * 18;
            row.style.paddingLeft = `${indentPixels + 8}px`;

            if (isHierarchy && level > 1) {
                row.classList.add("hierarchy-child-row");
            }

            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.id = `chk_item_${index}`;
            checkbox.value = item.value;

            // Soporte para Shift + Clic en rango
            checkbox.addEventListener("click", (e) => {
                if (e.shiftKey && this.lastCheckedIndex !== null) {
                    const start = Math.min(this.lastCheckedIndex, index);
                    const end = Math.max(this.lastCheckedIndex, index);
                    const allCheckboxes = container.querySelectorAll("input[type='checkbox']");
                    for (let i = start; i <= end; i++) {
                        allCheckboxes[i].checked = checkbox.checked;
                    }
                }
                this.lastCheckedIndex = index;
            });

            const label = document.createElement("label");
            label.htmlFor = `chk_item_${index}`;
            
            const prefix = (isHierarchy && level > 1) ? "└─ " : "";
            label.innerHTML = `<span class="level-indicator">L${level}</span> ${prefix}<strong>${item.value}</strong>`;

            row.appendChild(checkbox);
            row.appendChild(label);
            container.appendChild(row);
        });
    },

    filterItems(query) {
        const q = query.toLowerCase();
        const rows = document.querySelectorAll("#modalItemsContainer .modal-item-row");
        rows.forEach(row => {
            const text = row.innerText.toLowerCase();
            row.style.display = text.includes(q) ? "flex" : "none";
        });
    },

    apply() {
        const checkedValues = [];
        const checkboxes = document.querySelectorAll("#modalItemsContainer input[type='checkbox']:checked");
        checkboxes.forEach(chk => checkedValues.push(chk.value));

        console.log(`Filtros aplicados a ${this.currentData.name}:`, checkedValues);
        this.close();
    }
};

window.FilterModal = FilterModal;