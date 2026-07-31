/**
 * Gestor del diálogo Modal para el filtrado de valores
 */
const FilterModal = {
    currentField: null,
    allValues: [],

    init() {
        this.bindEvents();
    },

    bindEvents() {
        document.getElementById("btnCloseModal").addEventListener("click", () => this.hide());
        document.getElementById("btnCancelFilter").addEventListener("click", () => this.hide());
        
        document.getElementById("modalSearchInput").addEventListener("input", (e) => {
            this.filterList(e.target.value);
        });

        document.getElementById("chkSelectAll").addEventListener("change", (e) => {
            const checkboxes = document.querySelectorAll(".val-chk");
            checkboxes.forEach(chk => chk.checked = e.target.checked);
        });

        document.getElementById("btnApplyFilter").addEventListener("click", () => {
            this.applyFilter();
        });
    },

    async open(fieldData) {
        this.currentField = fieldData;
        document.getElementById("modalTitle").innerText = `Filtro: ${fieldData.name}`;
        
        const container = document.getElementById("modalValuesContainer");
        container.innerHTML = "<div style='padding: 8px;'>Cargando valores...</div>";
        document.getElementById("filterModal").classList.remove("hidden");

        // Leer datos de MET
        this.allValues = await ExcelService.readMetValuesForField(
            fieldData.dim,
            fieldData.name,
            fieldData.isHierarchy
        );

        this.renderValues(this.allValues);
    },

    renderValues(values) {
        const container = document.getElementById("modalValuesContainer");
        container.innerHTML = "";

        if (values.length === 0) {
            container.innerHTML = "<div style='padding: 8px; color: #a19f9d;'>Sin valores disponibles</div>";
            return;
        }

        values.forEach(val => {
            const item = document.createElement("label");
            item.className = "value-item";
            item.innerHTML = `
                <input type="checkbox" class="val-chk" value="${val}" checked />
                <span>${val}</span>
            `;
            container.appendChild(item);
        });
    },

    filterList(searchText) {
        const query = searchText.toLowerCase();
        const items = document.querySelectorAll(".value-item");
        items.forEach(item => {
            const text = item.innerText.toLowerCase();
            item.style.display = text.includes(query) ? "flex" : "none";
        });
    },

    applyFilter() {
        const selected = [];
        const isExclude = document.getElementById("chkExcludeMode").checked;

        document.querySelectorAll(".val-chk:checked").forEach(chk => {
            selected.push(chk.value);
        });

        console.log(`Filtro aplicado para [${this.currentField.name}]:`, {
            modoExcluir: isExclude,
            valoresSeleccionados: selected
        });

        this.hide();
    },

    hide() {
        document.getElementById("filterModal").classList.add("hidden");
        document.getElementById("modalSearchInput").value = "";
    }
};