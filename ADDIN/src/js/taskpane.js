/**
 * Lógica principal del TaskPane (Drag & Drop entre zonas, informe asimétrico y carga de interfaz)
 */
document.addEventListener("DOMContentLoaded", () => {
    if (typeof Office !== "undefined" && Office.onReady) {
        Office.onReady(() => {
            TaskPaneApp.init();
        });
    } else {
        TaskPaneApp.init();
    }
});

const TaskPaneApp = {
    draggedElementData: null,

    async init() {
        if (typeof FilterModal !== "undefined" && FilterModal.init) {
            FilterModal.init();
        }
        this.bindEvents();
        await this.loadFields();
    },

    bindEvents() {
        const btnRefresh = document.getElementById("btnRefresh");
        if (btnRefresh) {
            btnRefresh.addEventListener("click", () => this.loadFields());
        }

        const fieldSearch = document.getElementById("fieldSearch");
        if (fieldSearch) {
            fieldSearch.addEventListener("input", (e) => this.filterFields(e.target.value));
        }

        // Configurar Checkboxes de Informe Asimétrico
        const chkAsymmetricRows = document.getElementById("chkAsymmetricRows");
        const chkAsymmetricCols = document.getElementById("chkAsymmetricCols");

        if (chkAsymmetricRows) {
            chkAsymmetricRows.addEventListener("change", (e) => {
                console.log("Informe Asimétrico en Filas:", e.target.checked);
            });
        }
        if (chkAsymmetricCols) {
            chkAsymmetricCols.addEventListener("change", (e) => {
                console.log("Informe Asimétrico en Columnas:", e.target.checked);
            });
        }

        // Configurar Dropzones (Soporta las clases Fluent UI .zone-card y .dropzone-box)
        const zones = document.querySelectorAll(".zone-card, .dropzone-box");
        zones.forEach(zone => {
            zone.addEventListener("dragover", (e) => this.handleDragOver(e));
            zone.addEventListener("dragleave", (e) => this.handleDragLeave(e));
            zone.addEventListener("drop", (e) => this.handleDrop(e));
        });
    },

    async loadFields() {
        const container = document.getElementById("availableFieldsContainer");
        if (!container) return;

        container.innerHTML = "<div style='color: #605e5c;'>Cargando dimensiones...</div>";

        try {
            const result = await window.ExcelService.readDim2Data();

            if (result.error) {
                container.innerHTML = `<div style='color: #a80000; padding: 4px;'>⚠️ ${result.error}</div>`;
                return;
            }

            const dimensions = result.data || [];
            container.innerHTML = "";

            if (dimensions.length === 0) {
                container.innerHTML = "<div style='color: #605e5c;'>No se encontraron campos en DIM2</div>";
                return;
            }

            dimensions.forEach(dim => {
                const group = document.createElement("div");
                group.className = "dimension-group";

                const header = document.createElement("div");
                header.className = "dimension-header";
                header.innerText = dim.dimension.toLowerCase();
                group.appendChild(header);

                // 1. Jerarquías
                dim.hierarchies.forEach(hier => {
                    const item = this.createFieldElement(dim.dimension, hier, true);
                    group.appendChild(item);
                });

                // 2. Atributos
                dim.attributes.forEach(att => {
                    const item = this.createFieldElement(dim.dimension, att, false);
                    group.appendChild(item);
                });

                container.appendChild(group);
            });
        } catch (err) {
            console.error("Error al cargar dimensiones:", err);
            container.innerHTML = `<div style='color: #a80000; padding: 4px;'>❌ Error al cargar datos: ${err.message || err}</div>`;
        }
    },

    createFieldElement(dim, name, isHierarchy) {
        const div = document.createElement("div");
        div.className = "field-item";
        div.draggable = true;

        const iconClass = isHierarchy ? "field-icon hierarchy-icon" : "field-icon";
        const iconSymbol = isHierarchy ? "🗂️" : "📄";

        div.innerHTML = `
            <span class="${iconClass}">${iconSymbol}</span>
            <span class="field-label">${name.toLowerCase()}</span>
        `;

        const fieldData = { dim, name, isHierarchy };

        div.addEventListener("dragstart", (e) => {
            this.draggedElementData = { data: fieldData, sourceTag: null };
            div.classList.add("dragging");
            e.dataTransfer.setData("text/plain", JSON.stringify(fieldData));
        });

        div.addEventListener("dragend", () => {
            div.classList.remove("dragging");
        });

        return div;
    },

    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add("drag-over");
    },

    handleDragLeave(e) {
        e.currentTarget.classList.remove("drag-over");
    },

    handleDrop(e) {
        e.preventDefault();
        const dropzoneBox = e.currentTarget;
        dropzoneBox.classList.remove("drag-over");

        if (!this.draggedElementData) return;

        const { data, sourceTag } = this.draggedElementData;
        const targetContent = dropzoneBox.querySelector(".dropzone-content");
        const zoneId = dropzoneBox.getAttribute("data-zone");

        // Si se arrastra una etiqueta existente desde otra zona
        if (sourceTag) {
            if (sourceTag.parentElement !== targetContent) {
                sourceTag.remove();
                this.addTagToZone(targetContent, data, zoneId);
            }
        } else {
            // Se arrastra desde el panel lateral principal
            this.addTagToZone(targetContent, data, zoneId);
        }

        this.draggedElementData = null;
    },

    addTagToZone(container, data, zoneId) {
        const existing = Array.from(container.children).find(
            child => child.dataset.fieldName === data.name && child.dataset.dim === data.dim
        );

        if (existing) return;

        const tag = document.createElement("div");
        tag.className = "dropped-tag";
        tag.draggable = true;
        tag.dataset.dim = data.dim;
        tag.dataset.fieldName = data.name;
        tag.dataset.isHierarchy = data.isHierarchy;

        tag.innerHTML = `
            <span class="dropped-tag-title">${data.name}</span>
            <span class="dropped-tag-remove">&times;</span>
        `;

        // Permitir arrastrar la etiqueta ya soltada hacia otra zona
        tag.addEventListener("dragstart", (e) => {
            e.stopPropagation();
            this.draggedElementData = { data, sourceTag: tag };
            tag.classList.add("dragging");
            e.dataTransfer.setData("text/plain", JSON.stringify(data));
        });

        tag.addEventListener("dragend", () => {
            tag.classList.remove("dragging");
        });

        tag.addEventListener("dblclick", () => {
            if (typeof FilterModal !== "undefined" && FilterModal.open) {
                FilterModal.open(data);
            }
        });

        tag.querySelector(".dropped-tag-remove").addEventListener("click", (e) => {
            e.stopPropagation();
            tag.remove();
        });

        container.appendChild(tag);
    },

    filterFields(query) {
        const q = query.toLowerCase();
        const items = document.querySelectorAll(".field-item");
        items.forEach(item => {
            const label = item.querySelector(".field-label").innerText.toLowerCase();
            item.style.display = label.includes(q) ? "flex" : "none";
        });
    }
};