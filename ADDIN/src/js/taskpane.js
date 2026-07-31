/**
 * Lógica principal del TaskPane (Drag & Drop y carga de interfaz)
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        TaskPaneApp.init();
    } else {
        // Modo preview en navegador sin Excel JS
        TaskPaneApp.init();
    }
});

const TaskPaneApp = {
    draggedElementData: null,

    async init() {
        FilterModal.init();
        this.bindEvents();
        await this.loadFields();
    },

    bindEvents() {
        document.getElementById("btnRefresh").addEventListener("click", () => this.loadFields());
        document.getElementById("fieldSearch").addEventListener("input", (e) => this.filterFields(e.target.value));

        // Configurar Dropzones
        const zones = document.querySelectorAll(".dropzone-box");
        zones.forEach(zone => {
            zone.addEventListener("dragover", (e) => this.handleDragOver(e));
            zone.addEventListener("dragleave", (e) => this.handleDragLeave(e));
            zone.addEventListener("drop", (e) => this.handleDrop(e));
        });
    },

    async loadFields() {
        const container = document.getElementById("availableFieldsContainer");
        container.innerHTML = "Cargando dimensiones...";

        const dimensions = await ExcelService.readDim2Data();
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

            // 1. Añadir Jerarquías primero con icono de árbol
            dim.hierarchies.forEach(hier => {
                const item = this.createFieldElement(dim.dimension, hier, true);
                group.appendChild(item);
            });

            // 2. Añadir Atributos
            dim.attributes.forEach(att => {
                const item = this.createFieldElement(dim.dimension, att, false);
                group.appendChild(item);
            });

            container.appendChild(group);
        });
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
            this.draggedElementData = fieldData;
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

        const zoneId = dropzoneBox.getAttribute("data-zone");
        const targetContent = dropzoneBox.querySelector(".dropzone-content");

        this.addTagToZone(targetContent, this.draggedElementData, zoneId);
        this.draggedElementData = null;
    },

    addTagToZone(container, data, zoneId) {
        // Evitar duplicados en la misma zona
        const existing = Array.from(container.children).find(
            child => child.dataset.fieldName === data.name && child.dataset.dim === data.dim
        );

        if (existing) return;

        const tag = document.createElement("div");
        tag.className = "dropped-tag";
        tag.dataset.dim = data.dim;
        tag.dataset.fieldName = data.name;
        tag.dataset.isHierarchy = data.isHierarchy;

        tag.innerHTML = `
            <span class="dropped-tag-title">${data.name}</span>
            <span class="dropped-tag-remove">&times;</span>
        `;

        // Doble Click abre la ventana modal de filtro
        tag.addEventListener("dblclick", () => {
            FilterModal.open(data);
        });

        // Eliminar del área
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