/**
 * BucketPickerUI
 * ------------------------------------------------------------------------
 * Selector reutilizable "Proyecto de Google Cloud -> Bucket", usado tanto
 * en el diálogo "Abrir desde bucket" (bucketBrowser.html) como en el
 * diálogo "Guardar en bucket" (saveBucket.html).
 *
 * Antes este selector vivía dentro de la pantalla de conexión (login.html,
 * campo "Bucket de exportación"). Ahora se elige el proyecto/bucket en el
 * propio momento de abrir o guardar, no al configurar la conexión.
 *
 * Requiere que ya estén cargados js/bigquery.js y js/gcsExport.js.
 *
 * Uso:
 *   const picker = BucketPickerUI.mount({
 *       projectSelect: document.getElementById("pickerProject"),
 *       bucketSelect: document.getElementById("pickerBucket"),
 *       statusEl: document.getElementById("pickerStatus"),
 *       onBucketChange(bucketName) { ... }
 *   });
 *   await picker.init();
 */
(function () {

    function escapeHtml(str) {
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function mount({ projectSelect, bucketSelect, statusEl, onBucketChange }) {

        function setStatus(msg, isError) {
            if (!statusEl) return;
            statusEl.textContent = msg || "";
            statusEl.classList.toggle("error", !!isError);
        }

        async function loadProjects() {
            if (!window.BQ || !BQ.isConnected()) {
                projectSelect.innerHTML = `<option value="">— Conéctate primero —</option>`;
                projectSelect.disabled = true;
                bucketSelect.innerHTML = `<option value="">—</option>`;
                bucketSelect.disabled = true;
                setStatus("Conéctate desde \"Conexiones\" con una cuenta de Google/BigQuery para poder ver tus proyectos y buckets.", true);
                return;
            }

            projectSelect.innerHTML = `<option value="">Cargando proyectos…</option>`;
            projectSelect.disabled = true;
            setStatus("");

            try {
                const projects = await BQ.listProjects();
                if (!projects.length) {
                    projectSelect.innerHTML = `<option value="">No se encontraron proyectos</option>`;
                    setStatus("No se encontró ningún proyecto de Google Cloud visible para esta cuenta.", true);
                    return;
                }

                const lastProject = BQ.getExportProject();
                projectSelect.innerHTML = `<option value="">Selecciona un proyecto…</option>` +
                    projects.map(p => {
                        const id = p.id || (p.projectReference && p.projectReference.projectId);
                        const name = p.friendlyName || id;
                        const label = name !== id ? `${name} (${id})` : id;
                        const selected = id === lastProject ? " selected" : "";
                        return `<option value="${escapeHtml(id)}"${selected}>${escapeHtml(label)}</option>`;
                    }).join("");
                projectSelect.disabled = false;

                if (lastProject && projects.some(p => (p.id || (p.projectReference && p.projectReference.projectId)) === lastProject)) {
                    await loadBucketsForProject(lastProject, BQ.getExportBucket());
                }
            } catch (err) {
                console.error("Error al listar proyectos de Google Cloud:", err);
                projectSelect.innerHTML = `<option value="">Error al cargar proyectos</option>`;
                setStatus("Error al listar proyectos: " + (err.message || err), true);
            }
        }

        async function loadBucketsForProject(projectId, preselectBucket) {
            if (!projectId) {
                bucketSelect.innerHTML = `<option value="">Selecciona primero un proyecto…</option>`;
                bucketSelect.disabled = true;
                return;
            }

            bucketSelect.innerHTML = `<option value="">Cargando buckets…</option>`;
            bucketSelect.disabled = true;
            setStatus("");

            try {
                const buckets = await GCS.listBuckets(projectId);
                if (!buckets.length) {
                    bucketSelect.innerHTML = `<option value="">Este proyecto no tiene buckets</option>`;
                    setStatus("Este proyecto no tiene ningún bucket de Cloud Storage (o no tienes permiso para verlos).", true);
                    return;
                }
                bucketSelect.innerHTML = `<option value="">Selecciona un bucket…</option>` +
                    buckets.map(name => {
                        const selected = name === preselectBucket ? " selected" : "";
                        return `<option value="${escapeHtml(name)}"${selected}>${escapeHtml(name)}</option>`;
                    }).join("");
                bucketSelect.disabled = false;

                if (preselectBucket && buckets.includes(preselectBucket) && typeof onBucketChange === "function") {
                    onBucketChange(preselectBucket);
                }
            } catch (err) {
                console.error("Error al listar buckets:", err);
                bucketSelect.innerHTML = `<option value="">Error al cargar buckets</option>`;
                setStatus("Error al listar buckets: " + (err.message || err), true);
            }
        }

        projectSelect.addEventListener("change", (e) => {
            const projectId = e.target.value;
            BQ.setExportProject(projectId || "");
            loadBucketsForProject(projectId, "");
        });

        bucketSelect.addEventListener("change", (e) => {
            const bucket = e.target.value;
            BQ.setExportBucket(bucket || "");
            if (bucket && typeof onBucketChange === "function") onBucketChange(bucket);
        });

        return {
            init: loadProjects,
            reload: loadProjects
        };
    }

    window.BucketPickerUI = { mount };

})();
