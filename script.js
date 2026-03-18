const STORAGE_KEY = "listado-alumnos-state-v2";
const TOTAL_PHASES = 5; 

const state = {
  students: [],
  currentIndex: 0,
  currentPhase: 1,
  settings: {
    title: "Listado de alumnos",
    subtitle: "Fotos y nombres",
    perPage: 30,
    columns: 5,
    sortBy: "name",
  },
  save: {
    lastSavedAt: null,
  },
};

let saveTimer = null;

const $ = (id) => document.getElementById(id);

const namesInput = $("namesInput");
const sortByInput = $("sortByInput");
const loadNamesBtn = $("loadNamesBtn");
const sortAgainBtn = $("sortAgainBtn");
const clearAllBtn = $("clearAllBtn");

const pageTitleInput = $("pageTitle");
const pageSubtitleInput = $("pageSubtitle");
const perPageInput = $("perPage");
const columnsInput = $("columns");
const layoutStatus = $("layoutStatus");

const statusBox = $("statusBox");
const photoProgressText = $("photoProgressText");
const photoProgressFill = $("photoProgressFill");
const missingCount = $("missingCount");

const stepBox = $("stepBox");
const currentStudentName = $("currentStudentName");
const dropzone = $("dropzone");
const dropText = $("dropText");
const photoInput = $("photoInput");
const thumbPreview = $("thumbPreview");
const skipBtn = $("skipBtn");
const nextWithoutPhotoBtn = $("nextWithoutPhotoBtn");

const cropModal = $("cropModal");
const cropViewport = $("cropViewport");
const cropImage = $("cropImage");
const cropZoom = $("cropZoom");
const cropCancelBtn = $("cropCancelBtn");
const cropApplyBtn = $("cropApplyBtn");

const editList = $("editList");
const addStudentBtn = $("addStudentBtn");
const exportBtn = $("exportBtn");
const importBtn = $("importBtn");
const importFile = $("importFile");

const printBtn = $("printBtn");
const downloadDocBtn = $("downloadDocBtn");
const backToPhotosBtn = $("backToPhotosBtn");
const resetWorkflowBtn = $("resetWorkflowBtn");
const finalSummary = $("finalSummary");

const phaseLabel = $("phaseLabel");
const phaseSteps = $("phaseSteps");
const saveStatus = $("saveStatus");
const prevPhaseBtn = $("prevPhaseBtn");
const nextPhaseBtn = $("nextPhaseBtn");

const previewArea = $("previewArea");
const previewTitle = $("previewTitle");
const previewHint = $("previewHint");

const cropState = {
  imageElement: null,
  baseScale: 1,
  zoom: 1,
  offsetX: 0,
  offsetY: 0,
  isDragging: false,
  dragStartX: 0,
  dragStartY: 0,
  startOffsetX: 0,
  startOffsetY: 0,
  resolver: null,
};

const phaseTitles = {
  1: "Cargar nombres",
  2: "Configurar A4",
  3: "Asignar fotos",
  4: "Editar listado",
  5: "Revisión final",
};

function normalizeNames(raw) {
  const lines = raw
    .split(/\r?\n/)
    .map((value) => value.trim())
    .filter(Boolean);

  if (lines.length === 0) return [];

  const allLinesAreSingleCommaPair =
    lines.length > 1 &&
    lines.every((line) => {
      const commaCount = (line.match(/,/g) || []).length;
      return commaCount === 1;
    });

  if (allLinesAreSingleCommaPair) {
    return lines
      .map((line) => {
        const [left, right] = line.split(",").map((part) => part.trim());
        if (!left || !right) {
          return line.replace(/,/g, " ").replace(/\s+/g, " ").trim();
        }
        return `${right} ${left}`.replace(/\s+/g, " ").trim();
      })
      .filter(Boolean);
  }

  if (lines.length === 1) {
    return lines[0]
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean);
  }

  return lines
    .flatMap((line) => line.split(","))
    .map((value) => value.trim())
    .filter(Boolean);
}

function createStudent(name) {
  return {
    id: crypto.randomUUID(),
    name,
    photo: null,
  };
}

const nameCollator = new Intl.Collator("es", { sensitivity: "base" });

function getNameParts(value) {
  const fullName = String(value || "").trim();
  const parts = fullName.split(/\s+/).filter(Boolean);
  const surname = parts.length > 1 ? parts[parts.length - 1] : fullName;
  const firstNames = parts.length > 1 ? parts.slice(0, -1).join(" ") : fullName;

  return {
    fullName,
    surname,
    firstNames,
  };
}

function getDisplayName(value) {
  const { fullName, surname, firstNames } = getNameParts(value);

  if (state.settings.sortBy !== "surname") {
    return fullName;
  }

  if (!firstNames || firstNames === surname) {
    return fullName;
  }

  return `${surname}, ${firstNames}`;
}

function getLayoutMetrics() {
  const rowsPerPage = Math.max(
    1,
    Math.ceil(state.settings.perPage / state.settings.columns),
  );

  const pageHeightMm = 297;
  const pagePaddingTopMm = 10;
  const pagePaddingBottomMm = 10;
  const headerBlockMm = 18;
  const rowGapMm = 4;
  const usableHeightMm =
    pageHeightMm - pagePaddingTopMm - pagePaddingBottomMm - headerBlockMm;
  const totalGapMm = Math.max(0, rowsPerPage - 1) * rowGapMm;
  const rowHeightMm = (usableHeightMm - totalGapMm) / rowsPerPage;

  return {
    rowsPerPage,
    rowHeightMm,
    usableHeightMm,
  };
}

function renderLayoutStatus() {
  if (!layoutStatus) return;

  const metrics = getLayoutMetrics();
  layoutStatus.classList.remove("ok", "warn", "alert");

  const rowHeightRounded = Math.round(metrics.rowHeightMm);
  const headline = `A4 útil: ${Math.round(metrics.usableHeightMm)} mm · Filas estimadas: ${metrics.rowsPerPage} · Alto por card: ${rowHeightRounded} mm`;

  if (metrics.rowHeightMm < 28) {
    layoutStatus.classList.add("alert");
    layoutStatus.textContent = `${headline}. Muy ajustado: puede cortar o volver ilegibles fotos y nombres al imprimir.`;
    return;
  }

  if (metrics.rowHeightMm < 34) {
    layoutStatus.classList.add("warn");
    layoutStatus.textContent = `${headline}. Ajuste medio: imprime una prueba para validar legibilidad.`;
    return;
  }

  layoutStatus.classList.add("ok");
  layoutStatus.textContent = `${headline}. Buen margen para imprimir sin cortes.`;
}

function sortStudents() {
  if (state.settings.sortBy === "surname") {
    state.students.sort((a, b) => {
      const aParts = getNameParts(a.name);
      const bParts = getNameParts(b.name);

      const bySurname = nameCollator.compare(aParts.surname, bParts.surname);
      if (bySurname !== 0) return bySurname;

      const byFirstNames = nameCollator.compare(
        aParts.firstNames,
        bParts.firstNames,
      );
      if (byFirstNames !== 0) return byFirstNames;

      return nameCollator.compare(aParts.fullName, bParts.fullName);
    });
    return;
  }

  state.students.sort((a, b) => nameCollator.compare(a.name, b.name));
}

function rebuildCurrentIndex() {
  const nextMissingIndex = state.students.findIndex(
    (student) => !student.photo,
  );
  state.currentIndex =
    nextMissingIndex >= 0 ? nextMissingIndex : state.students.length;
}

function getPhotoStats() {
  const total = state.students.length;
  const withPhoto = state.students.filter((student) =>
    Boolean(student.photo),
  ).length;
  const pending = total - withPhoto;
  return { total, withPhoto, pending };
}

function clampSettings() {
  state.settings.perPage = Math.max(
    1,
    Math.min(60, Number(state.settings.perPage) || 30),
  );
  state.settings.columns = Math.max(
    2,
    Math.min(8, Number(state.settings.columns) || 5),
  );
  if (!["name", "surname"].includes(state.settings.sortBy)) {
    state.settings.sortBy = "name";
  }
  if (!state.settings.title?.trim()) {
    state.settings.title = "Listado de alumnos";
  }
  if (!state.settings.subtitle?.trim()) {
    state.settings.subtitle = "Fotos y nombres";
  }
}

function syncInputsFromState() {
  pageTitleInput.value = state.settings.title;
  pageSubtitleInput.value = state.settings.subtitle;
  perPageInput.value = state.settings.perPage;
  columnsInput.value = state.settings.columns;
  sortByInput.value = state.settings.sortBy;
}

function queueSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveState, 250);
}

function saveState() {
  try {
    const payload = {
      students: state.students,
      currentPhase: state.currentPhase,
      settings: state.settings,
      savedAt: new Date().toISOString(),
    };

    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    state.save.lastSavedAt = payload.savedAt;
    updateSaveStatus();
  } catch {
    saveStatus.textContent = "No se pudo guardar localmente";
  }
}

function loadSavedState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;

    const parsed = JSON.parse(raw);

    if (Array.isArray(parsed.students)) {
      state.students = parsed.students
        .filter((student) => student && typeof student.name === "string")
        .map((student) => ({
          id: student.id || crypto.randomUUID(),
          name: student.name,
          photo: typeof student.photo === "string" ? student.photo : null,
        }));
    }

    if (parsed.settings && typeof parsed.settings === "object") {
      state.settings = {
        title: parsed.settings.title || "Listado de alumnos",
        subtitle: parsed.settings.subtitle || "Fotos y nombres",
        perPage: Number(parsed.settings.perPage) || 30,
        columns: Number(parsed.settings.columns) || 5,
        sortBy: parsed.settings.sortBy || "name",
      };
    }

    if (Number(parsed.currentPhase)) {
      state.currentPhase = Math.max(
        1,
        Math.min(TOTAL_PHASES, Number(parsed.currentPhase)),
      );
    }

    state.save.lastSavedAt = parsed.savedAt || null;
    clampSettings();
    rebuildCurrentIndex();
    namesInput.value = state.students.map((student) => student.name).join("\n");
    syncInputsFromState();
    updateSaveStatus();
  } catch {
    saveStatus.textContent = "No se pudo recuperar el guardado local";
  }
}

function updateSaveStatus() {
  if (!state.save.lastSavedAt) {
    saveStatus.textContent = "Guardado local activo";
    return;
  }

  const date = new Date(state.save.lastSavedAt);
  saveStatus.textContent = `Guardado local: ${date.toLocaleTimeString("es-UY")}`;
}

function setStatus(text) {
  statusBox.textContent = text;
}

function validatePhaseTransition(targetPhase) {
  if (targetPhase <= state.currentPhase) return true;

  if (targetPhase >= 2 && state.students.length === 0) {
    alert("Primero carga al menos un alumno en el paso 1.");
    return false;
  }

  return true;
}

function goToPhase(nextPhase) {
  const target = Math.max(1, Math.min(TOTAL_PHASES, nextPhase));
  if (!validatePhaseTransition(target)) return;

  state.currentPhase = target;
  renderShell();
  queueSave();
}

function nextPhase() {
  goToPhase(state.currentPhase + 1);
}

function prevPhase() {
  goToPhase(state.currentPhase - 1);
}

function updatePhaseIndicator() {
  phaseLabel.textContent = `Paso ${state.currentPhase} de ${TOTAL_PHASES}: ${phaseTitles[state.currentPhase]}`;

  const stepItems = phaseSteps.querySelectorAll("li");
  stepItems.forEach((item) => {
    const step = Number(item.dataset.step);
    item.classList.remove("active", "completed");

    if (step < state.currentPhase) item.classList.add("completed");
    if (step === state.currentPhase) item.classList.add("active");
  });

  prevPhaseBtn.disabled = state.currentPhase === 1;

  if (state.currentPhase === TOTAL_PHASES) {
    nextPhaseBtn.disabled = true;
    nextPhaseBtn.textContent = "Finalizado";
  } else {
    nextPhaseBtn.disabled = false;
    nextPhaseBtn.textContent = "Siguiente";
  }
}

function updateVisiblePhase() {
  document.querySelectorAll(".phase").forEach((phase) => {
    const phaseNumber = Number(phase.dataset.phase);
    phase.classList.toggle("hidden", phaseNumber !== state.currentPhase);
  });
}

function updatePreviewMode() {
  const compact = state.currentPhase < 5;
  previewArea.classList.toggle("compact", compact);
  previewTitle.textContent = compact
    ? "Vista previa compacta"
    : "Vista final para imprimir";
  previewHint.textContent = compact
    ? "Vista rápida mientras avanzas por los pasos. En el paso final se expande a tamaño completo."
    : "Revisa el resultado final y luego imprime en formato A4 vertical.";
}

function renderPhotoProgress() {
  const stats = getPhotoStats();
  const percentage =
    stats.total > 0 ? Math.round((stats.withPhoto / stats.total) * 100) : 0;

  photoProgressText.textContent = `${stats.withPhoto}/${stats.total} fotos cargadas`;
  missingCount.textContent = `${stats.pending} pendientes`;
  photoProgressFill.style.width = `${percentage}%`;
}

function renderGuidedStep() {
  renderPhotoProgress();

  if (state.students.length === 0) {
    stepBox.classList.add("hidden");
    setStatus("Todavía no cargaste alumnos.");
    return;
  }

  if (state.currentIndex >= state.students.length) {
    stepBox.classList.add("hidden");
    setStatus(
      "Ya recorriste todos los alumnos. Puedes pasar al paso 4 para editar o al paso 5 para imprimir.",
    );
    return;
  }

  const student = state.students[state.currentIndex];
  stepBox.classList.remove("hidden");
  setStatus(`Alumno ${state.currentIndex + 1} de ${state.students.length}`);
  currentStudentName.textContent = student.name;
  dropText.textContent = `Arrastra la imagen de ${student.name}`;

  if (student.photo) {
    thumbPreview.src = student.photo;
    thumbPreview.classList.remove("hidden");
  } else {
    thumbPreview.classList.add("hidden");
    thumbPreview.removeAttribute("src");
  }
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

function renderPreview() {
  previewArea.innerHTML = "";

  if (state.students.length === 0) {
    previewArea.innerHTML = `
      <div class="empty-state">
        <h2>Vista previa</h2>
        <p>Carga alumnos para generar la grilla de impresión.</p>
      </div>
    `;
    return;
  }

  const pages = chunkArray(state.students, state.settings.perPage);

  pages.forEach((pageStudents, pageIndex) => {
    const page = document.createElement("section");
    page.className = "page";

    const header = document.createElement("div");
    header.className = "page-header";
    header.innerHTML = `
      <div class="page-title">${escapeHtml(state.settings.title)}</div>
      <div class="page-subtitle">${escapeHtml(state.settings.subtitle)} · Hoja ${pageIndex + 1} de ${pages.length}</div>
    `;

    const grid = document.createElement("div");
    grid.className = "grid";
    grid.style.gridTemplateColumns = `repeat(${state.settings.columns}, 1fr)`;
    const { rowsPerPage } = getLayoutMetrics();
    grid.style.gridTemplateRows = `repeat(${rowsPerPage}, minmax(0, 1fr))`;

    pageStudents.forEach((student) => {
      const displayName = getDisplayName(student.name);

      const card = document.createElement("div");
      card.className = "student-card";
      card.innerHTML = `
        ${
          student.photo
            ? `<img class="student-photo" src="${student.photo}" alt="${escapeHtml(displayName)}" />`
            : `<div class="student-placeholder" aria-hidden="true"></div>`
        }
        <div class="student-name">${escapeHtml(displayName)}</div>
      `;
      grid.appendChild(card);
    });

    page.appendChild(header);
    page.appendChild(grid);
    previewArea.appendChild(page);
  });
}

function renderEditList() {
  editList.innerHTML = "";

  if (state.students.length === 0) {
    editList.innerHTML =
      '<div class="muted small">No hay alumnos cargados.</div>';
    return;
  }

  state.students.forEach((student, index) => {
    const row = document.createElement("div");
    row.className = "edit-row";

    const badge = document.createElement("div");
    badge.className = "index-badge";
    badge.textContent = index + 1;

    const nameField = document.createElement("div");
    nameField.className = "edit-name";

    const input = document.createElement("input");
    input.type = "text";
    input.value = student.name;
    input.addEventListener("change", () => {
      student.name = input.value.trim() || student.name;
      sortStudents();
      rebuildCurrentIndex();
      namesInput.value = state.students.map((item) => item.name).join("\n");
      renderAll();
      queueSave();
    });

    const statusPill = document.createElement("span");
    statusPill.className = `photo-pill ${student.photo ? "ok" : "missing"}`;
    statusPill.textContent = student.photo ? "Con foto" : "Sin foto";

    nameField.appendChild(input);
    nameField.appendChild(statusPill);

    const actions = document.createElement("div");
    actions.className = "toolbar";

    const uploadBtn = document.createElement("button");
    uploadBtn.textContent = student.photo ? "Cambiar foto" : "Subir foto";
    uploadBtn.addEventListener("click", () => openFileForStudent(student.id));

    const removePhotoBtn = document.createElement("button");
    removePhotoBtn.textContent = "Quitar foto";
    removePhotoBtn.disabled = !student.photo;
    removePhotoBtn.addEventListener("click", () => {
      student.photo = null;
      rebuildCurrentIndex();
      renderAll();
      queueSave();
    });

    const removeBtn = document.createElement("button");
    removeBtn.className = "danger";
    removeBtn.textContent = "Eliminar";
    removeBtn.addEventListener("click", () => {
      const ok = confirm(`¿Eliminar a ${student.name}?`);
      if (!ok) return;
      state.students = state.students.filter((item) => item.id !== student.id);
      rebuildCurrentIndex();
      namesInput.value = state.students.map((item) => item.name).join("\n");
      renderAll();
      queueSave();
    });

    actions.appendChild(uploadBtn);
    actions.appendChild(removePhotoBtn);
    actions.appendChild(removeBtn);

    row.appendChild(badge);
    row.appendChild(nameField);
    row.appendChild(actions);
    editList.appendChild(row);
  });
}

function updateFinalSummary() {
  const stats = getPhotoStats();
  finalSummary.textContent = `Total: ${stats.total} alumnos. Con foto: ${stats.withPhoto}. Pendientes: ${stats.pending}.`;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
        return;
      }
      reject(new Error("Formato de imagen no válido"));
    };
    reader.onerror = () => reject(new Error("No se pudo leer la imagen"));
    reader.readAsDataURL(file);
  });
}

function clampCropOffsets() {
  if (!cropState.imageElement) return;

  const viewportRect = cropViewport.getBoundingClientRect();
  const viewportWidth = viewportRect.width;
  const viewportHeight = viewportRect.height;

  const imageWidth =
    cropState.imageElement.naturalWidth * cropState.baseScale * cropState.zoom;
  const imageHeight =
    cropState.imageElement.naturalHeight * cropState.baseScale * cropState.zoom;

  const minX = Math.min(0, viewportWidth - imageWidth);
  const minY = Math.min(0, viewportHeight - imageHeight);

  cropState.offsetX = Math.max(minX, Math.min(0, cropState.offsetX));
  cropState.offsetY = Math.max(minY, Math.min(0, cropState.offsetY));
}

function renderCropImage() {
  if (!cropState.imageElement) return;

  const width =
    cropState.imageElement.naturalWidth * cropState.baseScale * cropState.zoom;
  const height =
    cropState.imageElement.naturalHeight * cropState.baseScale * cropState.zoom;

  clampCropOffsets();

  cropImage.style.width = `${width}px`;
  cropImage.style.height = `${height}px`;
  cropImage.style.left = `${cropState.offsetX}px`;
  cropImage.style.top = `${cropState.offsetY}px`;
}

function buildCroppedImageDataUrl() {
  if (!cropState.imageElement) return null;

  const viewportRect = cropViewport.getBoundingClientRect();
  const viewportWidth = viewportRect.width;
  const viewportHeight = viewportRect.height;
  const displayScale = cropState.baseScale * cropState.zoom;

  let sx = -cropState.offsetX / displayScale;
  let sy = -cropState.offsetY / displayScale;
  let sw = viewportWidth / displayScale;
  let sh = viewportHeight / displayScale;

  sw = Math.min(sw, cropState.imageElement.naturalWidth);
  sh = Math.min(sh, cropState.imageElement.naturalHeight);
  sx = Math.max(0, Math.min(cropState.imageElement.naturalWidth - sw, sx));
  sy = Math.max(0, Math.min(cropState.imageElement.naturalHeight - sh, sy));

  const outputWidth = 600;
  const outputHeight = 800;
  const canvas = document.createElement("canvas");
  canvas.width = outputWidth;
  canvas.height = outputHeight;

  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  ctx.drawImage(
    cropState.imageElement,
    sx,
    sy,
    sw,
    sh,
    0,
    0,
    outputWidth,
    outputHeight,
  );
  return canvas.toDataURL("image/jpeg", 0.92);
}

function finishCrop(result) {
  if (cropModal.open) cropModal.close();
  const resolver = cropState.resolver;
  cropState.resolver = null;
  cropState.imageElement = null;
  cropImage.removeAttribute("src");
  if (resolver) resolver(result);
}

function handleCropPointerDown(event) {
  if (!cropState.imageElement) return;
  cropState.isDragging = true;
  cropState.dragStartX = event.clientX;
  cropState.dragStartY = event.clientY;
  cropState.startOffsetX = cropState.offsetX;
  cropState.startOffsetY = cropState.offsetY;
  cropViewport.setPointerCapture(event.pointerId);
}

function handleCropPointerMove(event) {
  if (!cropState.isDragging) return;

  cropState.offsetX =
    cropState.startOffsetX + (event.clientX - cropState.dragStartX);
  cropState.offsetY =
    cropState.startOffsetY + (event.clientY - cropState.dragStartY);
  renderCropImage();
}

function handleCropPointerUp(event) {
  cropState.isDragging = false;
  if (cropViewport.hasPointerCapture(event.pointerId)) {
    cropViewport.releasePointerCapture(event.pointerId);
  }
}

async function openCropTool(file) {
  try {
    const src = await readFileAsDataUrl(file);

    return await new Promise((resolve) => {
      const image = new Image();

      image.onload = () => {
        cropState.imageElement = image;
        cropState.zoom = 1;
        cropZoom.value = "1";

        if (!cropModal.open) cropModal.showModal();

        const viewportRect = cropViewport.getBoundingClientRect();
        const widthScale = viewportRect.width / image.naturalWidth;
        const heightScale = viewportRect.height / image.naturalHeight;
        cropState.baseScale = Math.max(widthScale, heightScale);

        const renderedWidth = image.naturalWidth * cropState.baseScale;
        const renderedHeight = image.naturalHeight * cropState.baseScale;

        cropState.offsetX = (viewportRect.width - renderedWidth) / 2;
        cropState.offsetY = (viewportRect.height - renderedHeight) / 2;

        cropImage.src = src;
        cropState.resolver = resolve;
        renderCropImage();
      };

      image.onerror = () => resolve(null);
      image.src = src;
    });
  } catch {
    return null;
  }
}

function renderShell() {
  updatePhaseIndicator();
  updateVisiblePhase();
  updatePreviewMode();
  updateFinalSummary();
}

function renderAll() {
  clampSettings();
  syncInputsFromState();
  renderShell();
  renderLayoutStatus();
  renderGuidedStep();
  renderEditList();
  renderPreview();
}

function loadNames() {
  const names = normalizeNames(namesInput.value);
  state.students = names.map(createStudent);
  sortStudents();
  rebuildCurrentIndex();
  renderAll();
  queueSave();
}

function resetWorkflow() {
  const ok = confirm(
    "Esto borrará alumnos, fotos y configuración del proceso actual. ¿Quieres empezar de cero?",
  );
  if (!ok) return;

  state.students = [];
  state.currentIndex = 0;
  state.currentPhase = 1;
  state.settings = {
    title: "Listado de alumnos",
    subtitle: "Fotos y nombres",
    perPage: 30,
    columns: 5,
    sortBy: "name",
  };

  namesInput.value = "";
  if (cropModal.open) cropModal.close();

  renderAll();
  queueSave();
}

function clearAll() {
  resetWorkflow();
}

function applySettingsLive() {
  state.settings.title = pageTitleInput.value.trim() || "Listado de alumnos";
  state.settings.subtitle = pageSubtitleInput.value.trim() || "Fotos y nombres";
  state.settings.perPage = Math.max(
    1,
    Math.min(60, Number(perPageInput.value) || 30),
  );
  state.settings.columns = Math.max(
    2,
    Math.min(8, Number(columnsInput.value) || 5),
  );
  renderLayoutStatus();
  renderPreview();
  updateFinalSummary();
  queueSave();
}

async function handleImageFile(file, forcedStudentId = null) {
  if (!file?.type?.startsWith("image/")) {
    alert("Selecciona una imagen válida.");
    return;
  }

  const croppedData = await openCropTool(file);
  if (!croppedData) return;

  if (forcedStudentId) {
    const forcedStudent = state.students.find(
      (student) => student.id === forcedStudentId,
    );
    if (forcedStudent) forcedStudent.photo = croppedData;
    rebuildCurrentIndex();
    renderAll();
    queueSave();
    return;
  }

  const student = state.students[state.currentIndex];
  if (!student) return;

  student.photo = croppedData;
  thumbPreview.src = croppedData;
  thumbPreview.classList.remove("hidden");
  state.currentIndex += 1;

  renderAll();
  queueSave();
}

function openFileForStudent(studentId) {
  const tempInput = document.createElement("input");
  tempInput.type = "file";
  tempInput.accept = "image/*";
  tempInput.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) void handleImageFile(file, studentId);
  });
  tempInput.click();
}

function skipCurrent() {
  if (state.currentIndex < state.students.length) {
    state.currentIndex += 1;
    renderGuidedStep();
  }
}

function addStudent() {
  const name = prompt("Nombre del nuevo alumno:");
  if (!name?.trim()) return;

  state.students.push(createStudent(name.trim()));
  sortStudents();
  rebuildCurrentIndex();
  namesInput.value = state.students.map((student) => student.name).join("\n");
  renderAll();
  queueSave();
}

function exportData() {
  const payload = {
    exportedAt: new Date().toISOString(),
    settings: state.settings,
    currentPhase: state.currentPhase,
    students: state.students,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "listado-alumnos.json";
  anchor.click();
  URL.revokeObjectURL(url);
}

function buildDocHtml() {
  const pages = chunkArray(state.students, state.settings.perPage);

  const pagesHtml = pages
    .map((pageStudents, pageIndex) => {
      const rows = [];

      for (let i = 0; i < pageStudents.length; i += state.settings.columns) {
        const rowStudents = pageStudents.slice(i, i + state.settings.columns);
        const cells = rowStudents
          .map((student) => {
            const photoHtml = student.photo
              ? `<img src="${student.photo}" alt="${escapeHtml(student.name)}" style="width:100%;aspect-ratio:3/4;object-fit:cover;border:1px solid #d1d9e6;border-radius:6px;" />`
              : '<div style="width:100%;aspect-ratio:3/4;border:1px dashed #c3cfdf;border-radius:6px;background:#f5f8fc;"></div>';

            return `
              <td style="width:${100 / state.settings.columns}%;vertical-align:top;padding:8px;">
                ${photoHtml}
                <div style="margin-top:6px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;text-align:center;word-break:break-word;">${escapeHtml(student.name)}</div>
              </td>
            `;
          })
          .join("");

        const emptyCellsCount = Math.max(0, state.settings.columns - rowStudents.length);
        const emptyCells = Array.from({ length: emptyCellsCount })
          .map(
            () =>
              `<td style="width:${100 / state.settings.columns}%;vertical-align:top;padding:8px;"></td>`,
          )
          .join("");

        rows.push(`<tr>${cells}${emptyCells}</tr>`);
      }

      return `
        <section style="page-break-after:always;min-height:1000px;">
          <h2 style="margin:0 0 4px;text-align:center;font-family:Arial,sans-serif;">${escapeHtml(state.settings.title)}</h2>
          <p style="margin:0 0 14px;text-align:center;color:#4d637a;font-family:Arial,sans-serif;font-size:12px;">${escapeHtml(state.settings.subtitle)} · Hoja ${pageIndex + 1} de ${pages.length}</p>
          <table style="width:100%;border-collapse:collapse;table-layout:fixed;">
            <tbody>
              ${rows.join("")}
            </tbody>
          </table>
        </section>
      `;
    })
    .join("");

  return `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8" />
      <title>${escapeHtml(state.settings.title)}</title>
    </head>
    <body style="margin:20px;background:#fff;">
      ${pagesHtml || '<p style="font-family:Arial,sans-serif;">No hay alumnos cargados.</p>'}
    </body>
    </html>
  `;
}

function downloadDoc() {
  const html = buildDocHtml();
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "listado-alumnos.doc";
  anchor.click();
  URL.revokeObjectURL(url);
}

async function importDataFromFile(file) {
  try {
    const text = await file.text();
    const data = JSON.parse(text);

    state.students = Array.isArray(data.students)
      ? data.students
          .filter((student) => student && typeof student.name === "string")
          .map((student) => ({
            id: student.id || crypto.randomUUID(),
            name: student.name,
            photo: typeof student.photo === "string" ? student.photo : null,
          }))
      : [];

    state.settings = {
      title: data.settings?.title || "Listado de alumnos",
      subtitle: data.settings?.subtitle || "Fotos y nombres",
      perPage: Number(data.settings?.perPage) || 30,
      columns: Number(data.settings?.columns) || 5,
      sortBy: data.settings?.sortBy || "name",
    };

    state.currentPhase = Math.max(
      1,
      Math.min(TOTAL_PHASES, Number(data.currentPhase) || 1),
    );

    clampSettings();
    rebuildCurrentIndex();
    namesInput.value = state.students.map((student) => student.name).join("\n");
    renderAll();
    queueSave();
  } catch {
    alert("El archivo no tiene un formato válido.");
  }
}

loadSavedState();

loadNamesBtn.addEventListener("click", loadNames);

sortAgainBtn.addEventListener("click", () => {
  sortStudents();
  rebuildCurrentIndex();
  namesInput.value = state.students.map((student) => student.name).join("\n");
  renderAll();
  queueSave();
});

clearAllBtn.addEventListener("click", clearAll);

pageTitleInput.addEventListener("input", applySettingsLive);
pageSubtitleInput.addEventListener("input", applySettingsLive);
perPageInput.addEventListener("input", applySettingsLive);
columnsInput.addEventListener("input", applySettingsLive);

sortByInput.addEventListener("change", () => {
  state.settings.sortBy = sortByInput.value;
  sortStudents();
  rebuildCurrentIndex();
  namesInput.value = state.students.map((student) => student.name).join("\n");
  renderAll();
  queueSave();
});

prevPhaseBtn.addEventListener("click", prevPhase);
nextPhaseBtn.addEventListener("click", nextPhase);

dropzone.addEventListener("click", () => photoInput.click());

photoInput.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (file) void handleImageFile(file);
  event.target.value = "";
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragover");
  const file = event.dataTransfer.files?.[0];
  if (file) void handleImageFile(file);
});

cropViewport.addEventListener("pointerdown", handleCropPointerDown);
cropViewport.addEventListener("pointermove", handleCropPointerMove);
cropViewport.addEventListener("pointerup", handleCropPointerUp);
cropViewport.addEventListener("pointercancel", handleCropPointerUp);

cropZoom.addEventListener("input", () => {
  cropState.zoom = Number(cropZoom.value) || 1;
  renderCropImage();
});

cropCancelBtn.addEventListener("click", () => finishCrop(null));
cropApplyBtn.addEventListener("click", () =>
  finishCrop(buildCroppedImageDataUrl()),
);

cropModal.addEventListener("click", (event) => {
  if (event.target === cropModal) finishCrop(null);
});

skipBtn.addEventListener("click", skipCurrent);
nextWithoutPhotoBtn.addEventListener("click", skipCurrent);

addStudentBtn.addEventListener("click", addStudent);
exportBtn.addEventListener("click", exportData);
importBtn.addEventListener("click", () => importFile.click());

importFile.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (file) void importDataFromFile(file);
  event.target.value = "";
});

printBtn.addEventListener("click", () => globalThis.print());
downloadDocBtn.addEventListener("click", downloadDoc);

backToPhotosBtn.addEventListener("click", () => {
  rebuildCurrentIndex();
  goToPhase(3);
});

resetWorkflowBtn.addEventListener("click", resetWorkflow);

renderAll();
