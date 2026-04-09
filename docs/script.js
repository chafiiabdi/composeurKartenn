const board = document.getElementById("board");
const boardGrid = document.getElementById("boardGrid");
const mapLayer = document.getElementById("mapLayer");
const mapImage = document.getElementById("mapImage");
const emptyMapState = document.getElementById("emptyMapState");
const mapUpload = document.getElementById("mapUpload");
const customPictoUpload = document.getElementById("customPictoUpload");
const uploadStatus = document.getElementById("uploadStatus");
const favoritePictos = document.getElementById("favoritePictos");
const legendSections = document.getElementById("legendSections");
const legendHint = document.getElementById("legendHint");
const legendPanel = document.querySelector(".carto-legend-panel");
const legendResizer = document.getElementById("legendResizer");
const legendAddSectionButton = document.getElementById("legendAddSection");
const legendResetButton = document.getElementById("legendReset");
const legendMoveUpButton = document.getElementById("legendMoveUp");
const legendMoveDownButton = document.getElementById("legendMoveDown");
const legendRenameSectionButton = document.getElementById("legendRenameSection");
const stageLayout = document.querySelector(".carto-stage-layout");
const mainContent = document.getElementById("main-content");
const layersList = document.getElementById("layersList");
const layerCount = document.getElementById("layerCount");
const zoomValue = document.getElementById("zoomValue");
const titleInput = document.getElementById("titleInput");
const subtitleInput = document.getElementById("subtitleInput");
const sourceInput = document.getElementById("sourceInput");
const northLabelInput = document.getElementById("northLabelInput");
const boardTitle = document.getElementById("boardTitle");
const boardSource = document.getElementById("boardSource");
const northLabel = document.getElementById("northLabel");
const selectedText = document.getElementById("selectedText");
const selectedFontSize = document.getElementById("selectedFontSize");
const selectedTextColor = document.getElementById("selectedTextColor");
const selectedTextBackground = document.getElementById("selectedTextBackground");
const selectedFontFamily = document.getElementById("selectedFontFamily");
const toggleBoldButton = document.getElementById("toggleBold");
const toggleItalicButton = document.getElementById("toggleItalic");
const toggleTransparentBackgroundButton = document.getElementById("toggleTransparentBackground");
const legendSelectionHint = document.getElementById("legendSelectionHint");
const legendSelectedText = document.getElementById("legendSelectedText");
const legendSelectedFontSize = document.getElementById("legendSelectedFontSize");
const legendSelectedTextColor = document.getElementById("legendSelectedTextColor");
const legendSelectedTextBackground = document.getElementById("legendSelectedTextBackground");
const legendSelectedFontFamily = document.getElementById("legendSelectedFontFamily");
const legendToggleBoldButton = document.getElementById("legendToggleBold");
const legendToggleItalicButton = document.getElementById("legendToggleItalic");
const legendToggleTransparentBackgroundButton = document.getElementById("legendToggleTransparentBackground");
const selectedWidth = document.getElementById("selectedWidth");
const selectedHeight = document.getElementById("selectedHeight");
const selectedRotation = document.getElementById("selectedRotation");
const legendColumns = document.getElementById("legendColumns");
const legendSymbolSize = document.getElementById("legendSymbolSize");
const selectionHint = document.getElementById("selectionHint");
const bringForwardButton = document.getElementById("bringForward");
const sendBackwardButton = document.getElementById("sendBackward");
const deleteLayerButton = document.getElementById("deleteLayer");
const exportImageButton = document.getElementById("exportImage");
const resetWorkspaceButton = document.getElementById("resetWorkspace");
const toggleGridButton = document.getElementById("toggleGrid");
const zoomInButton = document.getElementById("zoomIn");
const zoomOutButton = document.getElementById("zoomOut");
const openDisplaySettingsButton = document.getElementById("openDisplaySettings");
const closeDisplaySettingsButton = document.getElementById("closeDisplaySettings");
const displaySettings = document.getElementById("displaySettings");
const displaySettingsBackdrop = document.getElementById("displaySettingsBackdrop");
const addButtons = document.querySelectorAll("[data-add]");
const shapeButtons = document.querySelectorAll("[data-shape]");
const pictoLibrary = document.getElementById("pictoLibrary");
const toolboxSearch = document.getElementById("toolboxSearch");
const toolboxMenuButtons = document.querySelectorAll("[data-toolbox-target]");
const toolboxPanels = document.querySelectorAll(".carto-toolbox__panel");
const sideMenuButtons = document.querySelectorAll("[data-side-target]");
const sidePanels = document.querySelectorAll(".carto-side-panel");
const leftPanel = document.getElementById("leftPanel");
const leftPanelToggle = document.getElementById("leftPanelToggle");
const themeInputs = document.querySelectorAll('input[name="displayTheme"]');

const dsfrVersion = "1.14.4";
const dsfrArtworkBase = `https://cdn.jsdelivr.net/npm/@gouvfr/dsfr@${dsfrVersion}/dist/artwork/pictograms`;

const pictoCatalog = {
  administration: { label: "Service public", src: `${dsfrArtworkBase}/buildings/city-hall.svg` },
  information: { label: "Information", src: `${dsfrArtworkBase}/system/system.svg` },
  alerte: { label: "Alerte", src: `${dsfrArtworkBase}/system/technical-error.svg` },
  environnement: { label: "Environnement", src: `${dsfrArtworkBase}/environment/sun.svg` },
  nuit: { label: "Observation", src: `${dsfrArtworkBase}/environment/moon.svg` },
  territoire: { label: "Territoire", src: `${dsfrArtworkBase}/system/system.svg` },
};

const favoritePictoKeys = ["administration", "information", "alerte", "territoire"];
const shapeCatalog = {
  rectangle: { label: "Rectangle" },
  ellipse: { label: "Ellipse" },
  triangle: { label: "Triangle" },
  diamond: { label: "Losange" },
  line: { label: "Ligne" },
  "arrow-right": { label: "Fleche" },
  "black-bar": { label: "Trait noir" },
  "bubble-left": { label: "Bulle gauche" },
  "bubble-bottom": { label: "Bulle bas" },
  "callout-rect-left": { label: "Encadre gauche" },
  "callout-rect-right": { label: "Encadre droite" },
  "notch-top": { label: "Encoche haut" },
  "notch-bottom": { label: "Encoche bas" },
};

const layers = [
  { id: "map-base", type: "map", label: "Carte importee", locked: false, visible: true, listed: true },
  { id: "title-base", type: "title", label: "Titre principal", locked: false, visible: true, listed: true },
  { id: "source-base", type: "source", label: "Mention source", locked: false, visible: false, listed: false },
  { id: "north-base", type: "north", label: "Orientation", locked: false, visible: false, listed: false },
];

const ZOOM_BASELINE = 1.8; // Old 180% becomes new 100%
const ZOOM_MIN = 0.6 * ZOOM_BASELINE;
const ZOOM_MAX = 1.8 * ZOOM_BASELINE;
const ZOOM_STEP = 0.1 * ZOOM_BASELINE;
let zoomLevel = ZOOM_BASELINE;
let selectedLayerId = "title-base";
let selectedLayerIds = new Set(["title-base"]);
let interaction = null;
let mapImageDataUrl = "";
let customIndex = 0;
let copiedLayerPayload = null;
let historyStack = [];
let isRestoringHistory = false;
let suppressSelectionClickForLayerId = "";
const themeStorageKey = "carto-display-theme";
let legendSectionCounter = 0;
let legendSelection = null;
let draggedLegendItem = null;
let legendSectionsState = [];
let mapFrameSyncRaf = 0;
let legendResizeInteraction = null;
const collapsedLayerGroups = new Set();
const EXPORT_BASE_WIDTH = 790;
const EXPORT_BASE_HEIGHT = 506;
const EXPORT_OUTPUT_SCALE = 2;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function getElementByLayerId(layerId) {
  return board.querySelector(`[data-layer-id="${layerId}"]`);
}

function getLayer(layerId) {
  return layers.find((layer) => layer.id === layerId);
}

function getLayerIndex(layerId) {
  return layers.findIndex((layer) => layer.id === layerId);
}

function isLayerListed(layer) {
  return layer?.listed !== false;
}

function getNextLayerId(prefix) {
  customIndex += 1;
  return `${prefix}-${customIndex}`;
}

function getLayerMinSize(layer) {
  if (!layer) {
    return 1;
  }

  if (layer.type === "picto") {
    return 1;
  }

  return 1;
}

function getLegendWidthPercent() {
  const inlineWidth = parseFloat(legendPanel?.style?.width || "");
  if (Number.isFinite(inlineWidth)) {
    return clamp(inlineWidth, 10, 60);
  }

  const boardWidth = board?.clientWidth || 0;
  const legendWidthPx = legendPanel?.clientWidth || 0;
  if (!boardWidth || !legendWidthPx) {
    return 26;
  }

  return clamp((legendWidthPx / boardWidth) * 100, 10, 60);
}

function setLegendWidthPercent(value) {
  const next = clamp(Number(value) || 26, 10, 60);
  if (legendPanel) {
    legendPanel.style.width = `${next}%`;
  }
  if (legendResizer) {
    legendResizer.style.right = `calc(${next}% - 4px)`;
  }
}

function getMapWorkAreaWidthPercent() {
  return 100 - getLegendWidthPercent();
}

function resetMapLayerFrame() {
  const legendWidthPercent = getLegendWidthPercent();
  setLegendWidthPercent(legendWidthPercent);
  mapLayer.style.left = "0%";
  mapLayer.style.top = "0%";
  mapLayer.style.width = `${100 - legendWidthPercent}%`;
  mapLayer.style.height = "100%";
}

function scheduleMapLayerFrameSync() {
  if (mapFrameSyncRaf) {
    cancelAnimationFrame(mapFrameSyncRaf);
  }

  mapFrameSyncRaf = requestAnimationFrame(() => {
    resetMapLayerFrame();
    mapFrameSyncRaf = 0;
  });
}

function startLegendResize(event) {
  if (!legendResizer || !legendPanel || event.button !== 0) {
    return;
  }

  pushHistory();
  event.preventDefault();
  const boardRect = board.getBoundingClientRect();
  legendResizeInteraction = {
    pointerId: event.pointerId,
    boardLeft: boardRect.left,
    boardWidth: boardRect.width,
  };
  legendResizer.classList.add("is-active");
  legendResizer.setPointerCapture(event.pointerId);
}

function updateLegendResize(event) {
  if (!legendResizeInteraction) {
    return false;
  }

  const relativeX = clamp(event.clientX - legendResizeInteraction.boardLeft, 0, legendResizeInteraction.boardWidth);
  const legendWidthPercent = ((legendResizeInteraction.boardWidth - relativeX) / legendResizeInteraction.boardWidth) * 100;
  setLegendWidthPercent(legendWidthPercent);
  resetMapLayerFrame();
  return true;
}

function endLegendResize(event) {
  if (!legendResizeInteraction || !legendResizer) {
    return;
  }

  if (event?.pointerId !== undefined && legendResizer.hasPointerCapture(event.pointerId)) {
    legendResizer.releasePointerCapture(event.pointerId);
  }
  legendResizer.classList.remove("is-active");
  legendResizeInteraction = null;
  refreshSelectedControls();
}

function setActiveToolboxPanel(panelId) {
  if (!panelId) {
    return;
  }

  toolboxPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.id === panelId);
  });

  toolboxMenuButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.toolboxTarget === panelId);
  });
}

function initializeToolboxMenu() {
  toolboxMenuButtons.forEach((button) => {
    button.addEventListener("click", () => {
      setActiveToolboxPanel(button.dataset.toolboxTarget);
    });
  });

  if (toolboxSearch) {
    toolboxSearch.addEventListener("input", (event) => {
      const query = event.target.value.trim().toLowerCase();
      let firstVisibleButton = null;

      toolboxMenuButtons.forEach((button) => {
        const isMatch = button.textContent.toLowerCase().includes(query);
        button.classList.toggle("is-filtered-out", !isMatch);
        if (isMatch && !firstVisibleButton) {
          firstVisibleButton = button;
        }
      });

      const activeButton = [...toolboxMenuButtons].find((button) => button.classList.contains("is-active") && !button.classList.contains("is-filtered-out"));
      if (activeButton) {
        return;
      }

      if (firstVisibleButton) {
        setActiveToolboxPanel(firstVisibleButton.dataset.toolboxTarget);
      }
    });
  }

  const initiallyActiveButton = [...toolboxMenuButtons].find((button) => button.classList.contains("is-active"));
  setActiveToolboxPanel(initiallyActiveButton?.dataset.toolboxTarget || toolboxPanels[0]?.id || "");
}

function setActiveSidePanel(panelId) {
  if (!panelId) {
    return;
  }

  sidePanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.id === panelId);
  });

  sideMenuButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.sideTarget === panelId);
  });
}

function initializeSideMenu() {
  const rightPanel = document.querySelector(".carto-panel--right");
  rightPanel?.classList.add("js-side-panels");

  sideMenuButtons.forEach((button) => {
    button.addEventListener("click", () => {
      setActiveSidePanel(button.dataset.sideTarget);
    });
  });

  const initiallyActiveButton = [...sideMenuButtons].find((button) => button.classList.contains("is-active"));
  setActiveSidePanel(initiallyActiveButton?.dataset.sideTarget || sidePanels[0]?.id || "");
}

function updateLeftPanelToggle() {
  if (!leftPanel || !leftPanelToggle) {
    return;
  }

  const isCollapsed = leftPanel.classList.contains("is-collapsed");
  leftPanelToggle.setAttribute("aria-expanded", String(!isCollapsed));
  leftPanelToggle.setAttribute("aria-label", isCollapsed ? "Ouvrir la boite a outils" : "Replier la boite a outils");
  mainContent?.classList.toggle("is-left-panel-collapsed", isCollapsed);
}

function initializeLeftPanelToggle() {
  if (!leftPanel || !leftPanelToggle) {
    return;
  }

  updateLeftPanelToggle();
  leftPanelToggle.addEventListener("click", () => {
    leftPanel.classList.toggle("is-collapsed");
    updateLeftPanelToggle();
    scheduleMapLayerFrameSync();
  });
}

function applyDisplayTheme(theme) {
  const safeTheme = ["light", "dark", "system"].includes(theme) ? theme : "light";
  document.documentElement.setAttribute("data-fr-scheme", safeTheme);
  localStorage.setItem(themeStorageKey, safeTheme);
  themeInputs.forEach((input) => {
    input.checked = input.value === safeTheme;
  });
}

function openDisplaySettings() {
  if (!displaySettings) {
    return;
  }

  displaySettings.hidden = false;
}

function closeDisplaySettings() {
  if (!displaySettings) {
    return;
  }

  displaySettings.hidden = true;
}

function initializeDisplaySettings() {
  const storedTheme = localStorage.getItem(themeStorageKey) || document.documentElement.getAttribute("data-fr-scheme") || "light";
  applyDisplayTheme(storedTheme);

  openDisplaySettingsButton?.addEventListener("click", openDisplaySettings);
  closeDisplaySettingsButton?.addEventListener("click", closeDisplaySettings);
  displaySettingsBackdrop?.addEventListener("click", closeDisplaySettings);

  themeInputs.forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) {
        applyDisplayTheme(input.value);
      }
    });
  });
}

function initializeLegendToolbar() {
  legendAddSectionButton?.addEventListener("click", () => {
    const sectionName = prompt("Nom du sous-bloc de legende :", "Nouveau sous-bloc");
    if (!sectionName) {
      return;
    }
    legendSectionsState.push(createLegendSection(sectionName.trim() || "Nouveau sous-bloc", "custom"));
    renderLegend();
  });

  legendResetButton?.addEventListener("click", () => {
    legendSectionsState = [];
    legendSelection = null;
    renderLegend();
    refreshSelectedControls();
  });

  legendMoveUpButton?.addEventListener("click", () => {
    if (!legendSelection?.sectionId) {
      return;
    }
    const groupIndex = legendSectionsState.findIndex((group) => group.id === legendSelection.sectionId);
    if (groupIndex <= 0) {
      return;
    }
    const [sectionState] = legendSectionsState.splice(groupIndex, 1);
    legendSectionsState.splice(groupIndex - 1, 0, sectionState);
    renderLegend();
    refreshSelectedControls();
  });

  legendMoveDownButton?.addEventListener("click", () => {
    if (!legendSelection?.sectionId) {
      return;
    }
    const groupIndex = legendSectionsState.findIndex((group) => group.id === legendSelection.sectionId);
    if (groupIndex === -1 || groupIndex >= legendSectionsState.length - 1) {
      return;
    }
    const [sectionState] = legendSectionsState.splice(groupIndex, 1);
    legendSectionsState.splice(groupIndex + 1, 0, sectionState);
    renderLegend();
    refreshSelectedControls();
  });

  legendRenameSectionButton?.addEventListener("click", () => {
    if (!legendSelection?.sectionId) {
      return;
    }
    const section = legendSectionsState.find((group) => group.id === legendSelection.sectionId);
    if (!section) {
      return;
    }
    const nextName = prompt("Nom du sous-bloc :", section.title);
    if (!nextName) {
      return;
    }
    section.title = nextName.trim() || section.title;
    renderLegend();
    refreshSelectedControls();
  });
}


function isRotatableLayer(layer) {
  return Boolean(layer && ["title", "source", "text", "shape"].includes(layer.type));
}

function snapshotState() {
  return {
    mapImageDataUrl,
    legendWidthPercent: getLegendWidthPercent(),
    selectedLayerId,
    selectedLayerIds: [...selectedLayerIds],
    customIndex,
    titleInput: titleInput.value,
    subtitleInput: subtitleInput.value,
    sourceInput: sourceInput.value,
    northLabelInput: northLabelInput.value,
    gridHidden: boardGrid.classList.contains("is-hidden"),
    layers: layers.map((layer) => {
      const element = getElementByLayerId(layer.id);
      return {
        ...layer,
        html: element ? element.innerHTML : "",
        styles: element ? {
          left: element.style.left,
          top: element.style.top,
          width: element.style.width,
          height: element.style.height,
          visibility: element.style.visibility,
        } : null,
        textStyle: element && isTextStyleLayer(layer)
          ? {
            fontSize: element.dataset.textFontSize,
            textColor: element.dataset.textColor,
            backgroundColor: element.dataset.textBackground,
            fontFamily: element.dataset.textFontFamily,
            fontWeight: element.dataset.textFontWeight,
            fontStyle: element.dataset.textFontStyle,
            bgTransparent: element.dataset.bgTransparent === "true",
          }
          : null,
        shapeStyle: element && isShapeStyleLayer(layer)
          ? {
            fillColor: element.dataset.shapeFill,
            strokeColor: element.dataset.shapeStroke,
            fillTransparent: element.dataset.shapeTransparent === "true",
            shapeType: element.dataset.shapeType || "",
          }
          : null,
        rotation: element ? (element.dataset.rotation || "") : "",
      };
    }),
  };
}

function pushHistory() {
  if (isRestoringHistory) {
    return;
  }

  historyStack.push(snapshotState());
  if (historyStack.length > 80) {
    historyStack.shift();
  }
}

function restoreState(state) {
  isRestoringHistory = true;

  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    if (!["map-base", "title-base", "source-base", "north-base"].includes(element.dataset.layerId)) {
      element.remove();
    }
  });

  layers.length = 0;
  state.layers.forEach((layerState) => {
    layers.push({
      id: layerState.id,
      type: layerState.type,
      label: layerState.label,
      locked: layerState.locked,
      visible: layerState.visible,
      listed: layerState.listed,
      src: layerState.src,
      shapeType: layerState.shapeType,
    });

    let element = getElementByLayerId(layerState.id);
    if (!element && layerState.type !== "map") {
      element = document.createElement("div");
      element.className = `board-element board-element--${layerState.type}`;
      element.dataset.layerId = layerState.id;
      element.dataset.type = layerState.type;
      if (layerState.shapeStyle?.shapeType) {
        element.dataset.shapeType = layerState.shapeStyle.shapeType;
      }
      element.innerHTML = layerState.html;
      board.appendChild(element);
      attachElementEvents(element);
    }

    if (element && layerState.styles) {
      element.style.left = layerState.styles.left || "";
      element.style.top = layerState.styles.top || "";
      element.style.width = layerState.styles.width || "";
      element.style.height = layerState.styles.height || "";
      element.style.visibility = layerState.styles.visibility || "";
      element.innerHTML = layerState.html || element.innerHTML;
      element.classList.toggle("is-hidden", !layerState.visible);
      element.classList.toggle("is-locked", !!layerState.locked);
      if (layerState.textStyle && isTextStyleLayer(layerState)) {
        initializeTextStyleState(layerState.id, layerState.textStyle);
      }
      if (layerState.shapeStyle && isShapeStyleLayer(layerState)) {
        if (layerState.shapeStyle.shapeType) {
          element.dataset.shapeType = layerState.shapeStyle.shapeType;
        }
        initializeShapeStyleState(layerState.id, layerState.shapeStyle);
      }
      if (layerState.rotation) {
        element.dataset.rotation = layerState.rotation;
      }
    }
  });

  mapImageDataUrl = state.mapImageDataUrl || "";
  mapImage.src = mapImageDataUrl;
  mapLayer.classList.toggle("has-image", Boolean(mapImageDataUrl));
  emptyMapState.hidden = Boolean(mapImageDataUrl);
  setLegendWidthPercent(state.legendWidthPercent || 26);
  resetMapLayerFrame();

  titleInput.value = state.titleInput;
  subtitleInput.value = state.subtitleInput;
  sourceInput.value = state.sourceInput;
  northLabelInput.value = state.northLabelInput;
  customIndex = state.customIndex;
  selectedLayerIds = new Set(state.selectedLayerIds && state.selectedLayerIds.length ? state.selectedLayerIds : [state.selectedLayerId || "title-base"]);

  boardGrid.classList.toggle("is-hidden", state.gridHidden);

  syncBaseTexts();
  applyLayerOrder();
  renderLayers();
  if (selectedLayerIds.size > 1) {
    board.querySelectorAll("[data-layer-id]").forEach((element) => {
      element.classList.toggle("is-selected", selectedLayerIds.has(element.dataset.layerId));
    });
    refreshSelectedControls();
  } else {
    selectLayer(state.selectedLayerId || "title-base");
  }

  isRestoringHistory = false;
}

function undoLastAction() {
  const previousState = historyStack.pop();
  if (!previousState) {
    return;
  }

  restoreState(previousState);
}

function createResizeHandle(element) {
  if (element.querySelector(".resize-handle")) {
    return;
  }

  const edgeHandles = [
    { position: "top", cursor: "ns-resize" },
    { position: "right", cursor: "ew-resize" },
    { position: "bottom", cursor: "ns-resize" },
    { position: "left", cursor: "ew-resize" },
  ];

  const cornerHandles = [
    { position: "top-left", cursor: "nwse-resize" },
    { position: "top-right", cursor: "nesw-resize" },
    { position: "bottom-right", cursor: "nwse-resize" },
    { position: "bottom-left", cursor: "nesw-resize" },
  ];

  const isMapElement = element.dataset.type === "map";
  const handlePositions = isMapElement ? edgeHandles : [...edgeHandles, ...cornerHandles];

  handlePositions.forEach(({ position, cursor }) => {
    const handle = document.createElement("button");
    handle.type = "button";
    handle.className = `resize-handle resize-handle--${position}`;
    handle.dataset.resize = position;
    handle.setAttribute("aria-label", `Redimensionner ${position}`);
    handle.style.cursor = cursor;
    element.appendChild(handle);
  });
}

function setBoardText(target, value, fallback) {
  target.textContent = value.trim() || fallback;
}

function getTitleSubtitleNode(element) {
  return null;
}

function getSourceTextNode(element) {
  return element.querySelector("#boardSource") || element.querySelectorAll("p")[1] || null;
}

function isTextStyleLayer(layer) {
  return Boolean(layer && ["title", "source", "text"].includes(layer.type));
}

function isShapeStyleLayer(layer) {
  return Boolean(layer && layer.type === "shape");
}

function getDefaultShapeStyle(shapeType = "rectangle") {
  if (shapeType === "black-bar") {
    return {
      fillColor: "#000000",
      strokeColor: "#000000",
      fillTransparent: false,
    };
  }

  if (["line", "arrow-right"].includes(shapeType)) {
    return {
      fillColor: "#000091",
      strokeColor: "#000091",
      fillTransparent: false,
    };
  }

  if (["bubble-left", "bubble-bottom", "callout-rect-left", "callout-rect-right", "notch-top", "notch-bottom"].includes(shapeType)) {
    return {
      fillColor: "#ffffff",
      strokeColor: "#000091",
      fillTransparent: false,
    };
  }

  return {
    fillColor: "rgba(0, 0, 145, 0.18)",
    strokeColor: "#000091",
    fillTransparent: false,
  };
}

function initializeShapeStyleState(layerId, overrides = {}) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element || !isShapeStyleLayer(layer)) {
    return;
  }

  const shapeType = element.dataset.shapeType || "rectangle";
  const defaults = { ...getDefaultShapeStyle(shapeType), ...overrides };
  element.dataset.shapeFill = defaults.fillColor;
  element.dataset.shapeStroke = defaults.strokeColor;
  element.dataset.shapeTransparent = defaults.fillTransparent ? "true" : "false";
  syncShapeStylePresentation(layerId);
}

function getShapeStyleState(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element || !isShapeStyleLayer(layer)) {
    return null;
  }

  const shapeType = element.dataset.shapeType || "rectangle";
  const defaults = getDefaultShapeStyle(shapeType);
  return {
    fillColor: element.dataset.shapeFill || defaults.fillColor,
    strokeColor: element.dataset.shapeStroke || defaults.strokeColor,
    fillTransparent: (element.dataset.shapeTransparent || String(defaults.fillTransparent)) === "true",
  };
}

function syncShapeStylePresentation(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  const shapeState = getShapeStyleState(layerId);
  if (!layer || !element || !shapeState || !isShapeStyleLayer(layer)) {
    return;
  }

  const fill = shapeState.fillTransparent ? "transparent" : shapeState.fillColor;
  setInlineStyle(element, "--shape-fill", fill);
  setInlineStyle(element, "--shape-stroke", shapeState.strokeColor || "#000091");
}

function applyShapeStylesToElement(layerId, nextStyle = {}) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element || !isShapeStyleLayer(layer)) {
    return;
  }

  const currentState = getShapeStyleState(layerId);
  const merged = { ...currentState, ...nextStyle };
  element.dataset.shapeFill = merged.fillColor;
  element.dataset.shapeStroke = merged.strokeColor || "#000091";
  element.dataset.shapeTransparent = merged.fillTransparent ? "true" : "false";
  syncShapeStylePresentation(layerId);
}

function toHexColor(color) {
  if (!color || color === "transparent") {
    return "#000000";
  }

  if (color.startsWith("#")) {
    return color;
  }

  const match = color.match(/\d+/g);
  if (!match || match.length < 3) {
    return "#000000";
  }

  return `#${match.slice(0, 3).map((value) => Number(value).toString(16).padStart(2, "0")).join("")}`;
}

function getTextTargetNode(layer, element) {
  if (!layer || !element) {
    return null;
  }

  if (layer.type === "title") {
    return element.querySelector("h2");
  }

  if (layer.type === "source") {
    return getSourceTextNode(element);
  }

  if (layer.type === "text") {
    return element.querySelector("p");
  }

  return null;
}

function getTitleStyleNodes(element) {
  return [
    element.querySelector("h2"),
  ].filter(Boolean);
}

function setInlineStyle(node, property, value) {
  if (!node) {
    return;
  }

  node.style.setProperty(property, value, "important");
}

function getDefaultTextStyle(layerType) {
  if (layerType === "title") {
    return {
      fontSize: 13,
      textColor: "#ffffff",
      backgroundColor: "#000091",
      fontFamily: "Marianne, Arial, sans-serif",
      fontWeight: "700",
      fontStyle: "normal",
      bgTransparent: false,
    };
  }

  if (layerType === "text") {
    return {
      fontSize: 18,
      textColor: "#161616",
      backgroundColor: "#ffffff",
      fontFamily: "Marianne, Arial, sans-serif",
      fontWeight: "400",
      fontStyle: "normal",
      bgTransparent: true,
    };
  }

  return {
    fontSize: 14,
    textColor: "#5c5c5c",
    backgroundColor: "#ffffff",
    fontFamily: "Marianne, Arial, sans-serif",
    fontWeight: "400",
    fontStyle: "normal",
    bgTransparent: true,
  };
}

function initializeTextStyleState(layerId, overrides = {}) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element || !isTextStyleLayer(layer)) {
    return;
  }

  const defaults = { ...getDefaultTextStyle(layer.type), ...overrides };
  element.dataset.textFontSize = String(defaults.fontSize);
  element.dataset.textColor = defaults.textColor;
  element.dataset.textBackground = defaults.backgroundColor;
  element.dataset.textFontFamily = defaults.fontFamily;
  element.dataset.textFontWeight = String(defaults.fontWeight);
  element.dataset.textFontStyle = defaults.fontStyle;
  element.dataset.bgTransparent = defaults.bgTransparent ? "true" : "false";
  if (!element.dataset.rotation) {
    element.dataset.rotation = layer.type === "source" ? "180" : "0";
  }
  syncElementTextStylePresentation(layerId);
}

function getTextStyleState(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  const textNode = getTextTargetNode(layer, element);

  if (!layer || !element || !textNode || !isTextStyleLayer(layer)) {
    return null;
  }

  const defaults = getDefaultTextStyle(layer.type);

  return {
    fontSize: Number(element.dataset.textFontSize || defaults.fontSize),
    textColor: element.dataset.textColor || defaults.textColor,
    backgroundColor: element.dataset.textBackground || defaults.backgroundColor,
    fontFamily: element.dataset.textFontFamily || defaults.fontFamily,
    fontWeight: element.dataset.textFontWeight || defaults.fontWeight,
    fontStyle: element.dataset.textFontStyle || defaults.fontStyle,
    bgTransparent: (element.dataset.bgTransparent || String(defaults.bgTransparent)) === "true",
  };
}

function syncElementTextStylePresentation(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  const styleState = getTextStyleState(layerId);
  if (!layer || !element || !styleState || !isTextStyleLayer(layer)) {
    return;
  }

  const backgroundColor = styleState.bgTransparent ? "transparent" : styleState.backgroundColor;
  const rotation = Number(element.dataset.rotation || (layer.type === "source" ? 180 : 0));
  setInlineStyle(element, "--carto-text-color", styleState.textColor);
  setInlineStyle(element, "--carto-text-bg", backgroundColor);
  setInlineStyle(element, "--carto-font-family", styleState.fontFamily);
  setInlineStyle(element, "--carto-font-weight", String(styleState.fontWeight));
  setInlineStyle(element, "--carto-font-style", styleState.fontStyle);
  setInlineStyle(element, "background-color", backgroundColor);

  if (layer.type === "title") {
    setInlineStyle(element, "--carto-title-font-size", `${styleState.fontSize}px`);
    const titleNode = element.querySelector("h2");

    setInlineStyle(element, "font-family", styleState.fontFamily);
    setInlineStyle(element, "color", styleState.textColor);
    setInlineStyle(element, "background", backgroundColor);
    setInlineStyle(element, "transform", `rotate(${rotation}deg)`);

    setInlineStyle(titleNode, "font-family", styleState.fontFamily);
    setInlineStyle(titleNode, "font-weight", String(styleState.fontWeight));
    setInlineStyle(titleNode, "font-style", styleState.fontStyle);
    setInlineStyle(titleNode, "font-size", `${styleState.fontSize}px`);
    setInlineStyle(titleNode, "line-height", String(clamp(1.02 + styleState.fontSize / 180, 1.02, 1.22)));
    setInlineStyle(titleNode, "color", styleState.textColor);
  }

  if (layer.type === "source") {
    setInlineStyle(element, "--carto-source-font-size", `${styleState.fontSize}px`);
    const sourceNode = getSourceTextNode(element);

    setInlineStyle(element, "font-family", styleState.fontFamily);
    setInlineStyle(element, "font-weight", String(styleState.fontWeight));
    setInlineStyle(element, "font-style", styleState.fontStyle);
    setInlineStyle(element, "color", styleState.textColor);
    setInlineStyle(element, "background", backgroundColor);

    setInlineStyle(sourceNode, "font-family", styleState.fontFamily);
    setInlineStyle(sourceNode, "font-weight", String(styleState.fontWeight));
    setInlineStyle(sourceNode, "font-style", styleState.fontStyle);
    setInlineStyle(sourceNode, "font-size", `${styleState.fontSize}px`);
    setInlineStyle(sourceNode, "line-height", String(clamp(1.05 + styleState.fontSize / 160, 1.05, 1.3)));
    setInlineStyle(sourceNode, "color", styleState.textColor);
    setInlineStyle(element, "transform", `rotate(${rotation}deg)`);
  }

  if (layer.type === "text") {
    const textNode = element.querySelector("p");
    setInlineStyle(element, "--carto-text-font-size", `${styleState.fontSize}px`);
    setInlineStyle(element, "font-family", styleState.fontFamily);
    setInlineStyle(element, "font-weight", String(styleState.fontWeight));
    setInlineStyle(element, "font-style", styleState.fontStyle);
    setInlineStyle(element, "color", styleState.textColor);
    setInlineStyle(element, "background", backgroundColor);
    setInlineStyle(element, "transform", `rotate(${rotation}deg)`);
    setInlineStyle(textNode, "font-family", styleState.fontFamily);
    setInlineStyle(textNode, "font-weight", String(styleState.fontWeight));
    setInlineStyle(textNode, "font-style", styleState.fontStyle);
    setInlineStyle(textNode, "font-size", `${styleState.fontSize}px`);
    setInlineStyle(textNode, "color", styleState.textColor);
  }
}

function updateTextStyleButtons(styleState, disabled) {
  toggleBoldButton.disabled = disabled;
  toggleItalicButton.disabled = disabled;
  toggleTransparentBackgroundButton.disabled = disabled;
  toggleBoldButton.classList.toggle("is-active", !disabled && Number(styleState?.fontWeight || 400) >= 600);
  toggleItalicButton.classList.toggle("is-active", !disabled && styleState?.fontStyle === "italic");
  toggleTransparentBackgroundButton.classList.toggle("is-active", !disabled && Boolean(styleState?.bgTransparent));
}

function applyTextStylesToElement(layerId, nextStyle = {}) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  const textNode = getTextTargetNode(layer, element);

  if (!layer || !element || !textNode || !isTextStyleLayer(layer)) {
    return;
  }

  const currentStyle = getTextStyleState(layerId);
  const mergedStyle = { ...currentStyle, ...nextStyle };
  element.dataset.textFontSize = String(mergedStyle.fontSize);
  element.dataset.textColor = mergedStyle.textColor;
  element.dataset.textBackground = mergedStyle.backgroundColor;
  element.dataset.textFontFamily = mergedStyle.fontFamily;
  element.dataset.textFontWeight = String(mergedStyle.fontWeight);
  element.dataset.textFontStyle = mergedStyle.fontStyle;
  element.dataset.bgTransparent = mergedStyle.bgTransparent ? "true" : "false";
  syncElementTextStylePresentation(layerId);
  refreshSelectedControls();
}

function applyTextLayout(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element) {
    return;
  }

  const width = parseFloat(element.style.width || "12");
  const height = parseFloat(element.style.height || "12");

  if (layer.type === "title") {
    const titleNode = element.querySelector("h2");
    if (titleNode) {
      if (!titleNode.style.fontSize) {
        titleNode.style.fontSize = `${clamp(width * 0.09 + height * 0.04, 0.95, 2.1)}rem`;
      }
      if (!titleNode.style.lineHeight) {
        titleNode.style.lineHeight = String(clamp(1.05 + height * 0.01, 1.05, 1.25));
      }
    }
  }

  if (layer.type === "source") {
    const sourceNode = getSourceTextNode(element);
    if (sourceNode) {
      if (!sourceNode.style.fontSize) {
        sourceNode.style.fontSize = `${clamp(width * 0.028 + height * 0.018, 0.68, 1)}rem`;
      }
      if (!sourceNode.style.lineHeight) {
        sourceNode.style.lineHeight = String(clamp(1.2 + height * 0.012, 1.2, 1.5));
      }
    }
  }

}

function getLayerLabel(layer) {
  if (layer.type === "map") {
    return mapImageDataUrl ? "Carte importee" : "Fond cartographique";
  }
  return layer.label;
}

function setLayerLockedVisual(layerId) {
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element) {
    return;
  }
  element.classList.toggle("is-locked", layer.locked);
}

function applyLayerOrder() {
  layers.forEach((layer, index) => {
    const element = getElementByLayerId(layer.id);
    if (element) {
      element.style.zIndex = String(index + 1);
      setLayerLockedVisual(layer.id);
      applyTextLayout(layer.id);
      if (isShapeStyleLayer(layer)) {
        syncShapeStylePresentation(layer.id);
      }
    }
  });
}

function getSelectedLayers() {
  return [...selectedLayerIds]
    .map((layerId) => getLayer(layerId))
    .filter(Boolean);
}

function getUnlockedSelectedLayers() {
  return getSelectedLayers().filter((layer) => !layer.locked);
}

function getTextStyleTargetLayers() {
  return getUnlockedSelectedLayers().filter((layer) => isTextStyleLayer(layer));
}

function getRotationTargetLayers() {
  return getUnlockedSelectedLayers().filter((layer) => isRotatableLayer(layer));
}

function getResizeTargetLayers() {
  return getUnlockedSelectedLayers();
}

function clearSelection() {
  selectedLayerId = "";
  selectedLayerIds = new Set();
  legendSelection = null;
  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    element.classList.remove("is-selected");
  });
  renderLayers();
  refreshSelectedControls();
}

function selectAllLayers() {
  const visibleLayerIds = layers
    .filter((layer) => layer.visible)
    .map((layer) => layer.id);

  if (!visibleLayerIds.length) {
    return;
  }

  selectedLayerIds = new Set(visibleLayerIds);
  selectedLayerId = visibleLayerIds[visibleLayerIds.length - 1];
  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    element.classList.toggle("is-selected", selectedLayerIds.has(element.dataset.layerId));
  });
  renderLayers();
  refreshSelectedControls();
}

function updateZoom() {
  const supportsCssZoom = typeof document !== "undefined" && "zoom" in document.documentElement.style;
  if (supportsCssZoom) {
    board.style.zoom = String(zoomLevel);
    board.style.transform = "none";
  } else {
    board.style.zoom = "";
    board.style.transform = `scale(${zoomLevel})`;
  }
  zoomValue.textContent = `${Math.round((zoomLevel / ZOOM_BASELINE) * 100)} %`;
}

function selectLayer(layerId, options = {}) {
  const { additive = false } = options;
  legendSelection = null;
  selectedLayerId = layerId;
  if (additive) {
    if (selectedLayerIds.has(layerId)) {
      selectedLayerIds.delete(layerId);
      if (!selectedLayerIds.size) {
        selectedLayerId = "";
      } else if (selectedLayerId === layerId) {
        selectedLayerId = [...selectedLayerIds][selectedLayerIds.size - 1] || "";
      }
    } else {
      selectedLayerIds.add(layerId);
      selectedLayerId = layerId;
    }
  } else {
    selectedLayerIds = new Set([layerId]);
  }
  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    element.classList.toggle("is-selected", selectedLayerIds.has(element.dataset.layerId));
  });
  renderLayers();
  refreshSelectedControls();
}

function refreshLayerCount() {
  layerCount.textContent = `${layers.filter((layer) => isLayerListed(layer)).length} calques`;
}

function getLegendSymbolClass(layerType) {
  if (layerType === "picto") {
    return "carto-legend-symbol carto-legend-symbol--point";
  }

  if (layerType === "north") {
    return "carto-legend-symbol carto-legend-symbol--area";
  }

  return "carto-legend-symbol carto-legend-symbol--line";
}

function getDefaultLegendTextStyle() {
  return {
    fontSize: 13,
    textColor: "#161616",
    backgroundColor: "#ffffff",
    fontFamily: "Marianne, Arial, sans-serif",
    fontWeight: "600",
    fontStyle: "normal",
    bgTransparent: false,
  };
}

function getDefaultLegendTitleStyle() {
  return {
    fontSize: 13,
    textColor: "#ffffff",
    backgroundColor: "#000091",
    fontFamily: "Marianne, Arial, sans-serif",
    fontWeight: "700",
    fontStyle: "normal",
    bgTransparent: false,
  };
}

function createLegendSection(title, type = "custom") {
  legendSectionCounter += 1;
  return {
    id: `legend-section-${legendSectionCounter}`,
    type,
    title,
    items: [],
    titleStyle: getDefaultLegendTitleStyle(),
    itemStyle: getDefaultLegendTextStyle(),
    bodyBackground: "#ffffff",
    columns: 1,
    symbolSize: 20,
  };
}

function styleLegendSymbolElement(symbol, layerType, size) {
  const safeSize = clamp(Number(size) || 20, 10, 44);

  if (layerType === "picto") {
    symbol.style.width = `${safeSize}px`;
    symbol.style.height = `${safeSize}px`;
    return;
  }

  if (layerType === "north") {
    const half = Math.round(safeSize * 0.42);
    symbol.style.width = "0";
    symbol.style.height = "0";
    symbol.style.borderLeftWidth = `${half}px`;
    symbol.style.borderRightWidth = `${half}px`;
    symbol.style.borderBottomWidth = `${safeSize}px`;
    return;
  }

  symbol.style.width = `${safeSize}px`;
  symbol.style.height = `${Math.max(2, Math.round(safeSize * 0.14))}px`;
}

function getLegendEntries() {
  const visibleLegendLayers = layers.filter((layer) => layer.visible && layer.type === "picto");
  const uniqueLegendLayers = [];
  const seen = new Set();

  visibleLegendLayers.forEach((layer) => {
    const key = layer.type === "picto"
      ? `picto:${layer.label}:${layer.src || ""}`
      : `${layer.type}:${getLayerLabel(layer)}`;

    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    uniqueLegendLayers.push({
      key,
      layerType: layer.type,
      label: getLayerLabel(layer),
      src: layer.src || "",
    });
  });

  return uniqueLegendLayers;
}

function ensureLegendBaseSections() {
  if (!legendSectionsState.length) {
    legendSectionsState = [
      createLegendSection("groupe", "picto"),
    ];
  }
}

function getSectionByType(type) {
  return legendSectionsState.find((section) => section.type === type);
}

function syncLegendSectionsWithVisibleEntries() {
  legendSectionsState = legendSectionsState.filter((section) => section.type !== "north");
  if (legendSelection && legendSelection.sectionId) {
    const selectedSectionStillExists = legendSectionsState.some((section) => section.id === legendSelection.sectionId);
    if (!selectedSectionStillExists) {
      legendSelection = null;
    }
  }

  ensureLegendBaseSections();

  const entries = getLegendEntries();
  const entryMap = new Map(entries.map((entry) => [entry.key, entry]));
  const alreadyAssigned = new Set();

  legendSectionsState.forEach((section) => {
    section.items = section.items.filter((item) => entryMap.has(item.key));
    section.items.forEach((item) => alreadyAssigned.add(item.key));
  });

  entries.forEach((entry) => {
    if (alreadyAssigned.has(entry.key)) {
      return;
    }

    const targetSection = getSectionByType(entry.layerType) || legendSectionsState[0];
    if (!targetSection) {
      return;
    }
    targetSection.items.push(entry);
  });
}

function setLegendSelection(selection) {
  legendSelection = selection;
  selectedLayerId = "";
  selectedLayerIds = new Set();
  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    element.classList.remove("is-selected");
  });
  renderLayers();
  refreshSelectedControls();
  renderLegend();
}

function getLegendSelectionStyleState() {
  if (!legendSelection) {
    return null;
  }

  const section = legendSectionsState.find((item) => item.id === legendSelection.sectionId);
  if (!section) {
    return null;
  }

  if (legendSelection.kind === "section-title") {
    return { ...section.titleStyle };
  }

  if (legendSelection.kind === "section-body") {
    return {
      ...section.itemStyle,
      backgroundColor: section.bodyBackground,
    };
  }

  if (legendSelection.kind === "item-label") {
    return { ...section.itemStyle };
  }

  return null;
}

function updateLegendSelectionText(value) {
  if (!legendSelection) {
    return;
  }

  const section = legendSectionsState.find((item) => item.id === legendSelection.sectionId);
  if (!section) {
    return;
  }

  if (legendSelection.kind === "section-title") {
    section.title = value || "Sans titre";
  }

  if (legendSelection.kind === "item-label" && legendSelection.itemKey) {
    const item = section.items.find((entry) => entry.key === legendSelection.itemKey);
    if (item) {
      item.label = value || item.label;
    }
  }

  renderLegend();
  refreshSelectedControls();
}

function updateLegendSelectionStyle(nextStyle) {
  if (!legendSelection) {
    return;
  }

  const section = legendSectionsState.find((item) => item.id === legendSelection.sectionId);
  if (!section) {
    return;
  }

  if (legendSelection.kind === "section-title") {
    section.titleStyle = { ...section.titleStyle, ...nextStyle };
  }

  if (legendSelection.kind === "item-label") {
    section.itemStyle = { ...section.itemStyle, ...nextStyle };
  }

  if (legendSelection.kind === "section-body") {
    if (nextStyle.backgroundColor) {
      section.bodyBackground = nextStyle.backgroundColor;
    }
    section.itemStyle = { ...section.itemStyle, ...nextStyle };
  }

  renderLegend();
  refreshSelectedControls();
}

function moveLegendItem(itemKey, targetSectionId, targetIndex = -1) {
  let sourceSection = null;
  let sourceIndex = -1;
  legendSectionsState.forEach((section) => {
    const index = section.items.findIndex((item) => item.key === itemKey);
    if (index !== -1) {
      sourceSection = section;
      sourceIndex = index;
    }
  });

  const targetSection = legendSectionsState.find((section) => section.id === targetSectionId);
  if (!sourceSection || !targetSection || sourceIndex === -1) {
    return;
  }

  const [item] = sourceSection.items.splice(sourceIndex, 1);
  if (targetIndex < 0 || targetIndex > targetSection.items.length) {
    targetSection.items.push(item);
  } else {
    targetSection.items.splice(targetIndex, 0, item);
  }
}

function getLegendGroups() {
  syncLegendSectionsWithVisibleEntries();
  return legendSectionsState;
}

function renderLegend() {
  const groups = getLegendGroups();
  legendSections.innerHTML = "";
  legendHint.textContent = "";
  legendHint.hidden = true;

  groups.forEach((group) => {
    const section = document.createElement("section");
    section.className = "carto-legend-section";
    section.dataset.sectionId = group.id;

    const titleRow = document.createElement("div");
    titleRow.className = "carto-legend-section__title-row";

    const title = document.createElement("div");
    title.className = "carto-legend-section__title";
    title.textContent = group.title;
    title.style.fontSize = `${group.titleStyle.fontSize}px`;
    title.style.color = group.titleStyle.textColor;
    title.style.backgroundColor = group.titleStyle.bgTransparent ? "transparent" : group.titleStyle.backgroundColor;
    title.style.fontFamily = group.titleStyle.fontFamily;
    title.style.fontWeight = group.titleStyle.fontWeight;
    title.style.fontStyle = group.titleStyle.fontStyle;
    title.addEventListener("click", () => setLegendSelection({ kind: "section-title", sectionId: group.id }));
    if (legendSelection?.kind === "section-title" && legendSelection.sectionId === group.id) {
      title.classList.add("is-selected");
    }

    titleRow.append(title);

    const body = document.createElement("div");
    body.className = "carto-legend-section__body";
    body.style.backgroundColor = group.bodyBackground;
    body.addEventListener("click", (event) => {
      if (event.target === body) {
        setLegendSelection({ kind: "section-body", sectionId: group.id });
      }
    });
    if (legendSelection?.kind === "section-body" && legendSelection.sectionId === group.id) {
      body.classList.add("is-selected");
    }
    body.addEventListener("dragover", (event) => {
      event.preventDefault();
      section.classList.add("is-drop-target");
    });
    body.addEventListener("dragleave", () => {
      section.classList.remove("is-drop-target");
    });
    body.addEventListener("drop", (event) => {
      event.preventDefault();
      section.classList.remove("is-drop-target");
      if (!draggedLegendItem) {
        return;
      }
      moveLegendItem(draggedLegendItem.itemKey, group.id);
      draggedLegendItem = null;
      renderLegend();
    });

    const list = document.createElement("ul");
    list.className = "carto-legend-list";
    const columns = clamp(Number(group.columns) || 1, 1, 6);
    const symbolSize = clamp(Number(group.symbolSize) || 20, 10, 44);
    list.style.gridTemplateColumns = `repeat(${columns}, minmax(0, 1fr))`;

    group.items.forEach((itemEntry) => {
      const item = document.createElement("li");
      item.draggable = true;
      item.dataset.itemKey = itemEntry.key;
      item.addEventListener("dragstart", () => {
        draggedLegendItem = { itemKey: itemEntry.key };
      });
      item.addEventListener("dragend", () => {
        draggedLegendItem = null;
        section.classList.remove("is-drop-target");
      });

      let symbol;
      if (itemEntry.layerType === "picto" && itemEntry.src) {
        symbol = document.createElement("img");
        symbol.className = "carto-legend-picto";
        symbol.src = itemEntry.src;
        symbol.alt = itemEntry.label;
      } else {
        symbol = document.createElement("span");
        symbol.className = getLegendSymbolClass(itemEntry.layerType);
      }
      styleLegendSymbolElement(symbol, itemEntry.layerType, symbolSize);
      item.style.gridTemplateColumns = `${symbolSize}px minmax(0, 1fr)`;

      const label = document.createElement("span");
      label.className = "carto-legend-label";
      label.textContent = itemEntry.label;
      label.style.fontSize = `${group.itemStyle.fontSize}px`;
      label.style.color = group.itemStyle.textColor;
      label.style.fontFamily = group.itemStyle.fontFamily;
      label.style.fontWeight = group.itemStyle.fontWeight;
      label.style.fontStyle = group.itemStyle.fontStyle;
      item.addEventListener("click", (event) => {
        event.stopPropagation();
        setLegendSelection({ kind: "item-label", sectionId: group.id, itemKey: itemEntry.key });
      });
      if (legendSelection?.kind === "item-label" && legendSelection.sectionId === group.id && legendSelection.itemKey === itemEntry.key) {
        item.classList.add("is-selected");
      }

      item.append(symbol, label);
      list.appendChild(item);
    });

    body.appendChild(list);
    section.append(titleRow, body);
    legendSections.appendChild(section);
  });
}

function renderLayers() {
  layersList.innerHTML = "";

  const groupedLayers = {
    block1: [],
    block2: [],
    block3: [],
    block4: [],
  };

  [...layers].reverse().filter((layer) => isLayerListed(layer)).forEach((layer) => {
    if (["title", "north", "source", "map"].includes(layer.type)) {
      groupedLayers.block1.push(layer);
      return;
    }

    if (layer.type === "text") {
      groupedLayers.block2.push(layer);
      return;
    }

    if (layer.type === "picto") {
      groupedLayers.block3.push(layer);
      return;
    }

    if (layer.type === "shape") {
      groupedLayers.block4.push(layer);
    }
  });

  const groups = [
    { key: "block1", title: "Bloc 1 : Titre, orientation, source", hint: "Habillage principal" },
    { key: "block2", title: "Bloc 2 : Textes libres", hint: "Tous les textes libres" },
    { key: "block3", title: "Bloc 3 : Pictogrammes", hint: "Tous les pictogrammes poses" },
    { key: "block4", title: "Bloc 4 : Forme", hint: "Toutes les formes posees" },
  ];

  groups.forEach((group) => {
    const section = document.createElement("section");
    section.className = "layer-group";
    if (collapsedLayerGroups.has(group.key)) {
      section.classList.add("is-collapsed");
    }

    const head = document.createElement("div");
    head.className = "layer-group__head";

    const titleWrap = document.createElement("div");
    titleWrap.className = "layer-group__title-wrap";

    const title = document.createElement("h3");
    title.className = "layer-group__title";
    title.textContent = group.title;

    const hint = document.createElement("p");
    hint.className = "layer-group__hint";
    hint.textContent = group.hint;
    titleWrap.append(title, hint);

    const toggleButton = document.createElement("button");
    toggleButton.type = "button";
    toggleButton.className = "layer-group__toggle";
    toggleButton.textContent = collapsedLayerGroups.has(group.key) ? "+" : "-";
    toggleButton.setAttribute("aria-expanded", String(!collapsedLayerGroups.has(group.key)));
    toggleButton.title = collapsedLayerGroups.has(group.key) ? "Deplier le bloc" : "Plier le bloc";
    toggleButton.addEventListener("click", (event) => {
      event.stopPropagation();
      if (collapsedLayerGroups.has(group.key)) {
        collapsedLayerGroups.delete(group.key);
      } else {
        collapsedLayerGroups.add(group.key);
      }
      renderLayers();
    });

    head.append(titleWrap, toggleButton);
    section.appendChild(head);

    const list = document.createElement("div");
    list.className = "layer-group__list";

    if (!groupedLayers[group.key].length) {
      const empty = document.createElement("p");
      empty.className = "layer-group__empty";
      empty.textContent = "Aucun element dans ce bloc.";
      list.appendChild(empty);
    }

    groupedLayers[group.key].forEach((layer) => {
    const row = document.createElement("div");
    row.className = "layer-row";
    if (selectedLayerIds.has(layer.id)) {
      row.classList.add("is-active");
    }

    const swatch = document.createElement("span");
    swatch.className = "layer-row__swatch";

    const meta = document.createElement("div");
    meta.className = "layer-row__meta";

    const title = document.createElement("span");
    title.className = "layer-row__title";
    title.textContent = getLayerLabel(layer);

    const type = document.createElement("span");
    type.className = "layer-row__type";
    type.textContent = layer.type;

    meta.append(title, type);

    const actions = document.createElement("div");
    actions.className = "layer-row__actions";

    const visibilityButton = document.createElement("button");
    visibilityButton.type = "button";
    visibilityButton.className = "layer-icon-btn";
    visibilityButton.innerHTML = `<span class="${layer.visible ? "fr-icon-eye-line" : "fr-icon-eye-off-line"}" aria-hidden="true"></span>`;
    visibilityButton.setAttribute("aria-label", layer.visible ? "Calque visible" : "Calque masque");
    if (!layer.visible) {
      visibilityButton.classList.add("is-hidden-toggle");
    }
    visibilityButton.title = layer.visible ? "Masquer" : "Afficher";
    visibilityButton.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleLayerVisibility(layer.id);
    });

    const upButton = document.createElement("button");
    upButton.type = "button";
    upButton.className = "layer-icon-btn";
    upButton.textContent = "+";
    upButton.title = "Monter";
    upButton.addEventListener("click", (event) => {
      event.stopPropagation();
      moveLayer(layer.id, 1);
    });

    const downButton = document.createElement("button");
    downButton.type = "button";
    downButton.className = "layer-icon-btn";
    downButton.textContent = "-";
    downButton.title = "Descendre";
    downButton.addEventListener("click", (event) => {
      event.stopPropagation();
      moveLayer(layer.id, -1);
    });

    const lockButton = document.createElement("button");
    lockButton.type = "button";
    lockButton.className = "layer-icon-btn";
    lockButton.innerHTML = `<span class="${layer.locked ? "fr-icon-lock-line" : "fr-icon-lock-unlock-line"}" aria-hidden="true"></span>`;
    lockButton.setAttribute("aria-label", layer.locked ? "Calque verrouille" : "Calque deverrouille");
    if (layer.locked) {
      lockButton.classList.add("is-locked-toggle");
    }
    lockButton.title = layer.locked ? "Deverrouiller" : "Verrouiller";
    lockButton.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleLayerLock(layer.id);
    });

    actions.append(visibilityButton, upButton, downButton, lockButton);
    row.append(swatch, meta, actions);
    row.addEventListener("click", () => selectLayer(layer.id));
      list.appendChild(row);
    });

    section.appendChild(list);
    layersList.appendChild(section);
  });

  refreshLayerCount();
  renderLegend();
}

function buildCopiedLayerPayload(layerId) {
  const layer = getLayer(layerId);
  const element = layer ? getElementByLayerId(layer.id) : null;
  if (!layer || !element || layer.type === "map") {
    return null;
  }

  return {
    type: layer.type,
    label: layer.label,
    src: layer.src || "",
    shapeType: element.dataset.shapeType || "",
    shapeStyle: isShapeStyleLayer(layer) ? getShapeStyleState(layer.id) : null,
    html: element.innerHTML,
    left: parseFloat(element.style.left || "0"),
    top: parseFloat(element.style.top || "0"),
    width: parseFloat(element.style.width || "12"),
    height: parseFloat(element.style.height || "12"),
  };
}

function pasteCopiedLayer() {
  if (!copiedLayerPayload) {
    return;
  }

  pushHistory();
  const payload = copiedLayerPayload;
  const layerId = getNextLayerId(payload.type);
  const element = document.createElement("div");
  element.className = `board-element board-element--${payload.type}`;
  element.dataset.layerId = layerId;
  element.dataset.type = payload.type;
  if (payload.shapeType) {
    element.dataset.shapeType = payload.shapeType;
  }
  element.style.left = `${clamp(payload.left + 2, 0, 90)}%`;
  element.style.top = `${clamp(payload.top + 2, 0, 90)}%`;
  element.style.width = `${payload.width}%`;
  element.style.height = `${payload.height}%`;
  element.innerHTML = payload.html;

  board.appendChild(element);
  attachElementEvents(element);
  applyTextLayout(layerId);

  layers.push({
    id: layerId,
    type: payload.type,
    label: payload.label,
    src: payload.src || undefined,
    shapeType: payload.shapeType || undefined,
    locked: false,
    visible: true,
    listed: true,
  });

  if (payload.shapeStyle && payload.type === "shape") {
    initializeShapeStyleState(layerId, payload.shapeStyle);
  }

  applyLayerOrder();
  renderLayers();
  selectLayer(layerId);
}

function toggleLayerVisibility(layerId) {
  pushHistory();
  const layer = getLayer(layerId);
  const element = getElementByLayerId(layerId);
  if (!layer || !element) {
    return;
  }

  layer.visible = !layer.visible;
  layer.listed = true;
  element.classList.toggle("is-hidden", !layer.visible);
  if (layer.type === "map") {
    element.style.visibility = layer.visible ? "visible" : "hidden";
  }
  renderLayers();
}

function toggleLayerLock(layerId) {
  pushHistory();
  const layer = getLayer(layerId);
  if (!layer) {
    return;
  }

  layer.locked = !layer.locked;
  setLayerLockedVisual(layerId);
  renderLayers();
  refreshSelectedControls();
}

function moveLayer(layerId, direction) {
  pushHistory();
  const index = getLayerIndex(layerId);
  if (index === -1) {
    return;
  }

  const nextIndex = index + direction;
  if (nextIndex < 0 || nextIndex >= layers.length) {
    return;
  }

  [layers[index], layers[nextIndex]] = [layers[nextIndex], layers[index]];
  applyLayerOrder();
  renderLayers();
}

function syncBaseTexts() {
  setBoardText(boardTitle, titleInput.value, "Titre de la carte");
  setBoardText(boardSource, sourceInput.value, "Source : a completer");
  setBoardText(northLabel, northLabelInput.value, "N");
  if (!boardTitle.closest("[data-layer-id]")?.dataset.textFontSize) {
    initializeTextStyleState("title-base");
  }
  if (!boardSource.closest("[data-layer-id]")?.dataset.textFontSize) {
    initializeTextStyleState("source-base");
  }
  refreshSelectedControls();
}

function refreshLegendManagerControls(styleState = null, enabled = false, textValue = "", hint = "Selectionnez un titre de sous-bloc ou un libelle de picto dans la legende.") {
  if (!legendSelectedText) {
    return;
  }

  legendSelectionHint.textContent = hint;
  legendSelectedText.value = textValue;
  legendSelectedFontSize.value = String(styleState?.fontSize || 13);
  legendSelectedTextColor.value = styleState?.textColor || "#161616";
  legendSelectedTextBackground.value = styleState?.backgroundColor || "#ffffff";
  legendSelectedFontFamily.value = styleState?.fontFamily || "Marianne, Arial, sans-serif";

  legendSelectedText.disabled = !enabled;
  legendSelectedFontSize.disabled = !enabled;
  legendSelectedTextColor.disabled = !enabled;
  legendSelectedTextBackground.disabled = !enabled;
  legendSelectedFontFamily.disabled = !enabled;
  legendToggleBoldButton.disabled = !enabled;
  legendToggleItalicButton.disabled = !enabled;
  legendToggleTransparentBackgroundButton.disabled = !enabled;

  legendToggleBoldButton.classList.toggle("is-active", enabled && Number(styleState?.fontWeight || 400) >= 600);
  legendToggleItalicButton.classList.toggle("is-active", enabled && styleState?.fontStyle === "italic");
  legendToggleTransparentBackgroundButton.classList.toggle("is-active", enabled && Boolean(styleState?.bgTransparent));

  const selectedSectionId = legendSelection?.sectionId || "";
  const selectedSectionIndex = selectedSectionId ? legendSectionsState.findIndex((section) => section.id === selectedSectionId) : -1;
  const hasSectionSelection = selectedSectionIndex !== -1;
  legendMoveUpButton.disabled = !hasSectionSelection || selectedSectionIndex <= 0;
  legendMoveDownButton.disabled = !hasSectionSelection || selectedSectionIndex >= legendSectionsState.length - 1;
  legendRenameSectionButton.disabled = !hasSectionSelection;
}

function refreshSelectedControls() {
  if (legendSelection) {
    const styleState = getLegendSelectionStyleState();
    const isSectionBody = legendSelection.kind === "section-body";
    const isEditableText = legendSelection.kind !== "section-body";
    const section = legendSectionsState.find((item) => item.id === legendSelection.sectionId);

    if (!section || !styleState) {
      legendSelection = null;
      refreshSelectedControls();
      return;
    }

    if (legendSelection.kind === "section-title") {
      selectionHint.textContent = `Legende : sous-bloc "${section.title}" selectionne.`;
      selectedText.value = section.title;
      refreshLegendManagerControls(styleState, true, section.title, `Legende : sous-bloc "${section.title}" selectionne.`);
    } else if (legendSelection.kind === "item-label") {
      const selectedItem = section.items.find((item) => item.key === legendSelection.itemKey);
      selectionHint.textContent = `Legende : element "${selectedItem?.label || ""}" selectionne.`;
      selectedText.value = selectedItem?.label || "";
      refreshLegendManagerControls(styleState, true, selectedItem?.label || "", `Legende : element "${selectedItem?.label || ""}" selectionne.`);
    } else {
      selectionHint.textContent = `Legende : fond du sous-bloc "${section.title}" selectionne.`;
      selectedText.value = "";
      refreshLegendManagerControls(styleState, false, "", `Legende : fond du sous-bloc "${section.title}" selectionne. Le texte n'est pas editable pour ce mode.`);
    }

    selectedFontSize.value = String(styleState.fontSize || 13);
    selectedTextColor.value = styleState.textColor || "#161616";
    selectedTextBackground.value = styleState.backgroundColor || "#ffffff";
    selectedFontFamily.value = styleState.fontFamily || "Marianne, Arial, sans-serif";
    legendColumns.value = String(clamp(Number(section.columns) || 1, 1, 6));
    legendSymbolSize.value = String(clamp(Number(section.symbolSize) || 20, 10, 44));
    selectedRotation.value = "0";
    selectedRotation.disabled = true;
    selectedWidth.disabled = true;
    selectedHeight.disabled = true;
    bringForwardButton.disabled = true;
    sendBackwardButton.disabled = true;
    deleteLayerButton.disabled = true;
    selectedText.disabled = !isEditableText;
    selectedFontSize.disabled = false;
    selectedTextColor.disabled = false;
    selectedTextBackground.disabled = false;
    selectedFontFamily.disabled = false;
    legendColumns.disabled = false;
    legendSymbolSize.disabled = false;
    updateTextStyleButtons(styleState, false);
    return;
  }

  const layer = getLayer(selectedLayerId);
  const element = layer ? getElementByLayerId(layer.id) : null;
  const selectedLayers = getSelectedLayers();
  const isMultiSelection = selectedLayers.length > 1;
  const resizeTargetLayers = getResizeTargetLayers();
  const textStyleTargetLayers = getTextStyleTargetLayers();
  const rotationTargetLayers = getRotationTargetLayers();
  const isEditable = Boolean(layer && element);
  const isLocked = Boolean(layer && layer.locked);
  const isDeletable = Boolean(layer && !["map", "title", "source", "north"].includes(layer.type));
  const isTextStyleEditable = isMultiSelection
    ? textStyleTargetLayers.length > 0
    : isEditable && !isLocked && isTextStyleLayer(layer);
  const isShapeStyleEditable = !isMultiSelection && isEditable && !isLocked && isShapeStyleLayer(layer);
  const isRotationEditable = isMultiSelection
    ? rotationTargetLayers.length > 0
    : isEditable && !isLocked && isRotatableLayer(layer);

  if (!layer || !element) {
    selectionHint.textContent = "Selectionnez un element sur la planche pour ajuster son texte ou sa taille.";
    selectedText.value = "";
    selectedFontSize.value = "28";
    selectedTextColor.value = "#000000";
    selectedTextBackground.value = "#ffffff";
    selectedFontFamily.value = "Marianne, Arial, sans-serif";
    selectedFontSize.disabled = true;
    selectedTextColor.disabled = true;
    selectedTextBackground.disabled = true;
    selectedFontFamily.disabled = true;
    selectedRotation.value = "0";
    selectedRotation.disabled = true;
    legendColumns.value = "1";
    legendSymbolSize.value = "20";
    legendColumns.disabled = true;
    legendSymbolSize.disabled = true;
    selectedWidth.disabled = true;
    selectedHeight.disabled = true;
    bringForwardButton.disabled = true;
    sendBackwardButton.disabled = true;
    deleteLayerButton.disabled = true;
    updateTextStyleButtons(null, true);
    refreshLegendManagerControls(null, false, "");
    return;
  }

  if (isMultiSelection) {
    selectionHint.textContent = `${selectedLayers.length} elements selectionnes sur la planche.`;
  } else {
    selectionHint.textContent = isLocked
    ? `Element selectionne : ${getLayerLabel(layer)} (verrouille).`
    : `Element selectionne : ${getLayerLabel(layer)}.`;
  }

  if (isMultiSelection && resizeTargetLayers.length) {
    const firstResizeElement = getElementByLayerId(resizeTargetLayers[0].id);
    selectedWidth.value = Math.round(parseFloat(firstResizeElement?.style.width || "12"));
    selectedHeight.value = Math.round(parseFloat(firstResizeElement?.style.height || "12"));
  } else {
    selectedWidth.value = Math.round(parseFloat(element.style.width || "12"));
    selectedHeight.value = Math.round(parseFloat(element.style.height || "12"));
  }

  if (isMultiSelection) {
    selectedText.value = "";
  } else if (layer.id === "title-base") {
    selectedText.value = titleInput.value;
  } else if (layer.id === "source-base") {
    selectedText.value = sourceInput.value;
  } else if (layer.type === "north") {
    selectedText.value = northLabelInput.value;
  } else if (layer.type === "map") {
    selectedText.value = mapImageDataUrl ? "Carte importee" : "Fond cartographique";
  } else if (layer.type === "title") {
    selectedText.value = element.querySelector("h2")?.textContent || layer.label || "";
  } else if (layer.type === "source") {
    selectedText.value = getSourceTextNode(element)?.textContent || layer.label || "";
  } else {
    selectedText.value = layer.label || "";
  }

  const styleState = isMultiSelection && textStyleTargetLayers.length
    ? getTextStyleState(textStyleTargetLayers[0].id)
    : getTextStyleState(layer.id);
  const shapeStyleState = !isMultiSelection ? getShapeStyleState(layer.id) : null;
  selectedFontSize.value = String(styleState?.fontSize || 28);
  selectedTextColor.value = styleState?.textColor || "#000000";
  selectedTextBackground.value = shapeStyleState
    ? toHexColor(shapeStyleState.fillColor)
    : styleState?.backgroundColor || "#ffffff";
  selectedFontFamily.value = styleState?.fontFamily || "Marianne, Arial, sans-serif";
  if (isMultiSelection && rotationTargetLayers.length) {
    const firstRotationElement = getElementByLayerId(rotationTargetLayers[0].id);
    selectedRotation.value = firstRotationElement?.dataset.rotation || "0";
  } else {
    selectedRotation.value = element.dataset.rotation || (layer.type === "source" ? "180" : "0");
  }

  selectedText.disabled = isMultiSelection || !isEditable || isLocked || layer.type === "map";
  selectedFontSize.disabled = !isTextStyleEditable;
  selectedTextColor.disabled = !isTextStyleEditable;
  selectedTextBackground.disabled = !(isTextStyleEditable || isShapeStyleEditable);
  selectedFontFamily.disabled = !isTextStyleEditable;
  selectedRotation.disabled = !isRotationEditable;
  legendColumns.value = "1";
  legendSymbolSize.value = "20";
  legendColumns.disabled = true;
  legendSymbolSize.disabled = true;
  if (isShapeStyleEditable) {
    updateTextStyleButtons({ bgTransparent: shapeStyleState?.fillTransparent }, false);
    toggleBoldButton.disabled = true;
    toggleItalicButton.disabled = true;
  } else {
    updateTextStyleButtons(styleState, !isTextStyleEditable);
  }
  selectedWidth.disabled = !resizeTargetLayers.length;
  selectedHeight.disabled = !resizeTargetLayers.length;
  bringForwardButton.disabled = isMultiSelection || !isEditable || isLocked;
  sendBackwardButton.disabled = isMultiSelection || !isEditable || isLocked;
  deleteLayerButton.disabled = isMultiSelection ? !selectedLayers.some((selected) => !["map", "title", "source", "north"].includes(selected.type)) : !isDeletable;
  refreshLegendManagerControls(null, false, "");
}

function updateLegendLayoutStyle(nextLayout) {
  if (!legendSelection) {
    return;
  }

  const section = legendSectionsState.find((item) => item.id === legendSelection.sectionId);
  if (!section) {
    return;
  }

  if (typeof nextLayout.columns !== "undefined") {
    section.columns = clamp(Number(nextLayout.columns) || 1, 1, 6);
  }

  if (typeof nextLayout.symbolSize !== "undefined") {
    section.symbolSize = clamp(Number(nextLayout.symbolSize) || 20, 10, 44);
  }

  renderLegend();
  refreshSelectedControls();
}

function getPictoMarkup(src, label) {
  return `<img src="${src}" alt="${label}" crossorigin="anonymous">`;
}

function buildPictoButton(config, pictoKey) {
  const button = document.createElement("button");
  button.className = "carto-picto";
  button.type = "button";
  button.dataset.picto = pictoKey;
  button.innerHTML = `
    <span class="carto-picto__icon">${getPictoMarkup(config.src, config.label)}</span>
    <span class="carto-picto__label">${config.label}</span>
  `;
  button.addEventListener("click", () => {
    createPictoElement(pictoKey);
  });
  return button;
}

function renderFavoritePictos() {
  favoritePictos.innerHTML = "";
  favoritePictoKeys.forEach((pictoKey) => {
    const config = pictoCatalog[pictoKey];
    if (config) {
      favoritePictos.appendChild(buildPictoButton(config, pictoKey));
    }
  });
}

function renderPictoLibrary() {
  pictoLibrary.innerHTML = "";
  Object.entries(pictoCatalog).forEach(([pictoKey, config]) => {
    pictoLibrary.appendChild(buildPictoButton(config, pictoKey));
  });
}

function attachElementEvents(element) {
  createResizeHandle(element);

  element.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) {
      return;
    }

    const layerId = element.dataset.layerId;
    const layer = getLayer(layerId);
    const shouldAddToSelection = event.ctrlKey || event.metaKey || event.shiftKey;
    const wasAlreadySelected = selectedLayerIds.has(layerId);
    const hadMultipleSelection = selectedLayerIds.size > 1;

    if (shouldAddToSelection) {
      suppressSelectionClickForLayerId = layerId;
      selectLayer(layerId, { additive: true });
      return;
    }

    if (!(wasAlreadySelected && hadMultipleSelection)) {
      suppressSelectionClickForLayerId = layerId;
      selectLayer(layerId);
    } else {
      board.querySelectorAll("[data-layer-id]").forEach((boardElement) => {
        boardElement.classList.toggle("is-selected", selectedLayerIds.has(boardElement.dataset.layerId));
      });
      renderLayers();
      refreshSelectedControls();
    }

    if (!layer || layer.locked) {
      return;
    }

    pushHistory();

    const boardRect = board.getBoundingClientRect();
    const rect = element.getBoundingClientRect();
    const resizeDirection = event.target.classList.contains("resize-handle")
      ? event.target.dataset.resize || "bottom-right"
      : "";

    interaction = {
      mode: resizeDirection ? "resize" : "move",
      resizeDirection,
      layerId,
      startX: event.clientX,
      startY: event.clientY,
      startLeft: parseFloat(element.style.left || "0"),
      startTop: parseFloat(element.style.top || "0"),
      startWidth: parseFloat(element.style.width || "12"),
      startHeight: parseFloat(element.style.height || "12"),
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top,
      boardWidth: boardRect.width,
      boardHeight: boardRect.height,
      selectedLayerIds: !resizeDirection && wasAlreadySelected ? [...selectedLayerIds] : [layerId],
      initialPositions: !resizeDirection && wasAlreadySelected
        ? [...selectedLayerIds].map((selectedId) => {
          const selectedElement = getElementByLayerId(selectedId);
          return selectedElement ? {
            layerId: selectedId,
            left: parseFloat(selectedElement.style.left || "0"),
            top: parseFloat(selectedElement.style.top || "0"),
          } : null;
        }).filter(Boolean)
        : [{ layerId, left: parseFloat(element.style.left || "0"), top: parseFloat(element.style.top || "0") }],
    };

    element.setPointerCapture(event.pointerId);
  });

  element.addEventListener("click", (event) => {
    if (suppressSelectionClickForLayerId === element.dataset.layerId) {
      suppressSelectionClickForLayerId = "";
      return;
    }
    selectLayer(element.dataset.layerId, { additive: event.ctrlKey || event.metaKey || event.shiftKey });
  });
}

function createPictoElement(pictoKey) {
  const config = pictoCatalog[pictoKey];
  if (!config) {
    return;
  }

  pushHistory();
  customIndex += 1;
  const layerId = `picto-${customIndex}`;
  const element = document.createElement("div");
  element.className = "board-element board-element--picto";
  element.dataset.layerId = layerId;
  element.dataset.type = "picto";
  element.style.left = `${20 + customIndex * 2}%`;
  element.style.top = `${22 + customIndex * 2}%`;
  element.style.width = "12%";
  element.style.height = "12%";
  element.innerHTML = `
    <span class="board-picto-badge">${getPictoMarkup(config.src, config.label)}</span>
  `;

  board.appendChild(element);
  attachElementEvents(element);
  applyTextLayout(layerId);

  layers.push({
    id: layerId,
    type: "picto",
    label: config.label,
    src: config.src,
    locked: false,
    visible: true,
    listed: true,
  });

  applyLayerOrder();
  renderLayers();
  selectLayer(layerId);
}

function createTextElement() {
  pushHistory();
  customIndex += 1;
  const layerId = `text-${customIndex}`;
  const element = document.createElement("div");
  element.className = "board-element board-element--text";
  element.dataset.layerId = layerId;
  element.dataset.type = "text";
  element.style.left = `${18 + customIndex * 2}%`;
  element.style.top = `${18 + customIndex * 2}%`;
  element.style.width = "22%";
  element.style.height = "8%";
  element.innerHTML = "<p>Texte libre</p>";

  board.appendChild(element);
  attachElementEvents(element);

  layers.push({
    id: layerId,
    type: "text",
    label: "Texte libre",
    locked: false,
    visible: true,
    listed: true,
  });

  initializeTextStyleState(layerId);
  applyLayerOrder();
  renderLayers();
  selectLayer(layerId);
}

function createShapeElement(shapeType) {
  const shapeConfig = shapeCatalog[shapeType];
  if (!shapeConfig) {
    return;
  }

  pushHistory();
  customIndex += 1;
  const layerId = `shape-${customIndex}`;
  const element = document.createElement("div");
  element.className = "board-element board-element--shape";
  element.dataset.layerId = layerId;
  element.dataset.type = "shape";
  element.dataset.shapeType = shapeType;
  element.style.left = `${22 + customIndex * 1.2}%`;
  element.style.top = `${22 + customIndex * 1.2}%`;
  element.style.width = ["line", "black-bar"].includes(shapeType) ? "22%" : "10%";
  element.style.height = ["line", "black-bar"].includes(shapeType) ? "2.4%" : "10%";
  element.innerHTML = `<span class="board-shape board-shape--${shapeType}" aria-hidden="true"></span>`;

  board.appendChild(element);
  attachElementEvents(element);

  layers.push({
    id: layerId,
    type: "shape",
    label: shapeConfig.label,
    shapeType,
    locked: false,
    visible: true,
    listed: true,
  });

  initializeShapeStyleState(layerId);
  applyLayerOrder();
  renderLayers();
  selectLayer(layerId);
}

function cloneBaseLayer(type) {
  if (type === "text") {
    createTextElement();
    return;
  }

  const baseLayerId = `${type}-base`;
  const baseLayer = getLayer(baseLayerId);
  const baseElement = getElementByLayerId(baseLayerId);
  if (baseLayer && baseElement && !isLayerListed(baseLayer)) {
    pushHistory();
    baseLayer.visible = true;
    baseLayer.listed = true;
    baseElement.classList.remove("is-hidden");
    if (type === "title") {
      baseElement.querySelector("h2").textContent = titleInput.value || "Titre de la carte";
    }
    if (type === "source") {
      const sourceNode = getSourceTextNode(baseElement);
      if (sourceNode) {
        sourceNode.textContent = sourceInput.value || "Source : a completer";
      }
    }
    if (type === "north") {
      const northNode = baseElement.querySelector("span");
      if (northNode) {
        northNode.textContent = northLabelInput.value || "N";
      }
    }
    applyLayerOrder();
    renderLayers();
    selectLayer(baseLayerId);
    return;
  }

  const sourceElement = board.querySelector(`.board-element--${type}`);
  if (!sourceElement) {
    return;
  }

  pushHistory();
  customIndex += 1;
  const layerId = `${type}-${customIndex}`;
  const clone = sourceElement.cloneNode(true);
  clone.dataset.layerId = layerId;
  clone.style.left = `${12 + customIndex * 2}%`;
  clone.style.top = `${18 + customIndex * 2}%`;

  if (type === "title") {
    clone.querySelector("h2").id = "";
    clone.querySelector("h2").textContent = titleInput.value || "Titre de la carte";
    clone.style.width = "42%";
    clone.style.height = "16%";
  }

  if (type === "source") {
    const sourceNode = getSourceTextNode(clone);
    if (sourceNode) {
      sourceNode.id = "";
      sourceNode.textContent = sourceInput.value || "Source : a completer";
    }
    clone.style.width = "32%";
    clone.style.height = "10%";
  }

  if (type === "north") {
    clone.querySelector("span").id = "";
    clone.querySelector("span").textContent = northLabelInput.value || "N";
    clone.style.width = "10%";
    clone.style.height = "16%";
  }

  board.appendChild(clone);
  attachElementEvents(clone);
  if (["title", "source"].includes(type)) {
    initializeTextStyleState(layerId, getTextStyleState(`${type}-base`));
  }
  applyTextLayout(layerId);

  const labels = {
    title: "Titre secondaire",
    source: "Source secondaire",
    north: "Orientation secondaire",
  };

  layers.push({
    id: layerId,
    type,
    label: labels[type],
    locked: false,
    visible: true,
    listed: true,
  });

  applyLayerOrder();
  renderLayers();
  selectLayer(layerId);
}

function updateSelectedText(value) {
  if (legendSelection) {
    updateLegendSelectionText(value);
    return;
  }

  const layer = getLayer(selectedLayerId);
  const element = layer ? getElementByLayerId(layer.id) : null;
  if (!layer || !element || layer.locked) {
    return;
  }

  if (layer.id === "title-base") {
    titleInput.value = value;
    setBoardText(boardTitle, value, "Titre de la carte");
    return;
  }

  if (layer.id === "source-base") {
    sourceInput.value = value;
    setBoardText(boardSource, value, "Source : a completer");
    return;
  }

  if (layer.id === "north-base") {
    northLabelInput.value = value;
    setBoardText(northLabel, value, "N");
    return;
  }

  if (layer.type === "map") {
    return;
  }

  if (layer.type === "title") {
    element.querySelector("h2").textContent = value.trim() || "Titre de la carte";
    layer.label = value.trim() || "Titre secondaire";
  } else if (layer.type === "source") {
    const sourceNode = getSourceTextNode(element);
    if (sourceNode) {
      sourceNode.textContent = value.trim() || "Source : a completer";
    }
    layer.label = value.trim() || "Source secondaire";
  } else if (layer.type === "text") {
    const textNode = element.querySelector("p");
    if (textNode) {
      textNode.textContent = value.trim() || "Texte libre";
    }
    layer.label = value.trim() || "Texte libre";
  } else if (layer.type === "north") {
    element.querySelector("span").textContent = value.trim() || "N";
    layer.label = `Orientation ${value.trim() || "N"}`;
  } else {
    layer.label = value.trim() || "Pictogramme";
  }

  renderLayers();
}

function updateSelectedTextStyle(nextStyle) {
  if (legendSelection) {
    const sanitizedStyle = { ...nextStyle };
    if (typeof sanitizedStyle.fontSize === "number") {
      sanitizedStyle.fontSize = clamp(sanitizedStyle.fontSize, 4, 72);
    }
    updateLegendSelectionStyle(sanitizedStyle);
    return;
  }

  const selectedLayers = getUnlockedSelectedLayers();
  const shapeLayers = selectedLayers.filter((layer) => isShapeStyleLayer(layer));
  if (shapeLayers.length && !selectedLayers.some((layer) => isTextStyleLayer(layer))) {
    pushHistory();
    shapeLayers.forEach((layer) => {
      const shapeNext = {};
      if (typeof nextStyle.backgroundColor !== "undefined") {
        shapeNext.fillColor = nextStyle.backgroundColor;
        if (typeof nextStyle.bgTransparent === "undefined") {
          shapeNext.fillTransparent = false;
        }
      }
      if (typeof nextStyle.bgTransparent !== "undefined") {
        shapeNext.fillTransparent = nextStyle.bgTransparent;
      }
      applyShapeStylesToElement(layer.id, shapeNext);
    });
    refreshSelectedControls();
    return;
  }

  const targetLayers = getTextStyleTargetLayers();
  if (!targetLayers.length) {
    return;
  }

  const sanitizedStyle = { ...nextStyle };
  if (typeof sanitizedStyle.fontSize === "number") {
    sanitizedStyle.fontSize = clamp(sanitizedStyle.fontSize, 4, 72);
  }

  pushHistory();
  targetLayers.forEach((layer) => {
    applyTextStylesToElement(layer.id, sanitizedStyle);
  });
}

function updateSelectedRotation(value) {
  const targetLayers = getRotationTargetLayers();
  if (!targetLayers.length) {
    return;
  }

  pushHistory();
  targetLayers.forEach((layer) => {
    const element = getElementByLayerId(layer.id);
    if (!element) {
      return;
    }
    element.dataset.rotation = String(value);
    if (isTextStyleLayer(layer)) {
      syncElementTextStylePresentation(layer.id);
    } else {
      setInlineStyle(element, "transform", `rotate(${value}deg)`);
    }
  });
  refreshSelectedControls();
}

function updateSelectedWidth(value) {
  const targetLayers = getResizeTargetLayers();
  if (!targetLayers.length) {
    return;
  }

  targetLayers.forEach((layer) => {
    const element = getElementByLayerId(layer.id);
    if (!element) {
      return;
    }
    const minSize = getLayerMinSize(layer);
    element.style.width = `${clamp(Number(value), minSize, 95)}%`;
    applyTextLayout(layer.id);
  });
}

function updateSelectedHeight(value) {
  const targetLayers = getResizeTargetLayers();
  if (!targetLayers.length) {
    return;
  }

  targetLayers.forEach((layer) => {
    const element = getElementByLayerId(layer.id);
    if (!element) {
      return;
    }
    const minSize = getLayerMinSize(layer);
    element.style.height = `${clamp(Number(value), minSize, 95)}%`;
    applyTextLayout(layer.id);
  });
}

function removeSelectedLayer() {
  const selectedLayers = getSelectedLayers();
  const deletableLayers = selectedLayers.filter((layer) => !["map", "title", "source", "north"].includes(layer.type));
  if (!deletableLayers.length) {
    return;
  }

  pushHistory();
  deletableLayers.forEach((layer) => {
    const element = getElementByLayerId(layer.id);
    const index = getLayerIndex(layer.id);
    if (index !== -1) {
      layers.splice(index, 1);
    }
    if (element) {
      element.remove();
    }
    selectedLayerIds.delete(layer.id);
  });

  selectLayer("title-base");
  applyLayerOrder();
  renderLayers();
}

function updateInteraction(event) {
  if (!interaction) {
    return;
  }

  const layer = getLayer(interaction.layerId);
  const element = getElementByLayerId(interaction.layerId);
  if (!layer || !element || layer.locked) {
    return;
  }

  if (interaction.mode === "move") {
    const left = ((event.clientX - interaction.offsetX - board.getBoundingClientRect().left) / interaction.boardWidth) * 100;
    const top = ((event.clientY - interaction.offsetY - board.getBoundingClientRect().top) / interaction.boardHeight) * 100;
    const deltaLeft = left - interaction.startLeft;
    const deltaTop = top - interaction.startTop;

    interaction.initialPositions.forEach((position) => {
      const selectedElement = getElementByLayerId(position.layerId);
      const selectedLayer = getLayer(position.layerId);
      if (!selectedElement) {
        return;
      }
      const selectedWidth = parseFloat(selectedElement.style.width || "12");
      const selectedHeight = parseFloat(selectedElement.style.height || "12");
      const maxRightBoundary = selectedLayer?.type === "map" ? getMapWorkAreaWidthPercent() : 100;
      selectedElement.style.left = `${clamp(position.left + deltaLeft, 0, maxRightBoundary - selectedWidth)}%`;
      selectedElement.style.top = `${clamp(position.top + deltaTop, 0, 100 - selectedHeight)}%`;
    });
    return;
  }

  const deltaX = ((event.clientX - interaction.startX) / interaction.boardWidth) * 100;
  const deltaY = ((event.clientY - interaction.startY) / interaction.boardHeight) * 100;
  const minSize = getLayerMinSize(layer);
  const maxRightBoundary = layer.type === "map" ? getMapWorkAreaWidthPercent() : 100;
  let nextLeft = interaction.startLeft;
  let nextTop = interaction.startTop;
  let nextWidth = interaction.startWidth;
  let nextHeight = interaction.startHeight;
  const direction = interaction.resizeDirection || "bottom-right";

  if (direction.includes("right")) {
    nextWidth = clamp(interaction.startWidth + deltaX, minSize, maxRightBoundary - interaction.startLeft);
  }

  if (direction.includes("left")) {
    nextLeft = clamp(interaction.startLeft + deltaX, 0, interaction.startLeft + interaction.startWidth - minSize);
    nextWidth = clamp(interaction.startWidth - deltaX, minSize, maxRightBoundary);
    if (nextLeft + nextWidth > maxRightBoundary) {
      nextWidth = maxRightBoundary - nextLeft;
    }
  }

  if (direction.includes("bottom")) {
    nextHeight = clamp(interaction.startHeight + deltaY, minSize, 95 - interaction.startTop);
  }

  if (direction.includes("top")) {
    nextTop = clamp(interaction.startTop + deltaY, 0, interaction.startTop + interaction.startHeight - minSize);
    nextHeight = clamp(interaction.startHeight - deltaY, minSize, 95);
    if (nextTop + nextHeight > 100) {
      nextHeight = 100 - nextTop;
    }
  }

  if (["top", "bottom"].includes(direction)) {
    nextWidth = interaction.startWidth;
  }

  if (["left", "right"].includes(direction)) {
    nextHeight = interaction.startHeight;
  }

  element.style.left = `${clamp(nextLeft, 0, maxRightBoundary - minSize)}%`;
  element.style.top = `${clamp(nextTop, 0, 100 - minSize)}%`;
  element.style.width = `${clamp(nextWidth, minSize, maxRightBoundary - nextLeft)}%`;
  element.style.height = `${clamp(nextHeight, minSize, 100 - nextTop)}%`;
  applyTextLayout(layer.id);
  refreshSelectedControls();
}

function endInteraction() {
  interaction = null;
  refreshSelectedControls();
}

function readFileAsDataUrl(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(file);
  });
}

function waitForImageLoad(image, src) {
  return new Promise((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error("image-load-failed"));
    image.src = src;
  });
}

async function importCustomPictos(files) {
  const fileList = [...files];
  for (const file of fileList) {
    customIndex += 1;
    const key = `custom-library-${customIndex}`;
    const dataUrl = await readFileAsDataUrl(file);
    pictoCatalog[key] = {
      label: file.name.replace(/\.[^.]+$/, ""),
      src: dataUrl,
    };
  }

  renderPictoLibrary();
}

function resetWorkspace() {
  pushHistory();
  mapImage.src = "";
  mapImageDataUrl = "";
  mapLayer.classList.remove("has-image");
  resetMapLayerFrame();
  mapLayer.style.visibility = "visible";
  emptyMapState.hidden = false;
  updateStatusMessage("Aucune image importee pour l'instant.");
  mapUpload.value = "";

  layers
    .filter((layer) => !["map", "title", "source", "north"].includes(layer.type))
    .forEach((layer) => {
      const element = getElementByLayerId(layer.id);
      if (element) {
        element.remove();
      }
    });

  while (layers.length > 4) {
    layers.pop();
  }

  const titleElement = getElementByLayerId("title-base");
  const sourceElement = getElementByLayerId("source-base");
  const northElement = getElementByLayerId("north-base");

  titleElement.style.left = "0%";
  titleElement.style.top = "0%";
  titleElement.style.width = "34%";
  titleElement.style.height = "6.5%";
  initializeTextStyleState("title-base");

  sourceElement.style.left = "0.25%";
  sourceElement.style.top = "18%";
  sourceElement.style.width = "2.2%";
  sourceElement.style.height = "22%";
  initializeTextStyleState("source-base");
  sourceElement.dataset.rotation = "180";

  northElement.style.left = "91%";
  northElement.style.top = "8%";
  northElement.style.width = "6%";
  northElement.style.height = "11%";

  layers.forEach((layer) => {
    layer.visible = false;
    layer.locked = false;
    layer.listed = false;
  });

  board.querySelectorAll("[data-layer-id]").forEach((element) => {
    element.classList.remove("is-hidden", "is-locked");
    if (element.dataset.layerId) {
      element.classList.add("is-hidden");
    }
  });

  titleInput.value = "Titre de la carte";
  subtitleInput.value = "Sous-titre ou precision territoriale";
  sourceInput.value = "Source : a completer";
  northLabelInput.value = "N";
  legendSectionsState = [];

  applyTextLayout("title-base");
  applyTextLayout("source-base");
  clearSelection();
  syncBaseTexts();
  applyLayerOrder();
  renderLayers();
}

function setExportState(isBusy, label = "Exporter PNG") {
  exportImageButton.disabled = isBusy;
  exportImageButton.textContent = label;
}

function updateStatusMessage(message, assertive = false) {
  uploadStatus.textContent = message;
  uploadStatus.setAttribute("aria-live", assertive ? "assertive" : "polite");
}

function downloadCanvas(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("png-blob-generation-failed"));
        return;
      }

      const link = document.createElement("a");
      const objectUrl = URL.createObjectURL(blob);
      link.href = objectUrl;
      link.download = "mise-en-page-cartographique.png";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(objectUrl);
      resolve();
    }, "image/png");
  });
}

function drawWrappedText(context, text, x, y, maxWidth, lineHeight) {
  const words = text.split(" ");
  let line = "";
  let currentY = y;

  words.forEach((word, index) => {
    const testLine = `${line}${word} `;
    if (context.measureText(testLine).width > maxWidth && index > 0) {
      context.fillText(line.trim(), x, currentY);
      line = `${word} `;
      currentY += lineHeight;
    } else {
      line = testLine;
    }
  });

  if (line.trim()) {
    context.fillText(line.trim(), x, currentY);
  }
}

function getWrappedLines(context, text, maxWidth) {
  const words = String(text || "").split(" ");
  const lines = [];
  let line = "";

  words.forEach((word, index) => {
    const testLine = `${line}${word} `;
    if (context.measureText(testLine).width > maxWidth && index > 0) {
      lines.push(line.trim());
      line = `${word} `;
    } else {
      line = testLine;
    }
  });

  if (line.trim()) {
    lines.push(line.trim());
  }

  return lines.length ? lines : [""];
}

function loadExportImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error(`Impossible de charger l'image: ${src}`));
    image.src = src;
  });
}

function drawImageCover(context, image, x, y, width, height) {
  const sourceWidth = image.naturalWidth || image.width;
  const sourceHeight = image.naturalHeight || image.height;
  if (!sourceWidth || !sourceHeight || !width || !height) {
    return;
  }

  const sourceRatio = sourceWidth / sourceHeight;
  const targetRatio = width / height;
  let sourceX = 0;
  let sourceY = 0;
  let cropWidth = sourceWidth;
  let cropHeight = sourceHeight;

  if (sourceRatio > targetRatio) {
    cropWidth = sourceHeight * targetRatio;
    sourceX = (sourceWidth - cropWidth) / 2;
  } else {
    cropHeight = sourceWidth / targetRatio;
    sourceY = (sourceHeight - cropHeight) / 2;
  }

  context.drawImage(image, sourceX, sourceY, cropWidth, cropHeight, x, y, width, height);
}

function drawImageContain(context, image, x, y, width, height) {
  const sourceWidth = image.naturalWidth || image.width;
  const sourceHeight = image.naturalHeight || image.height;
  if (!sourceWidth || !sourceHeight || !width || !height) {
    return;
  }

  const sourceRatio = sourceWidth / sourceHeight;
  const targetRatio = width / height;
  let drawWidth = width;
  let drawHeight = height;

  if (sourceRatio > targetRatio) {
    drawHeight = width / sourceRatio;
  } else {
    drawWidth = height * sourceRatio;
  }

  const drawX = x + (width - drawWidth) / 2;
  const drawY = y + (height - drawHeight) / 2;
  context.drawImage(image, drawX, drawY, drawWidth, drawHeight);
}

function mapDomRectToCanvas(domRect, boardRect, canvas) {
  const x = ((domRect.left - boardRect.left) / boardRect.width) * canvas.width;
  const y = ((domRect.top - boardRect.top) / boardRect.height) * canvas.height;
  const width = (domRect.width / boardRect.width) * canvas.width;
  const height = (domRect.height / boardRect.height) * canvas.height;
  return { x, y, width, height };
}

async function exportBoardAsImage() {
  setExportState(true, "Export en cours...");
  updateStatusMessage("Export PNG en cours...");

  try {
    const renderCanvas = document.createElement("canvas");
    const boardRect = board.getBoundingClientRect();
    renderCanvas.width = EXPORT_BASE_WIDTH * EXPORT_OUTPUT_SCALE;
    renderCanvas.height = EXPORT_BASE_HEIGHT * EXPORT_OUTPUT_SCALE;

    const context = renderCanvas.getContext("2d");
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = "high";
    context.fillStyle = "#fbfaf4";
    context.fillRect(0, 0, renderCanvas.width, renderCanvas.height);

    const mapRect = mapDomRectToCanvas(mapLayer.getBoundingClientRect(), boardRect, renderCanvas);
    const mapX = mapRect.x;
    const mapY = mapRect.y;
    const mapWidth = mapRect.width;
    const mapHeight = mapRect.height;

    context.strokeStyle = "#9096ba";
    context.lineWidth = 2;
    context.strokeRect(0, 0, renderCanvas.width, renderCanvas.height);

    context.fillStyle = "#ece7d7";
    context.fillRect(mapX, mapY, mapWidth, mapHeight);

    if (mapImageDataUrl && getLayer("map-base").visible && mapImage.complete && mapImage.naturalWidth > 0) {
      drawImageCover(context, mapImage, mapX, mapY, mapWidth, mapHeight);
    }

    await drawElementsOnCanvas(context, boardRect, renderCanvas);
    await drawLegendOnCanvas(context, boardRect, renderCanvas);

    await downloadCanvas(renderCanvas);
    updateStatusMessage("Export termine : fichier PNG telecharge.");
    setExportState(false);
  } catch (error) {
    setExportState(false);
    updateStatusMessage("Echec export PNG. Reessayez apres actualisation.", true);
    alert("L'export PNG a echoue. Rechargez la page puis reessayez.");
  }
}

async function drawLegendSymbol(context, entry, x, y, size) {
  if (entry.layerType === "picto" && entry.src) {
    try {
      const image = await loadExportImage(entry.src);
      drawImageContain(context, image, x, y, size, size);
      return;
    } catch (error) {
      context.fillStyle = "#67d59a";
      context.beginPath();
      context.arc(x + size / 2, y + size / 2, size / 2, 0, Math.PI * 2);
      context.fill();
      return;
    }
  }

  if (entry.layerType === "north") {
    context.fillStyle = "#000091";
    context.beginPath();
    context.moveTo(x + size / 2, y);
    context.lineTo(x, y + size);
    context.lineTo(x + size, y + size);
    context.closePath();
    context.fill();
    return;
  }

  context.fillStyle = "#111827";
  context.fillRect(x, y + size / 2 - 1.5, size, 3);
}

function getCanvasScale(boardRect, canvas) {
  return {
    x: canvas.width / Math.max(1, boardRect.width),
    y: canvas.height / Math.max(1, boardRect.height),
  };
}

function drawTextNodeInRect(context, node, text, rect, boardRect, canvas, options = {}) {
  if (!node || !text || !rect.width || !rect.height) {
    return;
  }

  const styles = window.getComputedStyle(node);
  const scale = getCanvasScale(boardRect, canvas);
  const fontScale = (scale.x + scale.y) / 2;
  const fontSize = Math.max(8, (parseFloat(styles.fontSize) || 12) * fontScale);
  const fontWeight = styles.fontWeight || "400";
  const fontStyle = styles.fontStyle || "normal";
  const fontFamily = styles.fontFamily || "Marianne, Arial, sans-serif";
  const lineHeightCss = parseFloat(styles.lineHeight);
  const lineHeight = Number.isFinite(lineHeightCss) ? Math.max(fontSize, lineHeightCss * scale.y) : Math.max(fontSize, fontSize * 1.2);
  const paddingX = options.paddingX ?? 0;
  const paddingY = options.paddingY ?? 0;
  const textAlign = options.textAlign || "left";
  const verticalAlign = options.verticalAlign || "top";
  const maxWidth = Math.max(2, rect.width - paddingX * 2);

  context.fillStyle = styles.color || "#161616";
  context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`;
  context.textAlign = textAlign;
  context.textBaseline = "top";

  context.save();
  context.beginPath();
  context.rect(rect.x, rect.y, rect.width, rect.height);
  context.clip();

  const lines = getWrappedLines(context, text, maxWidth);
  const maxLines = Math.max(1, Math.floor((rect.height - paddingY * 2) / lineHeight));
  const visibleLines = lines.slice(0, maxLines);
  const textBlockHeight = visibleLines.length * lineHeight;
  let startY = rect.y + paddingY;
  if (verticalAlign === "middle") {
    startY = rect.y + Math.max(paddingY, (rect.height - textBlockHeight) / 2);
  } else if (verticalAlign === "bottom") {
    startY = rect.y + Math.max(paddingY, rect.height - paddingY - textBlockHeight);
  }
  const textX = textAlign === "center"
    ? rect.x + (rect.width / 2)
    : textAlign === "right"
      ? rect.x + rect.width - paddingX
      : rect.x + paddingX;

  visibleLines.forEach((line, index) => {
    context.fillText(line, textX, startY + (index * lineHeight));
  });
  context.restore();
}

async function resolveLegendImageSource(symbolNode) {
  if (!symbolNode) {
    return null;
  }

  if (symbolNode.tagName === "IMG" && symbolNode.src) {
    return symbolNode.src;
  }

  const nestedImage = symbolNode.querySelector("img");
  if (nestedImage && nestedImage.src) {
    return nestedImage.src;
  }

  return null;
}

async function drawLegendSymbolFromNode(context, symbolNode, rect) {
  if (!symbolNode || rect.width <= 0 || rect.height <= 0) {
    return;
  }

  const imageSrc = await resolveLegendImageSource(symbolNode);
  if (imageSrc) {
    try {
      const image = await loadExportImage(imageSrc);
      drawImageContain(context, image, rect.x, rect.y, rect.width, rect.height);
      return;
    } catch (error) {
      // Fallback shape below
    }
  }

  if (symbolNode.classList.contains("carto-legend-symbol--area")) {
    context.fillStyle = "#000091";
    context.beginPath();
    context.moveTo(rect.x + rect.width / 2, rect.y);
    context.lineTo(rect.x, rect.y + rect.height);
    context.lineTo(rect.x + rect.width, rect.y + rect.height);
    context.closePath();
    context.fill();
    return;
  }

  if (symbolNode.classList.contains("carto-legend-symbol--point")) {
    const radius = Math.max(2, Math.min(rect.width, rect.height) / 2);
    context.fillStyle = "#67d59a";
    context.beginPath();
    context.arc(rect.x + rect.width / 2, rect.y + rect.height / 2, radius, 0, Math.PI * 2);
    context.fill();
    return;
  }

  context.fillStyle = "#111827";
  context.fillRect(rect.x, rect.y + (rect.height / 2) - 1, rect.width, 2);
}

async function drawLegendOnCanvas(context, boardRect, canvas) {
  if (!legendPanel) {
    return;
  }

  const panelRect = mapDomRectToCanvas(legendPanel.getBoundingClientRect(), boardRect, canvas);
  const panelStyles = window.getComputedStyle(legendPanel);
  context.fillStyle = panelStyles.backgroundColor || "#ffffff";
  context.fillRect(panelRect.x, panelRect.y, panelRect.width, panelRect.height);

  const legendCard = legendPanel.querySelector(".carto-legend-card");
  if (legendCard) {
    const cardRect = mapDomRectToCanvas(legendCard.getBoundingClientRect(), boardRect, canvas);
    const cardStyles = window.getComputedStyle(legendCard);
    context.fillStyle = cardStyles.backgroundColor || "#ffffff";
    context.fillRect(cardRect.x, cardRect.y, cardRect.width, cardRect.height);
  }

  const sections = [...legendPanel.querySelectorAll(".carto-legend-section")];
  for (const section of sections) {
    const titleRow = section.querySelector(".carto-legend-section__title-row");
    const titleNode = section.querySelector(".carto-legend-section__title");
    const bodyNode = section.querySelector(".carto-legend-section__body");

    if (titleRow) {
      const titleRowRect = mapDomRectToCanvas(titleRow.getBoundingClientRect(), boardRect, canvas);
      const titleRowStyles = window.getComputedStyle(titleRow);
      context.fillStyle = titleRowStyles.backgroundColor || "#000091";
      context.fillRect(titleRowRect.x, titleRowRect.y, titleRowRect.width, titleRowRect.height);
    }

    if (titleNode) {
      const titleRect = mapDomRectToCanvas(titleNode.getBoundingClientRect(), boardRect, canvas);
      drawTextNodeInRect(context, titleNode, titleNode.textContent?.trim() || "", titleRect, boardRect, canvas, {
        paddingX: 2,
        paddingY: 1,
        textAlign: "left",
        verticalAlign: "middle",
      });
    }

    // Intentionally skip legend action buttons (U/D/R) in export output.

    if (bodyNode) {
      const bodyRect = mapDomRectToCanvas(bodyNode.getBoundingClientRect(), boardRect, canvas);
      const bodyStyles = window.getComputedStyle(bodyNode);
      context.fillStyle = bodyStyles.backgroundColor || "#ffffff";
      context.fillRect(bodyRect.x, bodyRect.y, bodyRect.width, bodyRect.height);
    }

  const items = [...section.querySelectorAll(".carto-legend-list li")];
  for (const item of items) {
      const symbolNode = item.querySelector(".carto-legend-symbol, .carto-legend-picto");
      const labelNode = item.querySelector(".carto-legend-label");
      if (symbolNode) {
        const symbolRect = mapDomRectToCanvas(symbolNode.getBoundingClientRect(), boardRect, canvas);
        await drawLegendSymbolFromNode(context, symbolNode, symbolRect);
      }
      if (labelNode) {
        const labelRect = mapDomRectToCanvas(labelNode.getBoundingClientRect(), boardRect, canvas);
        drawTextNodeInRect(context, labelNode, labelNode.textContent?.trim() || "", labelRect, boardRect, canvas);
      }
    }
  }

  if (legendHint && legendHint.offsetParent !== null) {
    const hintRect = mapDomRectToCanvas(legendHint.getBoundingClientRect(), boardRect, canvas);
    drawTextNodeInRect(context, legendHint, legendHint.textContent?.trim() || "", hintRect, boardRect, canvas);
  }

  // Draw separator at the very end so it stays visible.
  context.strokeStyle = "#9096ba";
  context.lineWidth = 1;
  context.beginPath();
  context.moveTo(panelRect.x + 0.5, 0);
  context.lineTo(panelRect.x + 0.5, canvas.height);
  context.stroke();
}

async function drawElementsOnCanvas(context, boardRect, canvas) {
  for (const layer of layers) {
    if (!layer.visible || layer.type === "map") {
      continue;
    }

    const element = getElementByLayerId(layer.id);
    if (!element) {
      continue;
    }

    const elementRect = mapDomRectToCanvas(element.getBoundingClientRect(), boardRect, canvas);
    const x = elementRect.x;
    const y = elementRect.y;
    const width = elementRect.width;
    const height = elementRect.height;
    const elementStyles = window.getComputedStyle(element);

    if (layer.type === "picto") {
      if (layer.src) {
        try {
          const image = await loadExportImage(layer.src);
          const badge = element.querySelector(".board-picto-badge");
          if (badge) {
            const badgeRect = mapDomRectToCanvas(badge.getBoundingClientRect(), boardRect, canvas);
            drawImageContain(context, image, badgeRect.x, badgeRect.y, badgeRect.width, badgeRect.height);
          } else {
            drawImageContain(context, image, x, y, width, height);
          }
        } catch (error) {
          context.fillStyle = "#000091";
          context.fillRect(x, y, width, height);
        }
      }
      continue;
    }

    if (layer.type === "shape") {
      const shapeType = element.dataset.shapeType || "rectangle";
      const shapeStyle = getShapeStyleState(layer.id);
      const fillColor = shapeStyle?.fillTransparent ? "transparent" : (shapeStyle?.fillColor || "rgba(0, 0, 145, 0.22)");
      const strokeColor = shapeStyle?.strokeColor || "#000091";
      context.save();
      context.translate(x + width / 2, y + height / 2);
      const rotation = Number(element.dataset.rotation || "0");
      context.rotate((rotation * Math.PI) / 180);
      context.fillStyle = fillColor;
      context.strokeStyle = strokeColor;
      context.lineWidth = 1;

      if (shapeType === "ellipse") {
        context.beginPath();
        context.ellipse(0, 0, width / 2, height / 2, 0, 0, Math.PI * 2);
        context.fill();
        context.stroke();
      } else if (shapeType === "triangle") {
        context.beginPath();
        context.moveTo(0, -height / 2);
        context.lineTo(-width / 2, height / 2);
        context.lineTo(width / 2, height / 2);
        context.closePath();
        context.fill();
        context.stroke();
      } else if (shapeType === "diamond") {
        context.beginPath();
        context.moveTo(0, -height / 2);
        context.lineTo(width / 2, 0);
        context.lineTo(0, height / 2);
        context.lineTo(-width / 2, 0);
        context.closePath();
        context.fill();
        context.stroke();
      } else if (shapeType === "line") {
        context.fillStyle = fillColor === "transparent" ? strokeColor : fillColor;
        context.fillRect(-width / 2, -1, width, 2);
      } else if (shapeType === "black-bar") {
        context.fillStyle = fillColor === "transparent" ? "#000000" : fillColor;
        context.fillRect(-width / 2, -height / 2, width, height);
      } else if (shapeType === "bubble-left" || shapeType === "bubble-bottom") {
        context.beginPath();
        context.ellipse(0, 0, width / 2, height / 2, 0, 0, Math.PI * 2);
        context.fillStyle = fillColor;
        context.fill();
        context.stroke();
        context.beginPath();
        if (shapeType === "bubble-left") {
          context.moveTo(-width * 0.5, height * 0.1);
          context.lineTo(-width * 0.62, height * 0.28);
          context.lineTo(-width * 0.44, height * 0.26);
        } else {
          context.moveTo(width * 0.02, height * 0.48);
          context.lineTo(width * 0.16, height * 0.66);
          context.lineTo(width * 0.26, height * 0.44);
        }
        context.closePath();
        context.fillStyle = fillColor;
        context.fill();
        context.stroke();
      } else if (shapeType === "callout-rect-left" || shapeType === "callout-rect-right" || shapeType === "notch-top" || shapeType === "notch-bottom") {
        context.fillStyle = fillColor;
        context.fillRect(-width / 2, -height / 2, width, height);
        context.strokeRect(-width / 2, -height / 2, width, height);
        context.beginPath();
        if (shapeType === "callout-rect-left") {
          context.moveTo(-width / 2, height * 0.1);
          context.lineTo(-width * 0.64, height * 0.24);
          context.lineTo(-width / 2, height * 0.32);
        } else if (shapeType === "callout-rect-right") {
          context.moveTo(width / 2, height * 0.1);
          context.lineTo(width * 0.64, height * 0.24);
          context.lineTo(width / 2, height * 0.32);
        } else if (shapeType === "notch-top") {
          context.moveTo(-width * 0.08, -height / 2);
          context.lineTo(0, -height * 0.7);
          context.lineTo(width * 0.08, -height / 2);
        } else {
          context.moveTo(-width * 0.08, height / 2);
          context.lineTo(0, height * 0.7);
          context.lineTo(width * 0.08, height / 2);
        }
        context.closePath();
        context.fillStyle = fillColor;
        context.fill();
        context.stroke();
      } else if (shapeType === "arrow-right") {
        const arrowWidth = width;
        const arrowHeight = height;
        context.fillStyle = "#000091";
        context.beginPath();
        context.moveTo(-arrowWidth / 2, -arrowHeight * 0.15);
        context.lineTo(arrowWidth * 0.2, -arrowHeight * 0.15);
        context.lineTo(arrowWidth * 0.2, -arrowHeight * 0.35);
        context.lineTo(arrowWidth / 2, 0);
        context.lineTo(arrowWidth * 0.2, arrowHeight * 0.35);
        context.lineTo(arrowWidth * 0.2, arrowHeight * 0.15);
        context.lineTo(-arrowWidth / 2, arrowHeight * 0.15);
        context.closePath();
        context.fill();
      } else {
        context.fillRect(-width / 2, -height / 2, width, height);
        context.strokeRect(-width / 2, -height / 2, width, height);
      }
      context.restore();
      continue;
    }

    if (layer.type === "north") {
      context.fillStyle = "#000091";
      context.font = "700 34px sans-serif";
      context.textAlign = "center";
      context.fillText(element.textContent.trim() || "N", x + width / 2, y + 28);
      context.beginPath();
      context.moveTo(x + width / 2, y + height - 6);
      context.lineTo(x + width / 2 - 18, y + height - 42);
      context.lineTo(x + width / 2 + 18, y + height - 42);
      context.closePath();
      context.fill();
      continue;
    }

    if (layer.type === "title") {
      const titleNode = element.querySelector("h2");
      const titleStyles = window.getComputedStyle(titleNode);
      const backgroundColor = element.dataset.bgTransparent === "true" ? "transparent" : elementStyles.backgroundColor;
      const rotation = Number(element.dataset.rotation || "0");

      context.save();
      context.translate(x + width / 2, y + height / 2);
      context.rotate((rotation * Math.PI) / 180);
      if (backgroundColor !== "transparent" && backgroundColor !== "rgba(0, 0, 0, 0)") {
        context.fillStyle = backgroundColor;
        context.fillRect(-width / 2, -height / 2, width, height);
      }

      context.fillStyle = titleStyles.color;
      const titleFontSize = Math.max(8, Math.round(parseFloat(titleStyles.fontSize) * (canvas.width / boardRect.width)));
      context.font = `${titleStyles.fontStyle} ${titleStyles.fontWeight} ${titleFontSize}px ${titleStyles.fontFamily}`;
      context.textAlign = "left";
      context.textBaseline = "top";
      const lineHeight = Math.max(12, titleFontSize * 1.15);
      const text = titleNode.textContent.trim();
      const textAreaWidth = Math.max(8, width - 28);
      const lines = getWrappedLines(context, text, textAreaWidth);
      const textBlockHeight = lines.length * lineHeight;
      const startY = (-height / 2) + Math.max(4, (height - textBlockHeight) / 2);
      drawWrappedText(context, text, -width / 2 + 14, startY, textAreaWidth, lineHeight);
      context.restore();
      continue;
    }

    if (layer.type === "source") {
      const sourceNode = getSourceTextNode(element);
      const sourceStyles = sourceNode ? window.getComputedStyle(sourceNode) : null;
      const backgroundColor = element.dataset.bgTransparent === "true" ? "transparent" : elementStyles.backgroundColor;
      const rotation = Number(element.dataset.rotation || "180");

      context.save();
      context.translate(x + width / 2, y + height / 2);
      // Source text is vertical in UI (writing-mode: vertical-rl) with an extra element rotation.
      // Approximate this in canvas by rotating a continuous text line.
      context.rotate(((rotation - 90) * Math.PI) / 180);

      if (backgroundColor !== "transparent" && backgroundColor !== "rgba(0, 0, 0, 0)") {
        context.fillStyle = backgroundColor;
        context.fillRect(-width / 2, -height / 2, width, height);
      }

      const fontSizePx = sourceStyles
        ? Math.max(8, Math.round(parseFloat(sourceStyles.fontSize) * (canvas.width / boardRect.width)))
        : 14;
      context.fillStyle = sourceStyles ? sourceStyles.color : "#5c5c5c";
      context.font = sourceStyles
        ? `${sourceStyles.fontStyle} ${sourceStyles.fontWeight} ${fontSizePx}px ${sourceStyles.fontFamily}`
        : "400 14px sans-serif";
      context.textAlign = "left";
      context.textBaseline = "middle";

      const text = sourceNode ? sourceNode.textContent.trim() : "";
      const measuredWidth = context.measureText(text).width;
      const textStartX = -measuredWidth / 2;

      context.beginPath();
      context.rect(-width / 2, -height / 2, width, height);
      context.clip();
      context.fillText(text, textStartX, 0);
      context.restore();
      continue;
    }

    if (layer.type === "text") {
      const textNode = element.querySelector("p");
      const textStyles = textNode ? window.getComputedStyle(textNode) : null;
      const backgroundColor = element.dataset.bgTransparent === "true" ? "transparent" : elementStyles.backgroundColor;
      const rotation = Number(element.dataset.rotation || "0");

      context.save();
      context.translate(x + width / 2, y + height / 2);
      context.rotate((rotation * Math.PI) / 180);

      if (backgroundColor !== "transparent" && backgroundColor !== "rgba(0, 0, 0, 0)") {
        context.fillStyle = backgroundColor;
        context.fillRect(-width / 2, -height / 2, width, height);
      }

      context.fillStyle = textStyles ? textStyles.color : "#161616";
      context.font = textStyles
        ? `${textStyles.fontStyle} ${textStyles.fontWeight} ${Math.round(parseFloat(textStyles.fontSize) * (canvas.width / boardRect.width))}px ${textStyles.fontFamily}`
        : "400 18px sans-serif";
      context.textAlign = "left";
      drawWrappedText(context, textNode ? textNode.textContent.trim() : "", -width / 2 + 8, 8, width - 16, 20);
      context.restore();
      continue;
    }

    context.fillStyle = "#ffffff";
    context.fillRect(x, y, width, height);
    context.strokeStyle = "rgba(0,0,145,0.22)";
    context.lineWidth = 2;
    context.strokeRect(x, y, width, height);
    context.fillStyle = "#3a3a3a";
    context.font = "400 18px sans-serif";
    const sourceNode = getSourceTextNode(element);
    drawWrappedText(context, sourceNode ? sourceNode.textContent.trim() : "", x + 16, y + 44, width - 32, 22);
  }
}

mapUpload.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) {
    return;
  }

  try {
    const dataUrl = await readFileAsDataUrl(file);
    await waitForImageLoad(mapImage, dataUrl);
    mapImageDataUrl = dataUrl;
    mapLayer.classList.add("has-image");
    resetMapLayerFrame();
    const mapBaseLayer = getLayer("map-base");
    if (mapBaseLayer) {
      mapBaseLayer.visible = true;
      mapBaseLayer.listed = true;
    }
    mapLayer.classList.remove("is-hidden");
    emptyMapState.hidden = true;
    updateStatusMessage(`Image chargee : ${file.name}`);
    renderLayers();
  } catch (error) {
    mapImageDataUrl = "";
    mapImage.removeAttribute("src");
    mapLayer.classList.remove("has-image");
    emptyMapState.hidden = false;
    updateStatusMessage("Le fichier n'a pas pu etre charge comme carte.", true);
  }
});

customPictoUpload.addEventListener("change", async (event) => {
  if (!event.target.files.length) {
    return;
  }
  await importCustomPictos(event.target.files);
  customPictoUpload.value = "";
});

titleInput.addEventListener("input", syncBaseTexts);
subtitleInput.addEventListener("input", syncBaseTexts);
sourceInput.addEventListener("input", syncBaseTexts);
northLabelInput.addEventListener("input", syncBaseTexts);

selectedText.addEventListener("input", (event) => {
  updateSelectedText(event.target.value);
});

legendSelectedText?.addEventListener("input", (event) => {
  if (!legendSelection) {
    return;
  }
  updateLegendSelectionText(event.target.value);
});

["input", "change"].forEach((eventName) => {
  selectedFontSize.addEventListener(eventName, (event) => {
    updateSelectedTextStyle({ fontSize: Number(event.target.value) });
  });

  selectedTextColor.addEventListener(eventName, (event) => {
    updateSelectedTextStyle({ textColor: event.target.value });
  });

  selectedTextBackground.addEventListener(eventName, (event) => {
    updateSelectedTextStyle({
      backgroundColor: event.target.value,
      bgTransparent: false,
    });
  });

  selectedFontFamily.addEventListener(eventName, (event) => {
    updateSelectedTextStyle({ fontFamily: event.target.value });
  });

  legendSelectedFontSize?.addEventListener(eventName, (event) => {
    if (!legendSelection) {
      return;
    }
    updateSelectedTextStyle({ fontSize: Number(event.target.value) });
  });

  legendSelectedTextColor?.addEventListener(eventName, (event) => {
    if (!legendSelection) {
      return;
    }
    updateSelectedTextStyle({ textColor: event.target.value });
  });

  legendSelectedTextBackground?.addEventListener(eventName, (event) => {
    if (!legendSelection) {
      return;
    }
    updateSelectedTextStyle({
      backgroundColor: event.target.value,
      bgTransparent: false,
    });
  });

  legendSelectedFontFamily?.addEventListener(eventName, (event) => {
    if (!legendSelection) {
      return;
    }
    updateSelectedTextStyle({ fontFamily: event.target.value });
  });
});

toggleBoldButton.addEventListener("click", () => {
  const styleState = legendSelection ? getLegendSelectionStyleState() : getTextStyleState(selectedLayerId);
  if (!styleState) {
    return;
  }
  updateSelectedTextStyle({ fontWeight: Number(styleState.fontWeight) >= 600 ? "400" : "700" });
});

toggleItalicButton.addEventListener("click", () => {
  const styleState = legendSelection ? getLegendSelectionStyleState() : getTextStyleState(selectedLayerId);
  if (!styleState) {
    return;
  }
  updateSelectedTextStyle({ fontStyle: styleState.fontStyle === "italic" ? "normal" : "italic" });
});

toggleTransparentBackgroundButton.addEventListener("click", () => {
  if (legendSelection) {
    const legendStyle = getLegendSelectionStyleState();
    if (!legendStyle) {
      return;
    }
    updateSelectedTextStyle({ bgTransparent: !legendStyle.bgTransparent });
    return;
  }

  const shapeState = getShapeStyleState(selectedLayerId);
  if (shapeState) {
    updateSelectedTextStyle({ bgTransparent: !shapeState.fillTransparent });
    return;
  }

  const textStyle = getTextStyleState(selectedLayerId);
  if (!textStyle) {
    return;
  }
  updateSelectedTextStyle({ bgTransparent: !textStyle.bgTransparent });
});

legendToggleBoldButton?.addEventListener("click", () => {
  if (!legendSelection) {
    return;
  }
  const styleState = getLegendSelectionStyleState();
  if (!styleState) {
    return;
  }
  updateSelectedTextStyle({ fontWeight: Number(styleState.fontWeight) >= 600 ? "400" : "700" });
});

legendToggleItalicButton?.addEventListener("click", () => {
  if (!legendSelection) {
    return;
  }
  const styleState = getLegendSelectionStyleState();
  if (!styleState) {
    return;
  }
  updateSelectedTextStyle({ fontStyle: styleState.fontStyle === "italic" ? "normal" : "italic" });
});

legendToggleTransparentBackgroundButton?.addEventListener("click", () => {
  if (!legendSelection) {
    return;
  }
  const styleState = getLegendSelectionStyleState();
  if (!styleState) {
    return;
  }
  updateSelectedTextStyle({ bgTransparent: !styleState.bgTransparent });
});

selectedWidth.addEventListener("input", (event) => {
  updateSelectedWidth(event.target.value);
});

selectedHeight.addEventListener("input", (event) => {
  updateSelectedHeight(event.target.value);
});

selectedRotation.addEventListener("input", (event) => {
  updateSelectedRotation(event.target.value);
});

legendColumns.addEventListener("input", (event) => {
  updateLegendLayoutStyle({ columns: event.target.value });
});

legendSymbolSize.addEventListener("input", (event) => {
  updateLegendLayoutStyle({ symbolSize: event.target.value });
});

bringForwardButton.addEventListener("click", () => moveLayer(selectedLayerId, 1));
sendBackwardButton.addEventListener("click", () => moveLayer(selectedLayerId, -1));
deleteLayerButton.addEventListener("click", removeSelectedLayer);
exportImageButton.addEventListener("click", exportBoardAsImage);
resetWorkspaceButton.addEventListener("click", resetWorkspace);

toggleGridButton.addEventListener("click", () => {
  boardGrid.classList.toggle("is-hidden");
});

zoomInButton.addEventListener("click", () => {
  zoomLevel = Math.min(ZOOM_MAX, zoomLevel + ZOOM_STEP);
  updateZoom();
});

zoomOutButton.addEventListener("click", () => {
  zoomLevel = Math.max(ZOOM_MIN, zoomLevel - ZOOM_STEP);
  updateZoom();
});

addButtons.forEach((button) => {
  button.addEventListener("click", () => {
    cloneBaseLayer(button.dataset.add);
  });
});

shapeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    createShapeElement(button.dataset.shape);
  });
});

document.addEventListener("pointermove", (event) => {
  if (updateLegendResize(event)) {
    return;
  }
  updateInteraction(event);
});

document.addEventListener("pointerup", (event) => {
  endLegendResize(event);
  endInteraction();
});

legendResizer?.addEventListener("pointerdown", startLegendResize);

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && displaySettings && !displaySettings.hidden) {
    closeDisplaySettings();
    return;
  }

  const target = event.target;
  const isTypingContext = target instanceof HTMLElement && (
    target.tagName === "INPUT" ||
    target.tagName === "TEXTAREA" ||
    target.isContentEditable
  );

  if (isTypingContext) {
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    const payload = buildCopiedLayerPayload(selectedLayerId);
    if (!payload) {
      return;
    }

    copiedLayerPayload = payload;
    event.preventDefault();
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
    if (!copiedLayerPayload) {
      return;
    }

    pasteCopiedLayer();
    event.preventDefault();
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
    undoLastAction();
    event.preventDefault();
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
    selectAllLayers();
    event.preventDefault();
  }

  if (event.key === "Delete" || event.key === "Backspace") {
    const selectedLayers = getSelectedLayers();
    const hasDeletableSelection = selectedLayers.some((layer) => !["map", "title", "source", "north"].includes(layer.type));
    if (!hasDeletableSelection) {
      return;
    }

    removeSelectedLayer();
    event.preventDefault();
  }
});

board.querySelectorAll(".board-element").forEach((element) => attachElementEvents(element));
attachElementEvents(mapLayer);
resetMapLayerFrame();
initializeTextStyleState("title-base");
initializeTextStyleState("source-base");
initializeToolboxMenu();
initializeSideMenu();
initializeLeftPanelToggle();
initializeDisplaySettings();
initializeLegendToolbar();
window.addEventListener("resize", scheduleMapLayerFrameSync);

if (typeof ResizeObserver !== "undefined") {
  const boardResizeObserver = new ResizeObserver(() => {
    scheduleMapLayerFrameSync();
  });
  boardResizeObserver.observe(board);
  if (legendPanel) {
    boardResizeObserver.observe(legendPanel);
  }
}

board.addEventListener("click", (event) => {
  if (event.target === board || event.target === boardGrid) {
    clearSelection();
  }
});

renderPictoLibrary();
renderFavoritePictos();
applyLayerOrder();
renderLayers();
syncBaseTexts();
updateZoom();
scheduleMapLayerFrameSync();
