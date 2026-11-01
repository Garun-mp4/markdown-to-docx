const form = document.querySelector("#convert-form");
const sourceInput = document.querySelector("#source");
const sourceDropzone = document.querySelector("#source-dropzone");
const sourceName = document.querySelector("#source-name");
const sourceFilename = document.querySelector("#source-filename");
const sourceError = document.querySelector("#source-error");
const assetsDropzone = document.querySelector("#assets-dropzone");
const assetsError = document.querySelector("#assets-error");
const assetFilesInput = document.querySelector("#asset-files");
const assetFolderInput = document.querySelector("#asset-folder");
const assetZipInput = document.querySelector("#asset-zip");
const assetList = document.querySelector("#asset-list");
const feedback = document.querySelector("#feedback");
const button = document.querySelector("#convert-button");
const resolution = document.querySelector("#resolution");
const markdownChoice = document.querySelector("#markdown-choice");
const mainPath = document.querySelector("#main-path");
const imagesDialog = document.querySelector("#images-dialog");
const result = document.querySelector("#result");
const downloadLink = document.querySelector("#download-link");
const resultFilename = document.querySelector("#result-filename");
const profileControls = document.querySelector("#profile-controls");
const profileError = document.querySelector("#profile-error");
const numberingStartPage = document.querySelector("#numbering-start-page");
const numberingMode = document.querySelector("#numbering-mode");
const numberingControls = document.querySelector("#numbering-controls");
const defaultProfile = JSON.parse(document.querySelector("#default-profile").textContent);
const validImageExtensions = new Set([".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"]);

let selectedSource = null;
let selectedAssets = [];
let downloadUrl = null;
let isConverting = false;

const profileLabels = {
  width_mm: "Ширина листа, мм",
  height_mm: "Высота листа, мм",
  left_margin_mm: "Левое поле, мм",
  right_margin_mm: "Правое поле, мм",
  top_margin_mm: "Верхнее поле, мм",
  bottom_margin_mm: "Нижнее поле, мм",
  header_distance_mm: "Отступ верхнего колонтитула, мм",
  footer_distance_mm: "Отступ нижнего колонтитула, мм",
  font_family: "Гарнитура",
  font_size_pt: "Размер шрифта, pt",
  color: "Цвет текста",
  alignment: "Выравнивание",
  line_spacing: "Межстрочный интервал",
  first_line_indent_mm: "Отступ первой строки, мм",
  space_before_pt: "Интервал перед, pt",
  space_after_pt: "Интервал после, pt",
  bold: "Полужирное начертание",
  italic: "Курсивное начертание",
  left_indent_mm: "Отступ слева, мм",
  depth_indent_mm: "Дополнительный отступ уровня, мм",
  hanging_indent_mm: "Выступ маркера, мм",
  right_indent_mm: "Отступ справа, мм",
  keep_with_next: "Оставлять со следующим абзацем",
  keep_together: "Не разрывать абзац между страницами",
  border_color: "Цвет линии",
  border_width_pt: "Толщина линии, pt",
  background_color: "Цвет фона",
  vertical_alignment: "Вертикальное выравнивание ячеек",
  padding_top_mm: "Внутренний отступ сверху, мм",
  padding_horizontal_mm: "Внутренний отступ по горизонтали, мм",
  padding_bottom_mm: "Внутренний отступ снизу, мм",
  max_width_percent: "Максимальная ширина изображения, %",
  underline: "Подчёркивание ссылок",
  text_alignment: "Выравнивание текста",
  page_alignment: "Выравнивание номера страницы",
  header_font_family: "Шрифт заголовка таблицы",
  header_font_size_pt: "Размер шрифта заголовка, pt",
  header_color: "Цвет заголовка таблицы",
  header_bold: "Полужирный заголовок таблицы",
  header_alignment: "Выравнивание заголовка таблицы",
  body_font_family: "Шрифт ячеек таблицы",
  body_font_size_pt: "Размер шрифта ячеек, pt",
  body_color: "Цвет текста ячеек",
  body_alignment: "Выравнивание ячеек",
};

const profileGroups = [
  { title: "Страница", path: "page", open: true },
  { title: "Основной текст", path: "body", open: true },
  {
    title: "Заголовки",
    children: defaultProfile.headings.map((_heading, index) => ({ title: `Заголовок ${index + 1}`, path: `headings.${index}` })),
  },
  { title: "Списки", path: "lists" },
  { title: "Цитаты", path: "quotes" },
  { title: "Код", path: "code" },
  { title: "Таблицы", path: "tables" },
  { title: "Изображения", path: "images" },
  { title: "Подписи", path: "captions" },
  {
    title: "Ссылки и разделители",
    children: [
      { title: "Ссылки", path: "links" },
      { title: "Горизонтальная линия", path: "horizontal_rule" },
    ],
  },
  { title: "Колонтитулы", path: "headers_footers" },
];

const profileFieldOrder = {
  page: ["width_mm", "height_mm", "left_margin_mm", "right_margin_mm", "top_margin_mm", "bottom_margin_mm", "header_distance_mm", "footer_distance_mm"],
  body: ["font_family", "font_size_pt", "color", "alignment", "line_spacing", "first_line_indent_mm", "space_before_pt", "space_after_pt"],
  "headings.0": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  "headings.1": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  "headings.2": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  "headings.3": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  "headings.4": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  "headings.5": ["font_family", "font_size_pt", "color", "bold", "italic", "alignment", "left_indent_mm", "first_line_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "keep_with_next", "keep_together"],
  lists: ["left_indent_mm", "depth_indent_mm", "hanging_indent_mm", "line_spacing", "space_after_pt"],
  quotes: ["left_indent_mm", "depth_indent_mm", "right_indent_mm", "line_spacing", "space_before_pt", "space_after_pt", "border_color", "border_width_pt"],
  code: ["font_family", "font_size_pt", "color", "background_color", "left_indent_mm", "line_spacing", "space_before_pt", "space_after_pt"],
  tables: ["alignment", "header_font_family", "header_font_size_pt", "header_color", "header_bold", "header_alignment", "body_font_family", "body_font_size_pt", "body_color", "body_alignment", "border_color", "border_width_pt", "vertical_alignment", "padding_top_mm", "padding_horizontal_mm", "padding_bottom_mm"],
  images: ["alignment", "max_width_percent"],
  captions: ["table", "figure", "continuation"],
  "captions.table": ["font_family", "font_size_pt", "color", "bold", "alignment", "space_before_pt", "space_after_pt"],
  "captions.figure": ["font_family", "font_size_pt", "color", "bold", "alignment", "space_before_pt", "space_after_pt"],
  "captions.continuation": ["font_family", "font_size_pt", "color", "bold", "alignment", "space_before_pt", "space_after_pt"],
  links: ["font_family", "font_size_pt", "color", "underline"],
  horizontal_rule: ["color", "width_pt", "space_before_pt", "space_after_pt"],
  headers_footers: ["font_family", "font_size_pt", "color", "text_alignment", "page_alignment"],
};

function getAtPath(object, path) {
  return path.split(".").reduce((value, key) => value[key], object);
}

function setAtPath(object, path, value) {
  const keys = path.split(".");
  const last = keys.pop();
  const parent = keys.reduce((target, key) => target[key], object);
  parent[last] = value;
}

function profileControlType(path, value) {
  if (typeof value === "boolean") return "checkbox";
  if (path.endsWith("color")) return "color";
  if (path === "tables.alignment") return "table-alignment";
  if (path.endsWith("alignment")) return path.endsWith("vertical_alignment") ? "vertical-alignment" : "alignment";
  if (path.endsWith("font_family")) return "font-family";
  if (typeof value === "number") return "number";
  return "text";
}

function createProfileControl(path, value) {
  const wrapper = document.createElement("div");
  const type = profileControlType(path, value);
  const id = `profile-${path.replace(/[^a-z0-9]+/gi, "-")}`;
  wrapper.className = `profile-control${type === "checkbox" ? " profile-control--checkbox" : ""}${type === "color" ? " profile-control--color" : ""}`;

  const label = document.createElement("label");
  label.htmlFor = id;
  label.textContent = profileLabels[path.split(".").at(-1)] ?? path;

  let control;
  if (type === "alignment" || type === "vertical-alignment" || type === "table-alignment") {
    control = document.createElement("select");
    const options = type === "vertical-alignment"
      ? [["top", "Сверху"], ["center", "По центру"], ["bottom", "Снизу"]]
      : type === "table-alignment"
        ? [["left", "Слева"], ["center", "По центру"], ["right", "Справа"]]
        : [["left", "Слева"], ["center", "По центру"], ["right", "Справа"], ["justify", "По ширине"]];
    for (const [optionValue, optionLabel] of options) {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionLabel;
      control.append(option);
    }
  } else {
    control = document.createElement("input");
    control.type = type === "font-family" ? "text" : type;
    if (type === "font-family") {
      control.setAttribute("list", "font-family-options");
      control.maxLength = 80;
      wrapper.classList.add("profile-control--wide");
    }
    if (type === "number") {
      control.step = "any";
      if (path === "page.width_mm" || path === "page.height_mm") {
        control.min = "100";
        control.max = "420";
      } else if (path.endsWith("line_spacing")) {
        control.min = "1";
        control.max = "3";
      } else if (path.endsWith("font_size_pt")) {
        control.min = "6";
        control.max = "72";
      } else if (path.endsWith("max_width_percent")) {
        control.min = "10";
        control.max = "100";
      } else if (path.endsWith("border_width_pt")) {
        control.min = "0.1";
        control.max = "3";
      } else if (path.includes("padding") && path.endsWith("_mm")) {
        control.min = "0";
        control.max = "10";
      } else if (path.endsWith("_mm")) {
        control.min = "0";
        control.max = "100";
      } else {
        control.min = "0";
        control.max = "72";
      }
    }
  }

  control.id = id;
  control.dataset.profilePath = path;
  control.required = type !== "checkbox";
  if (type === "checkbox") control.checked = value;
  else control.value = String(value);
  if (type !== "checkbox") wrapper.append(label, control);
  else wrapper.append(control, label);
  return wrapper;
}

function appendProfileFields(container, path) {
  const object = getAtPath(defaultProfile, path);
  const fields = document.createElement("div");
  fields.className = "profile-fields";
  const orderedKeys = profileFieldOrder[path] ?? Object.keys(object);
  const keys = [...orderedKeys.filter((key) => key in object), ...Object.keys(object).filter((key) => !orderedKeys.includes(key))];
  for (const key of keys) {
    const value = object[key];
    const fieldPath = `${path}.${key}`;
    if (value && typeof value === "object" && !Array.isArray(value)) {
      const subgroup = createDetails(key === "table" ? "Подпись таблицы" : key === "figure" ? "Подпись рисунка" : key === "continuation" ? "Продолжение таблицы" : key, false, "profile-subgroup");
      appendProfileFields(subgroup.content, fieldPath);
      fields.append(subgroup.element);
    } else {
      fields.append(createProfileControl(fieldPath, value));
    }
  }
  container.append(fields);
}

function createDetails(title, open = false, className = "profile-group") {
  const element = document.createElement("details");
  element.className = className;
  element.open = open;
  const summary = document.createElement("summary");
  summary.textContent = title;
  const content = document.createElement("div");
  element.append(summary, content);
  return { element, content };
}

function renderProfile() {
  profileControls.replaceChildren();
  for (const group of profileGroups) {
    const details = createDetails(group.title, Boolean(group.open));
    if (group.path) appendProfileFields(details.content, group.path);
    for (const child of group.children ?? []) {
      const nested = createDetails(child.title, false, "profile-subgroup");
      appendProfileFields(nested.content, child.path);
      details.content.append(nested.element);
    }
    profileControls.append(details.element);
  }
  const familyList = document.createElement("datalist");
  familyList.id = "font-family-options";
  for (const name of ["Times New Roman", "Arial", "Calibri", "Georgia", "Cambria", "Verdana", "Consolas", "Courier New"]) {
    const option = document.createElement("option");
    option.value = name;
    familyList.append(option);
  }
  profileControls.append(familyList);
}

function readProfile() {
  const profile = JSON.parse(JSON.stringify(defaultProfile));
  for (const control of profileControls.querySelectorAll("[data-profile-path]")) {
    let value;
    if (control.type === "checkbox") value = control.checked;
    else if (control.type === "number") value = Number(control.value);
    else value = control.value;
    setAtPath(profile, control.dataset.profilePath, value);
  }
  return profile;
}

function setFeedback(message, kind = "") {
  feedback.textContent = message;
  if (kind) feedback.dataset.kind = kind;
  else delete feedback.dataset.kind;
}

function setFieldError(field, message) {
  const targets = {
    source: sourceError,
    assets: assetsError,
    headers: document.querySelector("#headers-error"),
    profile: profileError,
  };
  const target = targets[field];
  if (!target) {
    setFeedback(message, "error");
    return;
  }
  target.textContent = message;
  target.hidden = false;
  if (field === "source") sourceDropzone.setAttribute("aria-invalid", "true");
  if (field === "assets") assetsDropzone.setAttribute("aria-invalid", "true");
  if (field === "profile") {
    const invalidPaths = [...message.matchAll(/«profile\.([^»]+)»/g)].map((match) => match[1]);
    for (const invalidPath of invalidPaths) {
      const control = [...profileControls.querySelectorAll("[data-profile-path]")].find((item) => item.dataset.profilePath === invalidPath);
      if (!control) continue;
      control.setAttribute("aria-invalid", "true");
      let parent = control.closest("details");
      while (parent) {
        parent.open = true;
        parent = parent.parentElement.closest("details");
      }
    }
  }
  setFeedback(message, "error");
}

function clearFieldError(field) {
  const targets = {
    source: sourceError,
    assets: assetsError,
    headers: document.querySelector("#headers-error"),
    profile: profileError,
  };
  if (targets[field]) {
    targets[field].textContent = "";
    targets[field].hidden = true;
  }
  if (field === "source") sourceDropzone.removeAttribute("aria-invalid");
  if (field === "assets") assetsDropzone.removeAttribute("aria-invalid");
  if (field === "profile") profileControls.querySelectorAll("[aria-invalid='true']").forEach((control) => control.removeAttribute("aria-invalid"));
  if (field === "headers") {
    document.querySelectorAll("#header-text, #footer-text, #numbering-start-page").forEach((control) => control.removeAttribute("aria-invalid"));
  }
  if (![sourceError, assetsError, document.querySelector("#headers-error"), profileError].some((node) => !node.hidden)) {
    setFeedback("");
  }
}

function clearAllErrors() {
  for (const field of ["source", "assets", "headers", "profile"]) clearFieldError(field);
  setFeedback("");
}

function clearResolution() {
  resolution.replaceChildren();
  resolution.hidden = true;
  markdownChoice.hidden = true;
  mainPath.replaceChildren();
}

function appendMarkdownChoices(paths) {
  markdownChoice.hidden = false;
  mainPath.replaceChildren();
  for (const path of paths) {
    const option = document.createElement("option");
    option.value = path;
    option.textContent = path;
    mainPath.append(option);
  }
}

function appendResolutionControls(payload) {
  resolution.replaceChildren();
  resolution.hidden = false;
  const heading = document.createElement("h3");
  heading.textContent = "Выберите подходящие файлы";
  resolution.append(heading);
  for (const item of payload.ambiguous_images ?? []) {
    const label = document.createElement("label");
    const select = document.createElement("select");
    select.dataset.assetSource = item.src;
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = `Выберите файл для ${item.src}`;
    placeholder.disabled = true;
    placeholder.selected = true;
    select.append(placeholder);
    for (const match of item.matches) {
      const option = document.createElement("option");
      option.value = match.id;
      option.textContent = match.path;
      select.append(option);
    }
    label.append(document.createTextNode(`Ссылка ${item.src}`), select);
    resolution.append(label);
  }
}

function collectChoices() {
  const choices = {};
  for (const select of resolution.querySelectorAll("select[data-asset-source]")) {
    if (select.value) choices[select.dataset.assetSource] = select.value;
  }
  return choices;
}

function filenameFromResponse(response) {
  const disposition = response.headers.get("Content-Disposition") ?? "";
  const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encoded) return decodeURIComponent(encoded[1]);
  const plain = disposition.match(/filename="?([^";]+)"?/i);
  return plain?.[1] ?? "document-formatted.docx";
}

function fileExtension(name) {
  const basename = name.replaceAll("\\", "/").split("/").at(-1);
  const dot = basename.lastIndexOf(".");
  return dot < 0 ? "" : basename.slice(dot).toLowerCase();
}

function invalidateResult(message = "") {
  const hadResult = !result.hidden;
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
  result.hidden = true;
  downloadLink.removeAttribute("href");
  if (hadResult && message) setFeedback(message);
}

function handleSourceFile(file) {
  if (!file) return;
  const extension = fileExtension(file.name);
  if (![".md", ".zip"].includes(extension)) {
    setFieldError("source", `${file.name}: в этой области нужен файл .md или ZIP-проект.`);
    sourceInput.value = "";
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    setFieldError("source", `${file.name}: размер превышает общий лимит 50 МБ.`);
    sourceInput.value = "";
    return;
  }
  clearFieldError("source");
  clearFieldError("assets");
  invalidateResult("Исходный файл изменён. Преобразуйте документ ещё раз.");
  clearResolution();
  selectedSource = file;
  sourceFilename.textContent = file.name;
  sourceName.hidden = false;
  if (extension === ".md" && typeof imagesDialog.showModal === "function") imagesDialog.showModal();
}

sourceInput.addEventListener("change", () => {
  const [file] = sourceInput.files;
  sourceInput.value = "";
  handleSourceFile(file);
});

function setDropState(element, value) {
  element.classList.toggle("is-dragging", value);
}

for (const eventName of ["dragenter", "dragover"]) {
  sourceDropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    setDropState(sourceDropzone, true);
  });
  assetsDropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    setDropState(assetsDropzone, true);
  });
}

for (const eventName of ["dragleave", "drop"]) {
  sourceDropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    setDropState(sourceDropzone, false);
  });
  assetsDropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    setDropState(assetsDropzone, false);
  });
}

sourceDropzone.addEventListener("drop", (event) => {
  const files = [...(event.dataTransfer?.files ?? [])];
  if (files.length !== 1) {
    setFieldError("source", "Перетащите сюда один файл .md или ZIP-проект.");
    return;
  }
  handleSourceFile(files[0]);
});

document.querySelector("#images-no").addEventListener("click", () => imagesDialog.close());
document.querySelector("#images-yes").addEventListener("click", () => {
  imagesDialog.close();
  assetFolderInput.click();
});

document.querySelector("#choose-assets").addEventListener("click", () => assetFilesInput.click());
document.querySelector("#choose-folder").addEventListener("click", () => assetFolderInput.click());
document.querySelector("#choose-assets-zip").addEventListener("click", () => assetZipInput.click());

function selectedFolderPath(file) {
  const path = file.webkitRelativePath || file.name;
  const firstSlash = path.indexOf("/");
  return firstSlash < 0 ? file.name : path.slice(firstSlash + 1);
}

function addAssetFiles(entries) {
  if (!entries.length) return;
  clearFieldError("assets");
  const rejected = [];
  const added = [];
  let totalBytes = selectedAssets.reduce((sum, asset) => sum + asset.file.size, 0);

  for (const entry of entries) {
    const file = entry.file;
    const path = entry.path || file.name;
    const extension = fileExtension(path);
    if (extension === ".zip") {
      if (file.size > 50 * 1024 * 1024) {
        rejected.push(`${path}: ZIP больше 50 МБ`);
        continue;
      }
    } else if (!validImageExtensions.has(extension)) {
      rejected.push(`${path}: нужен файл изображения или ZIP`);
      continue;
    } else if (file.size > 10 * 1024 * 1024) {
      rejected.push(`${path}: изображение больше 10 МБ`);
      continue;
    }
    if (totalBytes + file.size > 50 * 1024 * 1024) {
      rejected.push(`${path}: общий размер загрузки превысит 50 МБ`);
      continue;
    }
    totalBytes += file.size;
    added.push({ id: crypto.randomUUID?.() ?? `${Date.now()}-${Math.random()}`, file, path });
  }

  if (added.length) {
    selectedAssets.push(...added);
    invalidateResult();
    clearResolution();
    if (!rejected.length) setFeedback("Добавлено файлов: " + added.length + ".", "success");
  }
  renderAssetList();
  if (rejected.length) setFieldError("assets", `Не добавлены файлы: ${rejected.join("; ")}.`);
}

function renderAssetList() {
  assetList.replaceChildren();
  for (const asset of selectedAssets) {
    const item = document.createElement("li");
    item.className = "asset-item";
    const info = document.createElement("span");
    info.className = "asset-info";
    const name = document.createElement("span");
    name.className = "asset-name";
    name.textContent = asset.path;
    const meta = document.createElement("span");
    meta.className = "asset-meta";
    meta.textContent = `${(asset.file.size / 1024).toFixed(0)} КБ`;
    info.append(name, meta);
    const remove = document.createElement("button");
    remove.className = "remove-asset";
    remove.type = "button";
    remove.textContent = "Удалить";
    remove.setAttribute("aria-label", `Удалить файл ${asset.path}`);
    remove.addEventListener("click", () => {
      selectedAssets = selectedAssets.filter((item) => item.id !== asset.id);
      renderAssetList();
      clearFieldError("assets");
      invalidateResult("Список ресурсов изменён. Преобразуйте документ ещё раз.");
      clearResolution();
    });
    item.append(info, remove);
    assetList.append(item);
  }
}

assetFilesInput.addEventListener("change", () => {
  addAssetFiles([...assetFilesInput.files].map((file) => ({ file, path: file.name })));
  assetFilesInput.value = "";
});
assetFolderInput.addEventListener("change", () => {
  addAssetFiles([...assetFolderInput.files].map((file) => ({ file, path: selectedFolderPath(file) })));
  assetFolderInput.value = "";
});
assetZipInput.addEventListener("change", () => {
  addAssetFiles([...assetZipInput.files].map((file) => ({ file, path: file.name })));
  assetZipInput.value = "";
});

function readDirectoryEntries(entry) {
  return new Promise((resolve, reject) => {
    const reader = entry.createReader();
    const entries = [];
    const readBatch = () => reader.readEntries((batch) => {
      if (!batch.length) resolve(entries);
      else {
        entries.push(...batch);
        readBatch();
      }
    }, reject);
    readBatch();
  });
}

function readEntryFile(entry) {
  return new Promise((resolve, reject) => entry.file(resolve, reject));
}

async function walkDroppedEntry(entry, prefix = "", isRoot = false) {
  if (entry.isFile) {
    const file = await readEntryFile(entry);
    return [{ file, path: prefix ? `${prefix}/${entry.name}` : entry.name }];
  }
  if (!entry.isDirectory) return [];
  const children = await readDirectoryEntries(entry);
  const nextPrefix = isRoot ? prefix : (prefix ? `${prefix}/${entry.name}` : entry.name);
  const nested = await Promise.all(children.map((child) => walkDroppedEntry(child, nextPrefix)));
  return nested.flat();
}

assetsDropzone.addEventListener("drop", async (event) => {
  const items = [...(event.dataTransfer?.items ?? [])];
  try {
    const entries = items.map((item) => item.webkitGetAsEntry?.()).filter(Boolean);
    if (entries.length) {
      const files = (await Promise.all(entries.map((entry) => walkDroppedEntry(entry, "", entry.isDirectory)))).flat();
      addAssetFiles(files);
    } else {
      const files = [...(event.dataTransfer?.files ?? [])].map((file) => ({ file, path: file.name }));
      addAssetFiles(files);
    }
  } catch {
    setFieldError("assets", "Не удалось прочитать перетащенную папку. Выберите её кнопкой «Выбрать папку».");
  }
});

function updateHeaderPreview() {
  for (const name of ["header", "footer"]) {
    const mode = document.querySelector(`#${name}-mode`).value;
    const text = document.querySelector(`#${name}-text`);
    const preview = document.querySelector(`#${name}-preview`);
    text.hidden = mode !== "text";
    text.required = mode === "text";
    const value = mode === "none" ? "Пусто" : mode === "page" ? "№ " + (numberingMode.value === "restart" ? "1" : numberingStartPage.value || "1") : text.value.trim() || "Текст колонтитула";
    preview.textContent = value;
  }
  const hasPageNumber = ["header-mode", "footer-mode"].some((id) => document.querySelector(`#${id}`).value === "page");
  numberingControls.setAttribute("aria-disabled", String(!hasPageNumber));
  for (const control of numberingControls.querySelectorAll("input, select")) control.disabled = !hasPageNumber;
}

document.querySelectorAll("#header-mode, #footer-mode, #header-text, #footer-text").forEach((control) => {
  control.addEventListener("input", () => {
    updateHeaderPreview();
    clearFieldError("headers");
  });
  control.addEventListener("change", () => {
    updateHeaderPreview();
    clearFieldError("headers");
  });
});
numberingStartPage.addEventListener("input", updateHeaderPreview);
numberingMode.addEventListener("change", updateHeaderPreview);

document.querySelector("#preview-header-target").addEventListener("click", () => document.querySelector("#header-mode").focus());
document.querySelector("#preview-footer-target").addEventListener("click", () => document.querySelector("#footer-mode").focus());
mainPath.addEventListener("change", () => clearFieldError("source"));
resolution.addEventListener("change", () => clearFieldError("assets"));

document.querySelector("#reset-profile").addEventListener("click", () => {
  renderProfile();
  clearFieldError("profile");
  invalidateResult();
  setFeedback("Исходные параметры оформления восстановлены.", "success");
});

function firstInvalidProfileControl() {
  for (const control of profileControls.querySelectorAll("input, select")) {
    if (!control.checkValidity()) {
      let parent = control.closest("details");
      while (parent) {
        parent.open = true;
        parent = parent.parentElement.closest("details");
      }
      return control;
    }
  }
  return null;
}

function buildFormData(profile) {
  const data = new FormData();
  data.append("source", selectedSource, selectedSource.name);
  for (const asset of selectedAssets) data.append("assets", asset.file, asset.path);
  data.append("main_path", mainPath.value);
  data.append("asset_choices", JSON.stringify(collectChoices()));
  data.append("profile", JSON.stringify(profile));
  data.append("header_mode", document.querySelector("#header-mode").value);
  data.append("header_text", document.querySelector("#header-text").value);
  data.append("footer_mode", document.querySelector("#footer-mode").value);
  data.append("footer_text", document.querySelector("#footer-text").value);
  data.append("numbering_start_page", numberingStartPage.value);
  data.append("numbering_mode", numberingMode.value);
  return data;
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (isConverting) return;
  clearAllErrors();
  clearFieldError("profile");

  if (!selectedSource) {
    setFieldError("source", "Сначала выберите Markdown-файл или ZIP-проект.");
    sourceInput.focus();
    return;
  }

  const invalidProfile = firstInvalidProfileControl();
  if (invalidProfile) {
    setFieldError("profile", "Проверьте выделенные параметры оформления.");
    invalidProfile.setAttribute("aria-invalid", "true");
    invalidProfile.focus();
    return;
  }

  const headerMode = document.querySelector("#header-mode").value;
  const footerMode = document.querySelector("#footer-mode").value;
  if ((headerMode === "text" && !document.querySelector("#header-text").value.trim()) || (footerMode === "text" && !document.querySelector("#footer-text").value.trim())) {
    setFieldError("headers", "Введите текст выбранного колонтитула или измените его режим.");
    document.querySelector(headerMode === "text" ? "#header-text" : "#footer-text").focus();
    return;
  }
  if (!numberingStartPage.checkValidity()) {
    setFieldError("headers", "Укажите целый номер листа от 1 до 1000.");
    numberingStartPage.setAttribute("aria-invalid", "true");
    numberingStartPage.focus();
    return;
  }

  isConverting = true;
  button.disabled = true;
  button.textContent = "Подготовка документа…";
  setFeedback("Проверяем вложения и собираем DOCX.");
  invalidateResult();

  try {
    const response = await fetch("/api/convert", { method: "POST", body: buildFormData(readProfile()) });
    const contentType = response.headers.get("Content-Type") ?? "";
    if (!response.ok || contentType.includes("application/json")) {
      const payload = await response.json();
      if (payload.markdown_files) appendMarkdownChoices(payload.markdown_files);
      if (payload.ambiguous_images) appendResolutionControls(payload);
      setFieldError(payload.field ?? "source", payload.error ?? "Не удалось создать документ.");
      return;
    }

    const blob = await response.blob();
    downloadUrl = URL.createObjectURL(blob);
    downloadLink.href = downloadUrl;
    downloadLink.download = filenameFromResponse(response);
    resultFilename.textContent = downloadLink.download;
    result.hidden = false;
    downloadLink.click();
    setFeedback("DOCX готов. Файл скачан; его можно скачать ещё раз или начать новый документ.", "success");
    result.scrollIntoView({ behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "nearest" });
  } catch {
    setFeedback("Не удалось связаться с сервером. Проверьте соединение и повторите попытку.", "error");
  } finally {
    isConverting = false;
    button.disabled = false;
    button.textContent = "Преобразовать в DOCX";
  }
});

function resetForNewFile() {
  if (imagesDialog.open) imagesDialog.close();
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
  selectedSource = null;
  selectedAssets = [];
  form.reset();
  sourceInput.value = "";
  assetFilesInput.value = "";
  assetFolderInput.value = "";
  assetZipInput.value = "";
  sourceFilename.textContent = "";
  sourceName.hidden = true;
  renderAssetList();
  clearResolution();
  clearAllErrors();
  result.hidden = true;
  downloadLink.removeAttribute("href");
  numberingStartPage.value = "1";
  numberingMode.value = "continue";
  renderProfile();
  updateHeaderPreview();
  window.scrollTo({ top: 0, behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
  sourceInput.focus({ preventScroll: true });
}

document.querySelector("#new-file-button").addEventListener("click", resetForNewFile);

function handleFormEdit(event) {
  if (event.target.matches("[data-profile-path]")) {
    event.target.removeAttribute("aria-invalid");
    clearFieldError("profile");
  }
  if (event.target.matches("#numbering-start-page, #numbering-mode")) clearFieldError("headers");
  if (!result.hidden) {
    invalidateResult();
    setFeedback("Настройки изменены. Преобразуйте документ ещё раз.");
  }
}

form.addEventListener("input", handleFormEdit);
form.addEventListener("change", handleFormEdit);

renderProfile();
renderAssetList();
updateHeaderPreview();
