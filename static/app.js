const form = document.querySelector("#convert-form");
const sourceInput = document.querySelector("#source");
const sourceName = document.querySelector("#source-name");
const feedback = document.querySelector("#feedback");
const button = document.querySelector("#convert-button");
const resolution = document.querySelector("#resolution");
const markdownChoice = document.querySelector("#markdown-choice");
const mainPath = document.querySelector("#main-path");

const textModes = [
  ["#header-mode", "#header-text"],
  ["#footer-mode", "#footer-text"],
];

for (const [selectSelector, inputSelector] of textModes) {
  const select = document.querySelector(selectSelector);
  const input = document.querySelector(inputSelector);
  const update = () => {
    const show = select.value === "text";
    input.hidden = !show;
    input.required = show;
  };
  select.addEventListener("change", update);
  update();
}

sourceInput.addEventListener("change", () => {
  sourceName.textContent = sourceInput.files[0]
    ? `Выбран файл: ${sourceInput.files[0].name}`
    : "";
  clearResolution();
});

function clearResolution() {
  resolution.replaceChildren();
  resolution.hidden = true;
  markdownChoice.hidden = true;
  mainPath.replaceChildren();
}

function setFeedback(message, kind = "") {
  feedback.textContent = message;
  if (kind) feedback.dataset.kind = kind;
  else delete feedback.dataset.kind;
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
  const heading = document.createElement("h2");
  heading.textContent = "Нужно уточнить файлы";
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
    let readableSource = item.src;
    try { readableSource = decodeURIComponent(item.src); } catch { /* Keep the source text readable as-is. */ }
    label.append(document.createTextNode(`Ссылка ${readableSource}`), select);
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

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!sourceInput.files.length) {
    setFeedback("Сначала выберите Markdown-файл или ZIP-архив проекта.", "error");
    sourceInput.focus();
    return;
  }
  if (!form.reportValidity()) return;

  const data = new FormData(form);
  data.set("asset_choices", JSON.stringify(collectChoices()));
  button.disabled = true;
  button.textContent = "Подготовка документа…";
  setFeedback("Проверяем файлы и собираем DOCX.");

  try {
    const response = await fetch("/api/convert", { method: "POST", body: data });
    const contentType = response.headers.get("Content-Type") ?? "";
    if (!response.ok || contentType.includes("application/json")) {
      const payload = await response.json();
      if (payload.markdown_files) appendMarkdownChoices(payload.markdown_files);
      if (payload.ambiguous_images) appendResolutionControls(payload);
      setFeedback(payload.error ?? "Не удалось создать документ.", "error");
      return;
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filenameFromResponse(response);
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    clearResolution();
    setFeedback("DOCX готов и скачан.", "success");
  } catch {
    setFeedback("Не удалось связаться с сервером. Проверьте подключение и повторите попытку.", "error");
  } finally {
    button.disabled = false;
    button.textContent = "Преобразовать и скачать DOCX";
  }
});
