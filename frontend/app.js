const form = document.querySelector("#sopForm");
const topic = document.querySelector("#topic");
const model = document.querySelector("#model");
const generateBtn = document.querySelector("#generateBtn");
const saveDraftBtn = document.querySelector("#saveDraftBtn");
const logs = document.querySelector("#logs");
const clearLogs = document.querySelector("#clearLogs");
const preview = document.querySelector("#preview");
const emptyPreview = document.querySelector("#emptyPreview");
const docTitle = document.querySelector("#docTitle");
const summary = document.querySelector("#summary");
const kbNumber = document.querySelector("#kbNumber");
const kbStatus = document.querySelector("#kbStatus");
const layout = document.querySelector("#layout");
const downloadActions = document.querySelector("#downloadActions");
const downloadDocx = document.querySelector("#downloadDocx");
const downloadImage = document.querySelector("#downloadImage");
let currentDraftId = "";
let currentTopic = "";

function addLog(message, level = "info") {
  const item = document.createElement("li");
  item.className = level;
  item.innerHTML = `<span class="dot"></span><span>${escapeHtml(message)}</span>`;
  logs.appendChild(item);
  logs.scrollTop = logs.scrollHeight;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setLoading(isLoading) {
  generateBtn.disabled = isLoading;
  generateBtn.textContent = isLoading ? "Generating..." : "Generate SOP";
  emptyPreview.classList.toggle("is-loading", isLoading);
}

function setSaving(isSaving) {
  saveDraftBtn.disabled = isSaving;
  saveDraftBtn.textContent = isSaving ? "Saving..." : "Save as Draft";
}

function downloadUrl(path) {
  return `/download?file=${encodeURIComponent(path)}`;
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const request = {
    topic: topic.value.trim(),
    model: model.value.trim(),
  };

  if (!request.topic) {
    addLog("Please enter the problem or SOP topic.", "error");
    return;
  }

  logs.innerHTML = "";
  preview.removeAttribute("srcdoc");
  emptyPreview.classList.remove("hidden");
  summary.classList.add("hidden");
  downloadActions.classList.add("hidden");
  saveDraftBtn.classList.add("hidden");
  currentDraftId = "";
  currentTopic = "";
  docTitle.textContent = "Generating SOP draft";

  addLog(`Request accepted: ${request.topic}`);
  addLog("Ollama is building the SOP, DOCX, and ServiceNow-ready KB HTML.");
  setLoading(true);

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ topic: request.topic, model: request.model }),
    });

    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || "Generation failed.");
    }

    result.logs.forEach((entry) => addLog(entry.message, entry.level));
    currentDraftId = result.draft_id;
    currentTopic = result.topic || request.topic;

    docTitle.textContent = result.title || request.topic;
    preview.srcdoc = result.html;
    emptyPreview.classList.add("hidden");

    kbNumber.textContent = "Not saved";
    kbStatus.textContent = "Generated";
    layout.textContent = result.layout || "-";
    summary.classList.remove("hidden");

    downloadDocx.href = downloadUrl(result.paths.docx);
    downloadImage.href = downloadUrl(result.paths.image);
    downloadActions.classList.remove("hidden");
    saveDraftBtn.classList.remove("hidden");
    addLog("Review the generated SOP, then click Save as Draft to publish HTML to ServiceNow.", "info");
  } catch (error) {
    addLog(error.message, "error");
    if (!preview.srcdoc) {
      docTitle.textContent = "Generation failed";
    } else {
      kbNumber.textContent = "Publish failed";
      kbStatus.textContent = "Generated, not published";
    }
  } finally {
    setLoading(false);
  }
});

saveDraftBtn.addEventListener("click", async () => {
  if (!currentDraftId) {
    addLog("Generate an SOP before saving a ServiceNow draft.", "error");
    return;
  }

  setSaving(true);
  kbNumber.textContent = "Saving...";
  kbStatus.textContent = "Publishing";
  addLog("Saving generated SOP HTML as a ServiceNow KB draft...");

  try {
    const publishResponse = await fetch("/api/publish", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        draft_id: currentDraftId,
        topic: currentTopic || docTitle.textContent,
        html: preview.srcdoc,
      }),
    });
    const publishResult = await publishResponse.json();
    if (!publishResponse.ok || !publishResult.ok) {
      throw new Error(publishResult.error || "ServiceNow publish failed.");
    }

    publishResult.logs.forEach((entry) => addLog(entry.message, entry.level));
    const published = publishResult.service_now || {};
    kbNumber.textContent = published.number || "N/A";
    kbStatus.textContent = published.workflow_state || "draft";
  } catch (error) {
    addLog(error.message, "error");
    kbNumber.textContent = "Save failed";
    kbStatus.textContent = "Generated, not saved";
  } finally {
    setSaving(false);
  }
});

clearLogs.addEventListener("click", () => {
  logs.innerHTML = "";
});
