/* global Office, Excel, Word */

// Variable globale pour stocker l'image
let capturedImageBase64 = null;

// Quand Office est prêt
Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("captureBtn").onclick = captureRange;
    document.getElementById("copyBase64Btn").onclick = copyImageToClipboard;
    document.getElementById("insertWordBtn").onclick = insertInWord;
  }
});

// CAPTURE DE LA PLAGE
async function captureRange() {
  const btn = document.getElementById("captureBtn");
  btn.disabled = true;
  btn.textContent = "⏳ Capture en cours...";
  capturedImageBase64 = null;

  // Cacher les anciens résultats
  document.getElementById("previewBox").style.display = "none";
  document.getElementById("base64Box").style.display = "none";
  document.getElementById("copyBase64Btn").style.display = "none";
  document.getElementById("insertWordBtn").style.display = "none";
  document.getElementById("wordInstructions").style.display = "none";

  setStatus("info", "⏳ Lecture de la sélection...");

  try {
    await Excel.run(async function(context) {
      const range = context.workbook.getSelectedRange();
      range.load("address, cellCount");
      await context.sync();

      if (range.cellCount === 0) {
        setStatus("error", "❌ Aucune cellule sélectionnée.");
        btn.disabled = false;
        btn.textContent = "📷 Capturer la sélection";
        return;
      }

      setStatus("info", "📍 Plage : " + range.address + " — Génération...");

      // Capturer la plage en image PNG (base64)
      const imageResult = range.getImage();
      await context.sync();

      const base64 = imageResult.value;

      if (!base64 || base64.length === 0) {
        setStatus("error", "❌ Impossible de capturer. Essaie avec des cellules non vides.");
        btn.disabled = false;
        btn.textContent = "📷 Capturer la sélection";
        return;
      }

      capturedImageBase64 = base64;
      const dataUrl = "data:image/png;base64," + base64;

      // Afficher l'aperçu
      document.getElementById("preview").src = dataUrl;
      document.getElementById("previewBox").style.display = "block";

      // Afficher le code base64
      document.getElementById("base64Output").value = dataUrl;
      document.getElementById("base64Box").style.display = "block";

      // Afficher les boutons
      document.getElementById("copyBase64Btn").style.display = "block";
      document.getElementById("insertWordBtn").style.display = "block";
      document.getElementById("wordInstructions").style.display = "block";

      setStatus("success", "✅ Image capturée ! Clique sur Copier puis Ctrl+V dans Word.");
    });

  } catch (error) {
    setStatus("error", "❌ Erreur : " + error.message);
    console.error(error);
  }

  btn.disabled = false;
  btn.textContent = "📷 Capturer la sélection";
}

// COPIER L'IMAGE
async function copyImageToClipboard() {
  if (!capturedImageBase64) {
    setStatus("error", "❌ Aucune image. Clique d'abord sur Capturer.");
    return;
  }
  try {
    const dataUrl = "data:image/png;base64," + capturedImageBase64;
    const response = await fetch(dataUrl);
    const blob = await response.blob();
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    setStatus("success", "📋 Image copiée ! Fais Ctrl+V dans Word.");
  } catch (e) {
    // Fallback : copier le texte base64
    try {
      const dataUrl = "data:image/png;base64," + capturedImageBase64;
      await navigator.clipboard.writeText(dataUrl);
      setStatus("info", "📋 Code copié. Dans Word : Insertion > Image > URL.");
    } catch (e2) {
      document.getElementById("base64Output").select();
      document.execCommand("copy");
      setStatus("info", "📋 Code copié depuis la zone de texte !");
    }
  }
}

// INSÉRER DANS WORD
async function insertInWord() {
  if (!capturedImageBase64) {
    setStatus("error", "❌ Aucune image capturée.");
    return;
  }
  setStatus("info", "ℹ️ Utilise le bouton Copier puis Ctrl+V dans Word.");
}

// AFFICHER UN MESSAGE DE STATUT
function setStatus(type, message) {
  const el = document.getElementById("status");
  el.className = type;
  el.textContent = message;
  el.style.display = "block";
}