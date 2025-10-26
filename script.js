document.addEventListener("DOMContentLoaded", () => {
  const jalonsList = document.getElementById("jalonsList");
  const messagesTableBody = document.querySelector("#messagesTable tbody");
  const rdvList = document.getElementById("rdvList");
  const autresList = document.getElementById("autresList");
  const livrablesList = document.getElementById("livrablesList");
  const uploadJson = document.getElementById("uploadJson");
  const loadBtn = document.getElementById("loadBtn");
  const uploadStatus = document.getElementById("uploadStatus");
  const generateMailBtn = document.getElementById("generateMailBtn");
  const mailPromptSelect = document.getElementById("mailPromptSelect");

  const mailPrompts = {
    1: "Écris un email professionnel clair et concis pour :",
    2: "Écris un email amical et léger pour :"
  };

  let llmData = null;

  function renderModules() {
    if (!llmData) return;

    // --- Jalons ---
    jalonsList.innerHTML = "";
    (llmData.jalons || []).forEach(j => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${j.titre}</strong> (${j.datePrévue})`;
      if (j.sousActions?.length) {
        const subUl = document.createElement("ul");
        j.sousActions.forEach(s => {
          const subLi = document.createElement("li");
          const cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = s.statut === "fait";
          cb.addEventListener("change", () => s.statut = cb.checked ? "fait" : "à faire");
          subLi.appendChild(cb);
          subLi.appendChild(document.createTextNode(s.texte));
          subUl.appendChild(subLi);
        });
        li.appendChild(subUl);
      }
      jalonsList.appendChild(li);
    });

    // --- Messages ---
    messagesTableBody.innerHTML = "";
    (llmData.messages || []).forEach(m => {
      const tr = document.createElement("tr");
      const tdCheck = document.createElement("td");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = m.envoyé;
      cb.addEventListener("change", () => m.envoyé = cb.checked);
      tdCheck.appendChild(cb);
      tr.appendChild(tdCheck);
      tr.appendChild(document.createElement("td")).textContent = m.destinataire;
      tr.appendChild(document.createElement("td")).textContent = m.sujet;
      tr.appendChild(document.createElement("td")).textContent = m.texte;
      messagesTableBody.appendChild(tr);
    });

    // --- RDV ---
    rdvList.innerHTML = "";
    (llmData.rdv || []).forEach(r => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${r.titre}</strong> - ${r.date} (${r.durée}) - Participants: ${r.participants.join(", ")}`;
      rdvList.appendChild(li);
    });

    // --- Autres ressources ---
    autresList.innerHTML = "";
    (llmData.autresModules || []).forEach(m => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${m.titre}</strong>`;
      if (m.items?.length) {
        const subUl = document.createElement("ul");
        m.items.forEach(it => {
          const subLi = document.createElement("li");
          const a = document.createElement("a");
          a.href = it.lien;
          a.textContent = it.nom;
          a.target = "_blank";
          subLi.appendChild(a);
          subUl.appendChild(subLi);
        });
        li.appendChild(subUl);
      }
      autresList.appendChild(li);
    });

    // --- Livrables ---
    livrablesList.innerHTML = "";
    (llmData.livrables || []).forEach(l => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${l.titre}</strong> (${l.type}) `;
      const btn = document.createElement("button");
      btn.textContent = "Télécharger Template";
      btn.addEventListener("click", () => generateTemplate(l));
      li.appendChild(btn);
      livrablesList.appendChild(li);
    });
  }

  // --- Charger JSON ---
  loadBtn.addEventListener("click", () => {
    const file = uploadJson.files[0];
    if (!file) { alert("Choisis un fichier JSON LLM !"); return; }
    const reader = new FileReader();
    reader.onload = e => {
      try {
        llmData = JSON.parse(e.target.result);
        renderModules();
        uploadStatus.textContent = `Fichier "${file.name}" chargé avec succès !`;
      } catch(err) {
        console.error(err);
        alert("Fichier JSON invalide !");
        uploadStatus.textContent = "";
      }
    };
    reader.readAsText(file);
  });

  // --- Générer Mail GPT ---
  generateMailBtn.addEventListener("click", () => {
    if (!llmData?.messages) return;
    const selectedMessages = llmData.messages.filter(m => m.envoyé);
    if (!selectedMessages.length) { alert("Coche au moins un message !"); return; }
    const promptId = mailPromptSelect.value;
    const promptTexte = mailPrompts[promptId];
    const content = selectedMessages.map(m => `À: ${m.destinataire}\nSujet: ${m.sujet}\nMessage: ${m.texte}`).join("\n\n");
    navigator.clipboard.writeText(`${promptTexte}\n\n${content}`)
      .then(() => alert("Prompt + messages copiés dans le presse-papiers !"))
      .catch(err => console.error("Erreur copie: ", err));
    window.open("https://chat.openai.com/", "_blank");
  });

  // --- Générer livrables ---
  function generateTemplate(l) {
    // --- DOCX ---
    if (l.type === "docx") {
      const doc = new docx.Document({
        sections: [{
          children: (l.template.plan || []).map(p =>
            new docx.Paragraph({
              children: [new docx.TextRun({ text: p, bold:true, size:24 })]
            })
          )
        }]
      });
      docx.Packer.toBlob(doc).then(blob => {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `${l.titre}.docx`;
        a.click();
      });

    // --- PPTX ---
    } else if (l.type === "pptx") {
      const pptx = new PptxGenJS();
      (l.template.slides || []).forEach(s => {
        const slide = pptx.addSlide();
        slide.addText(s, { x:1, y:1, fontSize:24, color:"363636" });
      });
      pptx.writeFile({ fileName: `${l.titre}.pptx` });

    // --- XLSX ---
    } else if (l.type === "xlsx") {
      const wb = XLSX.utils.book_new();
      (l.template.sheets || []).forEach(sheetName => {
        const ws = XLSX.utils.aoa_to_sheet([[""]]);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
      XLSX.writeFile(wb, `${l.titre}.xlsx`);
    }
  }
});
