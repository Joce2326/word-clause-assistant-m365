/// <reference types="office-js" />

import { insertTextIntoWord } from "./services/wordService";
import { getClausesFromSharePoint, IClauseItem } from "./services/sharePointService";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }

    if (appBody) {
      appBody.style.display = "block";
      appBody.innerHTML = "<p>Loading clauses from SharePoint...</p>";
      initializeApp(appBody);
    }
  }
});

async function initializeApp(container: HTMLElement): Promise<void> {
  try {
    const clauses = await getClausesFromSharePoint();
    const approvedClauses = clauses.filter((c) => c.approved);

    if (approvedClauses.length === 0) {
      container.innerHTML = "<p>No approved clauses found.</p>";
      return;
    }

    renderClauseAssistant(container, approvedClauses);
  } catch (error) {
    console.error("Error loading clauses:", error);

    container.innerHTML = `
    <div>
      <h2>Legal Clause Assistant</h2>
      <p style="color:red;">Failed to load clauses from SharePoint.</p>
      <p>${error instanceof Error ? error.message : "Unknown error"}</p>
    </div>
  `;
  }
}

function renderClauseAssistant(container: HTMLElement, clauses: IClauseItem[]): void {
  container.innerHTML = `
  <h2 class="ms-font-xl">Legal Clause Assistant</h2>
  <p class="ms-font-m">Select a clause and insert it into the document.</p>

  <label for="clauseSelect" class="ms-font-m">Choose Clause</label>
  <br /><br />

  <select id="clauseSelect" style="width:100%; padding:8px; margin-bottom:16px;">
    ${clauses
      .map(
        (c) =>
          `<option value="${c.id}">${c.title} (${c.category}${c.versionLabel ? ` - v${c.versionLabel}` : ""})</option>`
      )
      .join("")}
  </select>

  <br />

  <label for="clientName" class="ms-font-m">Client Name</label>
  <br />
  <input id="clientName" type="text" style="width:100%; padding:8px; margin-bottom:12px;" />

  <label for="effectiveDate" class="ms-font-m">Effective Date</label>
  <br />
  <input id="effectiveDate" type="text" style="width:100%; padding:8px; margin-bottom:12px;" />

  <label for="governingLaw" class="ms-font-m">Governing Law</label>
  <br />
  <input id="governingLaw" type="text" style="width:100%; padding:8px; margin-bottom:16px;" />

  <button id="insertBtn" style="padding:10px 16px; cursor:pointer;">
    Insert Clause
  </button>

  <p id="statusMessage" style="margin-top:16px;"></p>
`;

  const button = document.getElementById("insertBtn");
  const statusMessage = document.getElementById("statusMessage");

  if (button) {
    button.onclick = async () => {
      const select = document.getElementById("clauseSelect") as HTMLSelectElement;
      const selectedId = Number(select.value);
      const selectedClause = clauses.find((c) => c.id === selectedId);

      if (!selectedClause) {
        if (statusMessage) {
          statusMessage.textContent = "Clause not found.";
        }
        return;
      }

      try {
        let finalText = selectedClause.clauseText;

        const clientName = (document.getElementById("clientName") as HTMLInputElement).value;
        const effectiveDate = (document.getElementById("effectiveDate") as HTMLInputElement).value;
        const governingLaw = (document.getElementById("governingLaw") as HTMLInputElement).value;

        finalText = finalText
          .replace(/{{ClientName}}/g, clientName)
          .replace(/{{EffectiveDate}}/g, effectiveDate)
          .replace(/{{GoverningLaw}}/g, governingLaw);

        await insertTextIntoWord(finalText);

        if (statusMessage) {
          statusMessage.textContent = "Clause inserted successfully.";
        }
      } catch (error) {
        console.error("Error inserting clause:", error);

        if (statusMessage) {
          statusMessage.textContent = "Error inserting clause.";
        }
      }
    };
  }
}