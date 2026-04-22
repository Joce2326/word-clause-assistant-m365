require("dotenv").config();

const express = require("express");
const cors = require("cors");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
const port = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
};

const cca = new ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
  const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"]
  };

  const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);

  if (!response || !response.accessToken) {
    throw new Error("Could not acquire Microsoft Graph access token.");
  }

  return response.accessToken;
}

async function getGraphJson(url, accessToken) {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph request failed: ${response.status} ${response.statusText} - ${errorText}`);
  }

  return response.json();
}

app.get("/", (req, res) => {
  res.json({
    message: "Clause API running with Microsoft Graph"
  });
});

app.get("/api/clauses", async (req, res) => {
  try {
    const accessToken = await getAccessToken();

    const hostname = process.env.SITE_HOSTNAME;
    const sitePath = process.env.SITE_PATH || "/";
    const listName = process.env.LIST_NAME || "Clause Library";

    // Step 1: get site ID
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`;
    const siteData = await getGraphJson(siteUrl, accessToken);

    const siteId = siteData.id;

    if (!siteId) {
      throw new Error("Could not resolve SharePoint site ID from Microsoft Graph.");
    }

    // Step 2: get lists in site
    const listsUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`;
    const listsData = await getGraphJson(listsUrl, accessToken);

    const targetList = (listsData.value || []).find(
      (list) => list.displayName === listName
    );

    if (!targetList) {
      throw new Error(`List '${listName}' not found in site.`);
    }

    const listId = targetList.id;

    // Step 3: get list items with fields
    const itemsUrl =
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items` +
      `?expand=fields`;

    const itemsData = await getGraphJson(itemsUrl, accessToken);

    const clauses = (itemsData.value || [])
      .map((item) => {
        const fields = item.fields || {};

        return {
          id: item.id ? Number(item.id) : 0,
          title: fields.Title || "",
          category: fields.Category || "",
          clauseText: fields.ClauseText || "",
          versionLabel: fields.VersionClause ? String(fields.VersionClause) : "",
          approved: fields.Approved === true
        };
      })
      .filter((item) => item.approved);

    res.json(clauses);
  } catch (error) {
    console.error("Error in /api/clauses:", error);

    res.status(500).json({
      error: error instanceof Error ? error.message : "Unknown server error"
    });
  }
});

app.listen(port, () => {
  console.log(`Clause API running on http://localhost:${port}`);
});