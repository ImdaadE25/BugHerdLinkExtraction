function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Jira Export")
    .addItem("Export UAT Bugs", "showDatePickerForm")
    .addToUi();
}

// Fetch all Jira projects
function getJiraProjects() {
  const email = '';
   const apiToken = ''; // Set your Jira API token here
   const domain = 'eight25media.atlassian.net';


  const maxResults = 100;
  let startAt = 0;
  let allProjects = [];
  let total = 0;

  do {
    const url = `https://${domain}/rest/api/3/project/search?maxResults=${maxResults}&startAt=${startAt}`;
    const headers = {
      "Authorization": "Basic " + Utilities.base64Encode(email + ":" + apiToken),
      "Accept": "application/json"
    };

    const response = UrlFetchApp.fetch(url, { method: "get", headers });
    const data = JSON.parse(response.getContentText());

    allProjects = allProjects.concat(data.values);
    total = data.total;
    startAt += maxResults;
  } while (startAt < total);

  return allProjects.map(project => ({
    key: project.key,
    name: project.name
  }));
}


// Show UI form
function showDatePickerForm() {
  const projects = getJiraProjects();

  const projectOptions = projects.map(p =>
    `<label><input type="checkbox" name="project" value="${p.key}" /> ${p.name} (${p.key})</label>`
  ).join('<br>');

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        /* Reset & base */
        body {
          font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
          padding: 20px;
          background: #f9fafb;
          color: #222;
          margin: 0;
        }
        h3 {
          text-align: center;
          color: #007bff;
          margin-bottom: 20px;
          font-weight: 600;
        }
        label {
          display: block;
          margin: 5px 0;
          cursor: pointer;
          font-weight: 500;
          color: #333;
        }
        input[type="checkbox"] {
          margin-right: 8px;
          cursor: pointer;
          transform: scale(1.1);
        }
        input[type="date"], input[type="text"] {
          width: 100%;
          padding: 10px 12px;
          margin: 8px 0 16px;
          border: 1.5px solid #ccc;
          border-radius: 6px;
          font-size: 14px;
          box-sizing: border-box;
          transition: border-color 0.3s ease;
        }
        input[type="date"]:focus, input[type="text"]:focus {
          outline: none;
          border-color: #007bff;
          box-shadow: 0 0 6px rgba(0,123,255,0.4);
        }
        button {
          background-color: #007bff;
          color: white;
          border: none;
          padding: 12px 0;
          width: 100%;
          font-size: 16px;
          font-weight: 600;
          cursor: pointer;
          border-radius: 8px;
          transition: background-color 0.3s ease;
        }
        button:hover {
          background-color: #0056b3;
        }
        #projectList {
          max-height: 230px;
          overflow-y: auto;
          border: 1.5px solid #ccc;
          border-radius: 6px;
          padding: 12px 15px;
          background: white;
          margin-bottom: 20px;
          box-sizing: border-box;
        }
        #loader-overlay {
          display: none;
          position: fixed;
          top: 0; left: 0;
          width: 100%; height: 100%;
          background: rgba(255, 255, 255, 0.85);
          z-index: 9999;
          text-align: center;
          padding-top: 180px;
          font-size: 18px;
          color: #333;
          font-weight: 600;
          font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        .spinner {
          margin: 0 auto 20px;
          border: 6px solid #f3f3f3;
          border-top: 6px solid #007bff;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        /* Scrollbar styling */
        #projectList::-webkit-scrollbar {
          width: 8px;
        }
        #projectList::-webkit-scrollbar-track {
          background: #f1f1f1;
          border-radius: 6px;
        }
        #projectList::-webkit-scrollbar-thumb {
          background: #007bff;
          border-radius: 6px;
        }
      </style>
    </head>
    <body>
      <h3>Export UAT Bugs</h3>

      <label for="search">🔍 Search Projects:</label>
      <input type="text" id="search" placeholder="Type to filter projects...">

      <div id="projectList">
        ${projectOptions}
      </div>

      <label for="startDate">Start Date:</label>
      <input type="date" id="startDate">

      <label for="endDate">End Date:</label>
      <input type="date" id="endDate">

      <button onclick="submitForm()">📤 Export</button>

      <div id="loader-overlay">
        <div class="spinner"></div>
        Exporting bugs... Please wait.
      </div>

      <script>
        const searchInput = document.getElementById('search');
        const projectList = document.getElementById('projectList');
        // Convert NodeList to Array for easy filtering/manipulation
        const labels = Array.from(projectList.querySelectorAll('label'));

        searchInput.addEventListener('input', () => {
          const query = searchInput.value.toLowerCase().trim();

          // Clear all projects
          projectList.innerHTML = '';

          // Filter matched projects
          const matched = labels.filter(label => label.textContent.toLowerCase().includes(query));

          // Append matched projects on top (only matched shown)
          matched.forEach(label => projectList.appendChild(label));
        });

        function submitForm() {
          const checkboxes = document.querySelectorAll('input[name="project"]:checked');
          const selectedProjects = Array.from(checkboxes).map(cb => cb.value);
          const startDate = document.getElementById('startDate').value;
          const endDate = document.getElementById('endDate').value;

          if (selectedProjects.length === 0) {
            alert("Please select at least one project.");
            return;
          }
          if (!startDate || !endDate) {
            alert("Please select valid start and end dates.");
            return;
          }
          if (startDate > endDate) {
            alert("Start Date cannot be after End Date.");
            return;
          }

          document.getElementById("loader-overlay").style.display = "block";

          google.script.run
            .withSuccessHandler(msg => {
              alert(msg);
              google.script.host.close();
            })
            .exportUATBugsMultiProject(selectedProjects, startDate, endDate);
        }
      </script>
    </body>
    </html>
  `).setWidth(520).setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "Export UAT Bugs");
}


// Main export function for multiple projects
function exportUATBugsMultiProject(projectKeys, startDate, endDate) {
  const email = '';
   const apiToken = ''; // Set your Jira API token here
   const domain = 'eight25media.atlassian.net';

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totalIssues = 0;

  projectKeys.forEach(projectKey => {
    const sheetName = `UAT Report [${projectKey}] ${startDate} to ${endDate}`;

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) spreadsheet.deleteSheet(sheet);

    sheet = spreadsheet.insertSheet(sheetName);

    const jql = `project = "${projectKey}" AND created >= "${startDate}" AND created <= "${endDate}" AND type = "UAT Bug"`;
    const url = `https://${domain}/rest/api/3/search?jql=${encodeURIComponent(jql)}&fields=issuetype,key,summary,description,customfield_10043`;

    const headers = {
      "Authorization": "Basic " + Utilities.base64Encode(email + ":" + apiToken),
      "Accept": "application/json"
    };

    const response = UrlFetchApp.fetch(url, { method: "get", headers });
    const data = JSON.parse(response.getContentText());
    const issues = data.issues;

    sheet.appendRow([
      "Issue Type",
      "Issue key",
      "Issue id",
      "Summary",
      "Defect Category",
      "BH TYPE",
      "CSM COMMENTS",
      "CATEGORY",
      "BugHerd Link"
    ]);

   issues.forEach(issue => {
  const fields = issue.fields;

  // Extract plain description text
  let plainText = "";
  if (typeof fields.description === "string") {
    plainText = fields.description;
  } else if (fields.description && fields.description.content) {
    fields.description.content.forEach(block => {
      if (block.content) {
        block.content.forEach(inner => {
          if (inner.text) plainText += inner.text + " ";
        });
      }
    });
  } else if (fields.description && fields.description.content?.[0]?.text) {
    plainText = fields.description.content[0].text;
  }

  // Combine summary + description for BugHerd link check
  const combinedText = (fields.summary || "") + " " + plainText;

  // Extract BugHerd link if any
  let bugherdLink = "–";
  const match = combinedText.match(/https:\/\/www\.bugherd\.com\/projects\/\d+\/tasks\/\d+/);
  if (match) {
    bugherdLink = match[0];
  }

  sheet.appendRow([
    fields.issuetype.name,
    issue.key,
    issue.id,
    fields.summary,
    fields.customfield_10043 || "–",
    "",
    "",
    "",
    bugherdLink
  ]);
});


    totalIssues += issues.length;
  });

  return `${totalIssues} UAT bugs exported across ${projectKeys.length} project(s).`;
}

