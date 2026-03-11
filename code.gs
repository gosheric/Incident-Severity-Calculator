/**
 * Serves the HTML file as a web app.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IT Incident Severity Calculator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches lists from CMDB tabs and defines question options.
 */
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const locationSheet = ss.getSheetByName('CMDB Location Tiering');
  const locations = locationSheet ? 
    locationSheet.getRange("A2:A" + locationSheet.getLastRow()).getValues().flat().filter(String) : [];

  const assetSheet = ss.getSheetByName('CMDB Asset Tiering');
  const assets = assetSheet ? 
    assetSheet.getRange("B2:B" + assetSheet.getLastRow()).getValues().flat().filter(String) : [];

  return {
    locations: [...new Set(locations)].sort(),
    assets: [...new Set(assets)].sort(),
    
    q1LocationOptions: ['2 = Only One Station/Location Impacted', '3 = More than One Station/Location Impacted', '4 = More than Five Stations/Locations Impacted', '5 = All stations Impacted'],
    q2LocationOptions: ['2 = Full Workaround available', '3 = Limited Workaround available (Manual Process)', '4 = Limited Workaround available', '5 = No Workaround available'],
    q3LocationOptions: ['2 = Tier 4 Application/System/Product/Service', '3 = Tier 3 Application/System/Product/Service', '4 = Tier 2 Application/System/Product/Service', '5 = Tier 1 Application/System/Product/Service'],

    q1AppOptions: ['1 = Single user', '2 = Single department', '3 = Multiple departments', '4 = Entire organization', '5 = Total outage (Global)'],
    q2AppOptions: ['1 = Full workaround', '2 = Partial workaround', '3 = Temporary manual process', '4 = Limited workaround', '5 = No workaround'],
    q3AppOptions: ['1 = Non-critical system', '2 = Supporting system', '3 = Important System', '4 = Critical System', '5 = Compliance/security critical']
  };
}

/**
 * Parses numeric score from selection strings.
 */
function extractScore(val) {
  if (!val) return 0;
  const match = val.toString().match(/\d+/);
  return match ? parseInt(match[0], 10) : 0;
}

/**
 * Main Calculation Engine - Updates the "Airport Sev" tab.
 */
function calculateSeverity(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Airport Sev');
  
  const s1 = extractScore(data.q1);
  const s2 = extractScore(data.q2);
  const s3 = extractScore(data.q3);
  const totalScore = s1 + s2 + s3;
  const hasLevelFive = (s1 === 5 || s2 === 5 || s3 === 5);

  let priority = "";
  let severity = "";
  let tierValue = "N/A";

  // --- 1. Base Scoring Logic ---
  if (hasLevelFive || totalScore >= 15) {
    priority = "P0 - Critical (15-30 min SLA)";
    severity = "Sev 1 = Services impacting everyone / Business Escalation";
  } else if (data.type === 'Location') {
    if (totalScore >= 11) {
      priority = "P1 - Major (1-2 hours SLA)";
      severity = "Sev 2 = Services impacting a few groups/stations";
    } else if (totalScore >= 8) {
      priority = "P2 - High (4-8 hours SLA)";
      severity = "Sev 3 = Services impacting a department";
    } else {
      priority = "P3 - Low (24+ hours SLA)";
      severity = "Sev 4 = Services impacting a single user";
    }
  } else {
    if (totalScore >= 10) {
      priority = "P1 - Major (1-2 hours SLA)";
      severity = "Sev 2 = Services impacting a few groups/stations";
    } else if (totalScore >= 7) {
      priority = "P2 - High (4-8 hours SLA)";
      severity = "Sev 3 = Services impacting a department";
    } else {
      priority = "P3 - Low (24+ hours SLA)";
      severity = "Sev 4 = Services impacting a single user";
    }
  }

  // --- 2. Location Specific Checks ---
  if (data.type === 'Location') {
    const locSheet = ss.getSheetByName('CMDB Location Tiering');
    if (locSheet) {
      const locData = locSheet.getDataRange().getValues();
      const match = locData.find(row => row[0] === data.selection);
      if (match) {
        tierValue = match[3];
        sheet.getRange('C5').setValue(tierValue);
        
        // TIER 1 OVERRIDE: Bump to P1 if Tier 1 and score is low
        if (!hasLevelFive && totalScore < 11 && (tierValue == "1" || tierValue.toString().toLowerCase().includes("tier 1"))) {
          priority = "P1 - Major (1-2 hours SLA) [Tier 1 Override]";
          severity = "Sev 2 = Major (Tier 1 Priority)";
        }
      }
    }

    // --- 3. CRITICAL SITE OVERRIDE (Mandatory P0/Sev 1) ---
    const criticalSites = [
      "CHINA DC", 
      "CYBERJAYA (PRISMA 9) - BCP", 
      "CYBERPOINT FINEXUS", 
      "REDQ DC", 
      "VITRO DC", 
      "TITIWANGSA FINEXUS"
    ];
    if (criticalSites.indexOf(data.selection) !== -1) {
      priority = "P0 - Critical (15-30 min SLA) [Critical Site]";
      severity = "Sev 1 = Services impacting everyone / Business Escalation";
    }

    // Record results to Spreadsheet
    sheet.getRange('C4').setValue(data.selection); 
    sheet.getRange('C6').setValue(data.q1);
    sheet.getRange('C7').setValue(data.q2);
    sheet.getRange('C8').setValue(data.q3);
    sheet.getRange('E4').setValue(totalScore);    
    sheet.getRange('E5').setValue(priority);      
    sheet.getRange('C9').setValue(severity);
  } else {
    // Application results
    sheet.getRange('C17').setValue(data.selection);
    sheet.getRange('C19').setValue(data.q1);
    sheet.getRange('C20').setValue(data.q2);
    sheet.getRange('C21').setValue(data.q3);
    sheet.getRange('E17').setValue(totalScore);
    sheet.getRange('E18').setValue(priority);
    sheet.getRange('C22').setValue(severity);
    
    const appSheet = ss.getSheetByName('CMDB Asset Tiering');
    if (appSheet) {
       const appData = appSheet.getDataRange().getValues();
       const match = appData.find(row => row[1] === data.selection);
       if (match) {
         tierValue = match[3];
         sheet.getRange('C18').setValue(tierValue);
       }
    }
  }

  SpreadsheetApp.flush();
  return { score: totalScore, priority: priority, severity: severity, tier: tierValue };
}
