// ==========================================================
// Google Apps Script - CSR Questionnaire → Google Sheet
// ==========================================================
// INSTRUCTIONS:
// 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1GHbnY0o34WUhY-Cl0C2MoZLC5uXWeiXxOt6lbDqNCBA/
// 2. Go to Extensions → Apps Script
// 3. Delete all existing code and paste this entire file
// 4. Click "Deploy" → "New deployment"
//    - Type: "Web app"
//    - Execute as: "Me"
//    - Who has access: "Anyone"
// 5. Click "Deploy" and authorize when prompted
// 6. Copy the Web App URL and paste it in index.html at line:
//    const SCRIPT_URL = 'YOUR_URL_HERE';
// ==========================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = e.parameter;

    // For checkbox fields (multiple values), we use getAll()
    var allData = e.parameters; // parameters (plural) returns arrays for multi-value fields

    // Helper: join array values with ", " for checkbox fields
    function getMulti(fieldName) {
      if (allData[fieldName]) {
        return allData[fieldName].join(", ");
      }
      return "";
    }

    // Helper: get single value
    function getSingle(fieldName) {
      return data[fieldName] || "";
    }

    // Build the row with all form fields
    var row = [
      new Date(),                              // Timestamp
      
      // Section 1: Identification
      getSingle("email"),                      // Q1 - Email
      getSingle("company_name"),               // Q2 - Company Name
      getSingle("address"),                    // Q3 - Address
      getSingle("siret"),                      // Q4 - SIRET
      getSingle("respondent_name"),            // Q5 - Respondent Name
      getSingle("respondent_title"),           // Q6 - Respondent Title
      getSingle("csr_contact"),                // Q7 - CSR Contact

      // Section 2: General CSR Commitment
      getSingle("structured_csr"),             // Q8 - Structured CSR (Yes/No)

      // Section 3A: Detailed Assessment (Branch: Yes)
      getSingle("csr_labeled"),                // Q9  - CSR labeled
      getSingle("csr_label_details"),          // Q10 - Label details
      getSingle("csr_signatory"),              // Q11 - CSR signatory
      getSingle("csr_signatory_details"),      // Q12 - Signatory details
      getSingle("csr_responsible_exists"),     // Q13 - CSR responsible exists
      getSingle("csr_resp_name"),              // Q14 - CSR resp name
      getSingle("csr_resp_title"),             // Q15 - CSR resp title
      getSingle("csr_resp_email"),             // Q16 - CSR resp email
      getSingle("csr_report"),                 // Q17 - CSR report
      getSingle("csr_report_link"),            // Q18 - CSR report link
      getSingle("code_of_conduct"),            // Q19 - Code of conduct
      getSingle("whistleblowing"),             // Q20 - Whistleblowing

      // Working Conditions and Human Rights
      getSingle("human_rights_policy"),        // Q21 - Human rights policy
      getMulti("hr_areas"),                    // Q22 - HR areas (checkboxes)
      getSingle("hr_other"),                   // Q23 - HR other

      // Occupational Health and Safety
      getSingle("ohs_policy"),                 // Q24 - OHS policy
      getSingle("ohs_actions"),                // Q25 - OHS actions
      getSingle("ohs_examples"),               // Q26 - OHS examples

      // Business Ethics
      getSingle("ethics_policy"),              // Q27 - Ethics policy
      getMulti("ethics_areas"),                // Q28 - Ethics areas (checkboxes)
      getSingle("ethics_other"),               // Q29 - Ethics other

      // Environment
      getSingle("env_policy"),                 // Q30 - Env policy
      getSingle("env_system"),                 // Q31 - Env system
      getSingle("env_kpi"),                    // Q32 - Env KPI
      getSingle("env_cert_details"),           // Q33 - Env cert details
      getSingle("substances"),                 // Q34 - Substances
      getSingle("substances_proc"),            // Q35 - Substances procedures

      // Responsible Procurement
      getSingle("supplier_csr"),               // Q36 - Supplier CSR
      getMulti("supplier_comm"),               // Q37 - Supplier comm (checkboxes)
      getSingle("supplier_other"),             // Q38 - Supplier other

      // Training
      getSingle("training_sessions"),          // Q39 - Training sessions

      // Section 3B: Unstructured (Branch: No)
      getSingle("informal_person"),            // Q40 - Informal person
      getSingle("informal_contact_details"),   // Q41 - Informal contact
      getMulti("basic_kpi"),                   // Q42 - Basic KPI (checkboxes)
      getSingle("basic_kpi_other"),            // Q43 - Basic KPI other
      getMulti("written_rules"),               // Q44 - Written rules (checkboxes)
      getSingle("written_rules_other"),        // Q45 - Written rules other
      getSingle("support_interest"),           // Q46 - Support interest

      // Section 4: Common Ending
      getSingle("waste_measure"),              // Q47 - Waste measure
      getSingle("waste_reduce"),               // Q48 - Waste reduce
      getSingle("waste_examples"),             // Q49 - Waste examples
      getSingle("recycling"),                  // Q50 - Recycling
      getSingle("recycling_types"),            // Q51 - Recycling types
      getSingle("energy_measure"),             // Q52 - Energy measure
      getSingle("energy_reduce"),              // Q53 - Energy reduce
      getSingle("energy_examples"),            // Q54 - Energy examples
      getSingle("water_measure"),              // Q55 - Water measure
      getSingle("water_reduce"),               // Q56 - Water reduce
      getSingle("water_examples"),             // Q57 - Water examples
      getSingle("transport_actions"),          // Q58 - Transport actions
      getSingle("transport_examples"),         // Q59 - Transport examples
      getSingle("co2_measure"),                // Q60 - CO2 measure
      getSingle("ecodesign"),                  // Q61 - Eco-design
      getSingle("ecodesign_products"),         // Q62 - Eco-designed products
      getSingle("comments"),                   // Q63 - Comments
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================================
// Run this function ONCE to create the header row
// ==========================================================
function createHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = [
    "Timestamp",
    "Email",
    "Company Name",
    "Address",
    "SIRET",
    "Respondent Name",
    "Respondent Title",
    "CSR Contact",
    "Q8 - Structured CSR",
    "Q9 - CSR Labeled",
    "Q10 - Label Details",
    "Q11 - CSR Signatory",
    "Q12 - Signatory Details",
    "Q13 - CSR Responsible",
    "Q14 - CSR Resp Name",
    "Q15 - CSR Resp Title",
    "Q16 - CSR Resp Email",
    "Q17 - CSR Report",
    "Q18 - Report Link",
    "Q19 - Code of Conduct",
    "Q20 - Whistleblowing",
    "Q21 - Human Rights Policy",
    "Q22 - HR Areas",
    "Q23 - HR Other",
    "Q24 - OHS Policy",
    "Q25 - OHS Actions",
    "Q26 - OHS Examples",
    "Q27 - Ethics Policy",
    "Q28 - Ethics Areas",
    "Q29 - Ethics Other",
    "Q30 - Env Policy",
    "Q31 - Env System",
    "Q32 - Env KPI",
    "Q33 - Env Cert Details",
    "Q34 - Substances",
    "Q35 - Substances Procedures",
    "Q36 - Supplier CSR",
    "Q37 - Supplier Communication",
    "Q38 - Supplier Other",
    "Q39 - Training Sessions",
    "Q40 - Informal Person",
    "Q41 - Informal Contact",
    "Q42 - Basic KPI",
    "Q43 - Basic KPI Other",
    "Q44 - Written Rules",
    "Q45 - Written Rules Other",
    "Q46 - Support Interest",
    "Q47 - Waste Measure",
    "Q48 - Waste Reduce",
    "Q49 - Waste Examples",
    "Q50 - Recycling",
    "Q51 - Recycling Types",
    "Q52 - Energy Measure",
    "Q53 - Energy Reduce",
    "Q54 - Energy Examples",
    "Q55 - Water Measure",
    "Q56 - Water Reduce",
    "Q57 - Water Examples",
    "Q58 - Transport Actions",
    "Q59 - Transport Examples",
    "Q60 - CO2 Measure",
    "Q61 - Eco-Design",
    "Q62 - Eco-Designed Products",
    "Q63 - Comments",
  ];

  // Insert headers in row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Bold and freeze the header row
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}
