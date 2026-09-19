<!-- Division Digital Summary Report Section -->
<div class="container my-3">
  <div class="card shadow-sm p-3">
    <h4 class="mb-3">Division Digital Summary Report</h4>
    <div class="table-responsive">
      <table class="table table-bordered align-middle text-nowrap mb-0" id="divisionSummaryTable">
        <thead style="background-color: #f7caac; color: #000;">
          <tr>
            <th>Division</th>
            <th class="text-end">Total Cash Txn</th>
            <th class="text-end">Total Digital Txn</th>
            <th class="text-end">Total Txn</th>
            <th class="text-end">% of Digital Txn</th>
          </tr>
        </thead>
        <tbody id="summaryTableBody">
          <!-- Data JavaScript se yahan aayega -->
        </tbody>
        <tfoot id="summaryTableFooter" style="background-color: #f7caac; font-weight: bold;">
          <!-- Total Row -->
        </tfoot>
      </table>
    </div>
  </div>
</div>
2. File Upload वाले Event Handler को Check करें
आपकी app.js में फाइल अपलोड होते ही टेबल अपडेट करने के लिए readWorkbook फंक्शन में generateDivisionSummary() का सही जगह कॉल होना जरूरी है।

app.js में readWorkbook फंक्शन को इस तरह अपडेट करें:

JavaScript
async function readWorkbook(arrayBuffer, sourceName) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets["Daily Booking Report"] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("Daily Booking Report sheet not found.");
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: 0 });
  if (!raw.length) throw new Error("The sheet is empty.");
  
  headers = raw[0];
  classifyColumns();
  allRows = raw.slice(1).filter(r => r.length && r[0] !== null && r[0] !== undefined && r[0] !== "");
  
  $("sourceStatus").textContent = `Source: ${sourceName} • ${allRows.length.toLocaleString("en-IN")} office rows`;
  
  buildDivisionOptions();
  
  // Highlighting: File load hote hi Summary Table render hogi
  generateDivisionSummary(); 
  
  applyFilters();
}
  
  applyFilters();
}
