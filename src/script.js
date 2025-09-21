document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("file-input");
  const dropArea = document.querySelector(".upload-card");
  const tableContainer = document.getElementById("table-container");
  const downloadBtn = document.getElementById("download-btn");
  const status = document.getElementById("status");
  const counter = document.getElementById("counter");

  // Drag and drop functionality
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropArea.addEventListener(eventName, highlight, false);
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, unhighlight, false);
  });

  function highlight() {
    dropArea.style.border = "2px dashed " + "#4361ee";
    status.textContent = "Drop file to upload";
    status.className = "status info";
  }

  function unhighlight() {
    dropArea.style.border = "";
    status.textContent = "Ready to upload your Excel file";
    status.className = "status info";
  }

  dropArea.addEventListener("drop", handleDrop, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length) {
      fileInput.files = files;
      handleFile(files[0]);
    }
  }

  fileInput.addEventListener("change", function () {
    if (this.files.length) {
      handleFile(this.files[0]);
    }
  });

  function handleFile(file) {
    status.textContent = `Processing: ${file.name}...`;
    status.className = "status info";

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      processExcel(data);
    };
    reader.onerror = function () {
      status.textContent = "Error reading file";
      status.className = "status error";
    };
    reader.readAsArrayBuffer(file);
  }

  function processExcel(data) {
    try {
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Filter out empty rows
      const filteredData = jsonData.filter(
        (row) =>
          row.length > 0 && row.some((cell) => cell !== null && cell !== "")
      );

      if (filteredData.length === 0) {
        status.textContent = "No data found in the Excel file";
        status.className = "status error";
        return;
      }

      // Display preview
      displayPreview(filteredData);

      // Generate VCF
      generateVCF(filteredData);

      const contactCount =
        filteredData.length - (isHeaderRow(filteredData[0]) ? 1 : 0);
      status.textContent = `Successfully processed ${contactCount} contacts`;
      status.className = "status success";
      counter.textContent = `${contactCount} contacts`;
    } catch (error) {
      console.error(error);
      status.textContent =
        "Error processing file. Please make sure it is a valid Excel file.";
      status.className = "status error";
    }
  }

  function displayPreview(data) {
    let tableHTML = "<table><thead><tr>";

    // Create headers
    for (let i = 0; i < data[0].length; i++) {
      tableHTML += `<th>Column ${i + 1}</th>`;
    }
    tableHTML += "</tr></thead><tbody>";

    // Create rows (show all rows)
    for (let i = 0; i < data.length; i++) {
      tableHTML += "<tr>";
      for (let j = 0; j < data[i].length; j++) {
        tableHTML += `<td>${data[i][j] || ""}</td>`;
      }
      tableHTML += "</tr>";
    }

    tableHTML += "</tbody></table>";

    tableContainer.innerHTML = tableHTML;
  }

  function generateVCF(data) {
    let vcfContent = "";

    // Start from index 1 to skip header row if exists
    const startIndex = data.length > 1 && isHeaderRow(data[0]) ? 1 : 0;
    let contactCount = 0;

    for (let i = startIndex; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 2) {
        const name = formatValue(row[0]);
        const phone = formatValue(row[1]);

        if (name && phone) {
          vcfContent += `BEGIN:VCARD
VERSION:3.0
FN:${name}
TEL;TYPE=CELL:${phone}
END:VCARD

`;
          contactCount++;
        }
      }
    }

    if (contactCount > 0) {
      downloadBtn.disabled = false;

      downloadBtn.onclick = function () {
        const blob = new Blob([vcfContent], { type: "text/vcard" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "contacts.vcf";
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        }, 100);
      };
    } else {
      downloadBtn.disabled = true;
      status.textContent = "No valid contacts found in the Excel file";
      status.className = "status error";
    }
  }

  function formatValue(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function isHeaderRow(row) {
    // Simple check to see if the row might be a header
    if (row.length === 0) return false;

    const secondCell = row.length > 1 ? String(row[1]) : "";
    // If second cell doesn't look like a phone number, it might be a header
    return !/\d{5,}/.test(secondCell);
  }
});
