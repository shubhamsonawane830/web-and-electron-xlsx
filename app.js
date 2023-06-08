const fileInput = document.getElementById("xlsxFile");
const convertBtn = document.getElementById("convertBtn");
const downloadBtn = document.getElementById("downloadBtn");
const sheetCheckboxContainer = document.getElementById(
  "sheetCheckboxContainer"
);
const selectAllBtn = document.getElementById("selectAllBtn");
const clearBtn = document.getElementById("clearBtn");

// Create an element to display the progress message
const progressMessageElement = document.getElementById("progressMessage");

convertBtn.addEventListener("click", convertToJSON);

const logFilesDownloadLocation = "logs/";

fileInput.addEventListener("change", function () {
  sheetCheckboxContainer.innerHTML = ""; // Clear existing checkboxes
  sheetCheckboxContainer.style.display = "block";

  const file = fileInput.files[0];

  if (!file) {
    alert("Please select an XLSX file.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    let workbook;
    try {
      workbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      alert("Invalid XLSX file. Please select a valid file.");
      return;
    }

    const sheetNames = workbook.SheetNames;

    sheetNames.forEach((sheetName) => {
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = sheetName;
      checkbox.checked = true;
      sheetCheckboxContainer.appendChild(checkbox);

      const label = document.createElement("label");
      label.appendChild(document.createTextNode(sheetName));
      sheetCheckboxContainer.appendChild(label);
      sheetCheckboxContainer.appendChild(document.createElement("br"));
    });

    // Reset progress bar and percentage
    document.getElementById("conversionProgress").value = 0;
    document.getElementById("progressPercentage").textContent = "0%";
  };

  reader.readAsArrayBuffer(file);
});

selectAllBtn.addEventListener("click", function () {
  const checkboxes = document.querySelectorAll(
    "#sheetCheckboxContainer input[type='checkbox']"
  );
  checkboxes.forEach((checkbox) => (checkbox.checked = true));
  updateDownloadButton();
});

clearBtn.addEventListener("click", function () {
  const checkboxes = document.querySelectorAll(
    "#sheetCheckboxContainer input[type='checkbox']"
  );
  checkboxes.forEach((checkbox) => (checkbox.checked = false));
  updateDownloadButton();
});

// Function to check if a cell value has a math formula
function hasMathFormula(cellValue, allowedFormulas) {
  if (typeof cellValue !== "string") {
    return false;
  }

  const formulaPattern = /^=([A-Z]+)\(/;
  const matches = cellValue.match(formulaPattern);
  if (matches && matches.length > 1) {
    const formula = matches[1];
    return !allowedFormulas.includes(formula);
  }
  return false;
}

// Function to clear existing download links
function clearDownloadLinks() {
  const existingLinks = document.querySelectorAll(".download-link");
  existingLinks.forEach((link) => link.remove());
}

// Function to generate the error log file
function generateErrorLog(invalidCellInfo) {
  if (invalidCellInfo.length === 0) {
    // No errors, no need to generate the error log
    return;
  }

  const timestamp = new Date().toISOString();
  let errorLogMessage = `Error Log (${timestamp}):\n\n`;

  invalidCellInfo.forEach((errorInfo) => {
    errorLogMessage += `Sheet: ${errorInfo.sheet}\n`;
    errorLogMessage += `Value: ${errorInfo.value}\n`;
    errorLogMessage += `Column: ${errorInfo.column}\n`;
    errorLogMessage += `Row: ${errorInfo.row}\n`;
    errorLogMessage += `Reasons: ${errorInfo.reason.join(", ")}\n\n`;
  });

  const errorLogContent = errorLogMessage;
  const blob = new Blob([errorLogContent], { type: "text/plain" });
  const downloadLink = URL.createObjectURL(blob);

  const anchor = document.createElement("a");
  anchor.href = downloadLink;
  anchor.download = `error_log_${timestamp}.txt`;
  anchor.className = "download-link";
  anchor.textContent = `error_log_${timestamp}.txt`;
  anchor.style.display = "block";

  document.body.appendChild(anchor);
}

// Function to convert XLSX file to JSON
function convertToJSON() {
  const invalidCellInfo = []; // Store information about invalid cells
  const convertedSheets = []; // Store names of successfully converted sheets
  const notConvertedSheets = []; // Store names of sheets that were not converted

  // Clear existing download links and error log link
  clearDownloadLinks();
  const existingErrorLogLink = document.querySelector(".error-log-link");
  if (existingErrorLogLink) {
    existingErrorLogLink.remove();
  }

  const checkboxes = document.querySelectorAll(
    "#sheetCheckboxContainer input[type='checkbox']"
  );
  const selectedSheets = Array.from(checkboxes).filter(
    (checkbox) => checkbox.checked
  );

  if (selectedSheets.length === 0) {
    alert("Please select at least one sheet.");
    return;
  }

  const file = fileInput.files[0];

  if (!file) {
    alert("Please select an XLSX file.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    let workbook;
    try {
      workbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      alert("Invalid XLSX file. Please select a valid file.");
      return;
    }

    const jsonArray = [];
    const totalSheets = selectedSheets.length;
    let completedSheets = 0;
    let convertedFiles = [];

    selectedSheets.forEach((checkbox) => {
      const sheetName = checkbox.value;
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet);

      const headerRow = sheetData[0];
      const columnKeys = Object.keys(headerRow);

      const exceptionCharacters = [
        "-",
        "_",
        " ",
        ".",
        "\u00ED",
        "\u00F3",
        "\u00E9",
      ]; // Include í, ó, and é as exceptions

      const allowedFormulas = ["SUM", "AVERAGE", "MAX", "MIN"];

      const specialCharactersRegex = new RegExp(
        `[^\\w${exceptionCharacters.join("\\\\")}]`,
        "g"
      );

      const jsonSheetData = {
        sheetName: sheetName,
        data: sheetData,
      };

      jsonArray.push(jsonSheetData);

      let hasErrors = false;

      sheetData.forEach((row, rowIndex) => {
        for (const key in row) {
          const cellValue = row[key];

          if (
            specialCharactersRegex.test(cellValue) ||
            cellValue.length > 128 ||
            hasMathFormula(cellValue, allowedFormulas)
          ) {
            const columnName = columnKeys.find((colKey) => colKey === key);
            const errorInfo = {
              sheet: sheetName,
              value: cellValue,
              column: columnName,
              row: rowIndex + 1,
              reason: [],
            };

            if (specialCharactersRegex.test(cellValue)) {
              errorInfo.reason.push("Contains invalid special character(s).");
            }
            if (cellValue.length > 128) {
              errorInfo.reason.push(
                "Exceeds the maximum allowed length of 128 characters."
              );
            }
            if (hasMathFormula(cellValue, allowedFormulas)) {
              errorInfo.reason.push("Contains disallowed formula(s).");
            }

            invalidCellInfo.push(errorInfo);
            hasErrors = true;
          }
        }
      });

      if (!hasErrors) {
        const jsonContent = JSON.stringify(jsonSheetData, null, 2);
        const blob = new Blob([jsonContent], { type: "application/json" });
        const downloadLink = URL.createObjectURL(blob);

        const anchor = document.createElement("a");
        anchor.href = downloadLink;
        anchor.download = `${sheetName}.json`;
        anchor.className = "download-link";
        anchor.textContent = `${sheetName}.json`;
        anchor.style.display = "block";

        document.body.appendChild(anchor);

        convertedSheets.push(sheetName);
      } else {
        notConvertedSheets.push(sheetName);
      }

      completedSheets++;

      // Update progress bar
      updateProgressBar(completedSheets, totalSheets);

      if (completedSheets === totalSheets) {
        if (invalidCellInfo.length > 0) {
          let alertMessage =
            "Conversion completed with errors. Please check the error log.\n\n";
          if (convertedSheets.length > 0) {
            alertMessage +=
              "The following sheets were converted successfully:\n" +
              convertedSheets.join(", ") +
              "\n\n";
          }
          if (notConvertedSheets.length > 0) {
            alertMessage +=
              "The following sheets could not be converted:\n" +
              notConvertedSheets.join(", ");
          }
          alert(alertMessage);
        } else {
          let alertMessage = "Conversion completed successfully.\n\n";
          if (convertedSheets.length > 0) {
            alertMessage +=
              "The following sheets were converted successfully:\n" +
              convertedSheets.join(", ") +
              "\n\n";
          }
          if (notConvertedSheets.length > 0) {
            alertMessage +=
              "The following sheets could not be converted:\n" +
              notConvertedSheets.join(", ");
          }
          alert(alertMessage);
        }

        generateErrorLog(invalidCellInfo);
        generateDownloadLinks(convertedSheets);
      }
    });
  };

  reader.readAsArrayBuffer(file);
}

// Function to update the progress bar
function updateProgressBar(completedSheets, totalSheets) {
  const progress = (completedSheets / totalSheets) * 100;
  document.getElementById("conversionProgress").value = progress;
  document.getElementById("progressPercentage").textContent = `${progress}%`;
}

// Function to reset the progress bar
function resetProgressBar() {
  document.getElementById("conversionProgress").value = 0;
  document.getElementById("progressPercentage").textContent = "0%";
}

function updateDownloadButton() {
  const checkboxes = document.querySelectorAll(
    "#sheetCheckboxContainer input[type='checkbox']"
  );
  const selectedSheets = Array.from(checkboxes).filter(
    (checkbox) => checkbox.checked
  );

  if (selectedSheets.length > 0) {
    downloadBtn.disabled = false;
  } else {
    downloadBtn.disabled = true;
  }
}
