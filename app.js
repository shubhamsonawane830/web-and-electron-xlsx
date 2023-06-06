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

    updateDownloadButton();

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

function convertToJSON() {
  const invalidCellInfo = []; // Store information about invalid cells

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
    const invalidCellInfo = [];
    const totalSheets = selectedSheets.length;
    let completedSheets = 0;
    let convertedFiles = [];
    let downloadedFiles = [];
    let downloadLink;

    // Add a new array to store the converted JSON files
    let jsonFiles = [];

    const downloadLinks = [];

    selectedSheets.forEach((checkbox) => {
      const sheetName = checkbox.value;
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet);

      const jsonContent = JSON.stringify(sheetData, null, 2);
      const blob = new Blob([jsonContent], { type: "application/json" });
      const downloadLink = URL.createObjectURL(blob);

      downloadLinks.push({
        sheetName: sheetName,
        downloadLink: downloadLink,
      });
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
      // new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');

      // const progress = Math.floor((completedSheets / totalSheets) * 100);
      document.getElementById("conversionProgress").max = totalSheets;
      // const progress = completedSheets + 1;
      const progress = totalSheets;

      document.getElementById("conversionProgress").value = progress;
      document.getElementById("progressPercentage").textContent = 100 + "%";
      // Math.floor((progress / totalSheets) * 100) + "%";

      // document.getElementById("conversionProgress").value = progress;
      // document.getElementById("progressPercentage").textContent =
      //   progress + "%";

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
              errorInfo.reason.push("Contains an invalid math formula.");
            }

            invalidCellInfo.push(errorInfo);
            hasErrors = true;
          }
        }
      });

      if (hasErrors) {
        console.log(
          `Skipping JSON conversion for sheet ${sheetName} due to errors.`
        );
        return; // Skip JSON conversion for this sheet
      }

      completedSheets++;

      // Add the converted sheet name to the convertedFiles array
      convertedFiles.push(sheetName);
    });

    // Update the progress message on the UI
    const progressMessage = `JSON conversion completed for ${completedSheets} sheet(s) of  total ${totalSheets} sheet(s).`;
    progressMessageElement.textContent = progressMessage;

    // Display completion message as a popup
    const message = `Conversion completed for the following sheet(s):\n\n${convertedFiles.join(
      "\n"
    )}`;
    alert(message);

    if (invalidCellInfo.length > 0) {
      const errorMessage =
        "Error: The XLSX data contains invalid cells. Please check the log file for details.";
      alert(errorMessage);

      const timestamp = new Date().toISOString().replace(/[-:.]/g, "");
      const logFileName = logFilesDownloadLocation + `error_${timestamp}.txt`;

      let logFileContent = `${errorMessage}\n\n`;

      invalidCellInfo.forEach((errorInfo) => {
        logFileContent += `Sheet: ${errorInfo.sheet}\n`;
        logFileContent += `Cell Value: ${errorInfo.value}\n`;
        logFileContent += `Column: ${errorInfo.column}\n`;
        logFileContent += `Row: ${errorInfo.row}\n`;
        logFileContent += `Reason: ${errorInfo.reason.join(", ")}\n\n`;
      });

      const blob = new Blob([logFileContent], { type: "text/plain" });
      saveAs(blob, logFileName);

      return;
    }

    if (completedSheets === 0) {
      alert("No data found in the selected sheets.");
      return;
    }

    downloadBtn.disabled = false;

    downloadBtn.addEventListener("click", function () {
      const oldLinks = document.querySelectorAll("#downloadLinksContainer a");
      oldLinks.forEach((link) => {
        link.remove();
      });
      convertToJSONDownloadclick();

      convertedFiles.forEach((sheetName) => {
        const jsonContent = JSON.stringify(
          jsonArray.find((json) => json.sheetName === sheetName).data,
          null,
          2
        );
        const blob = new Blob([jsonContent], { type: "application/json" });
        const downloadLink = URL.createObjectURL(blob);
        if (downloadLink) {
          const anchor = document.createElement("a");
          anchor.href = downloadLink;
          anchor.download = `${sheetName}.json`;
          anchor.textContent = `${sheetName}.json`;
          anchor.style.display = "block";

          document.getElementById("downloadLinksContainer").appendChild(anchor);
        }
      });

      // downloadLinks.forEach((link) => {
      //   const anchor = document.createElement("a");
      //   anchor.href = link.downloadLink;
      //   anchor.download = `${link.sheetName}.json`;
      //   anchor.textContent = `${link.sheetName}.json`;
      //   anchor.style.display = "block";

      //   document.body.appendChild(anchor);

      //   // Trigger the click event to initiate the download
      //   anchor.click();

      //   // Remove the dynamically created anchor element
      //   document.body.removeChild(anchor);
      // });

      // jsonArray.forEach((jsonSheetData) => {
      //   const sheetName = jsonSheetData.sheetName;
      //   const sheetData = jsonSheetData.data;

      //   const jsonContent = JSON.stringify(sheetData, null, 2);
      //   const blob = new Blob([jsonContent], { type: "application/json" });
      //   const fileName = `${sheetName}.json`;
      //   saveAs(blob, fileName);

      //   // // Add the downloaded file name to the downloadedFiles array
      //   downloadedFiles.push(fileName);
      // });
    });
  };

  reader.readAsArrayBuffer(file);
}

////////////

function convertToJSONDownloadclick() {
  const invalidCellInfo = []; // Store information about invalid cells

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
    const invalidCellInfo = [];
    const totalSheets = selectedSheets.length;
    let completedSheets = 0;
    let convertedFiles = [];
    let downloadedFiles = [];
    let downloadLink;

    // Add a new array to store the converted JSON files
    let jsonFiles = [];

    const downloadLinks = [];

    selectedSheets.forEach((checkbox) => {
      const sheetName = checkbox.value;
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet);

      const jsonContent = JSON.stringify(sheetData, null, 2);
      const blob = new Blob([jsonContent], { type: "application/json" });
      const downloadLink = URL.createObjectURL(blob);

      downloadLinks.push({
        sheetName: sheetName,
        downloadLink: downloadLink,
      });
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
      // new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');

      // const progress = Math.floor((completedSheets / totalSheets) * 100);
      document.getElementById("conversionProgress").max = totalSheets;
      // const progress = completedSheets + 1;
      const progress = totalSheets;

      document.getElementById("conversionProgress").value = progress;
      document.getElementById("progressPercentage").textContent = 100 + "%";
      // Math.floor((progress / totalSheets) * 100) + "%";

      // document.getElementById("conversionProgress").value = progress;
      // document.getElementById("progressPercentage").textContent =
      //   progress + "%";

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
              errorInfo.reason.push("Contains an invalid math formula.");
            }

            invalidCellInfo.push(errorInfo);
            hasErrors = true;
          }
        }
      });

      if (hasErrors) {
        console.log(
          `Skipping JSON conversion for sheet ${sheetName} due to errors.`
        );
        return; // Skip JSON conversion for this sheet
      }

      completedSheets++;

      // Add the converted sheet name to the convertedFiles array
      convertedFiles.push(sheetName);
    });

    // Update the progress message on the UI
    const progressMessage = `JSON conversion completed for ${completedSheets} sheet(s) of  total ${totalSheets} sheet(s).`;
    progressMessageElement.textContent = progressMessage;

    // Display completion message as a popup
    const message = `Conversion completed for the following sheet(s):\n\n${convertedFiles.join(
      "\n"
    )}`;
    // alert(message);

    if (invalidCellInfo.length > 0) {
      const errorMessage =
        "Error: The XLSX data contains invalid cells. Please check the log file for details.";
      // alert(errorMessage);

      const timestamp = new Date().toISOString().replace(/[-:.]/g, "");
      const logFileName = logFilesDownloadLocation + `error_${timestamp}.txt`;

      let logFileContent = `${errorMessage}\n\n`;

      invalidCellInfo.forEach((errorInfo) => {
        logFileContent += `Sheet: ${errorInfo.sheet}\n`;
        logFileContent += `Cell Value: ${errorInfo.value}\n`;
        logFileContent += `Column: ${errorInfo.column}\n`;
        logFileContent += `Row: ${errorInfo.row}\n`;
        logFileContent += `Reason: ${errorInfo.reason.join(", ")}\n\n`;
      });

      const blob = new Blob([logFileContent], { type: "text/plain" });
      saveAs(blob, logFileName);

      return;
    }

    if (completedSheets === 0) {
      alert("No data found in the selected sheets.");
      return;
    }

    downloadBtn.disabled = false;

    downloadBtn.addEventListener("click", function () {
      const oldLinks = document.querySelectorAll("#downloadLinksContainer a");
      oldLinks.forEach((link) => {
        link.remove();
      });

      convertedFiles.forEach((sheetName) => {
        const jsonContent = JSON.stringify(
          jsonArray.find((json) => json.sheetName === sheetName).data,
          null,
          2
        );
        const blob = new Blob([jsonContent], { type: "application/json" });
        const downloadLink = URL.createObjectURL(blob);
        if (downloadLink) {
          const anchor = document.createElement("a");
          anchor.href = downloadLink;
          anchor.download = `${sheetName}.json`;
          anchor.textContent = `${sheetName}.json`;
          anchor.style.display = "block";

          document.getElementById("downloadLinksContainer").appendChild(anchor);
        }
      });

      // downloadLinks.forEach((link) => {
      //   const anchor = document.createElement("a");
      //   anchor.href = link.downloadLink;
      //   anchor.download = `${link.sheetName}.json`;
      //   anchor.textContent = `${link.sheetName}.json`;
      //   anchor.style.display = "block";

      //   document.body.appendChild(anchor);

      //   // Trigger the click event to initiate the download
      //   anchor.click();

      //   // Remove the dynamically created anchor element
      //   document.body.removeChild(anchor);
      // });

      // jsonArray.forEach((jsonSheetData) => {
      //   const sheetName = jsonSheetData.sheetName;
      //   const sheetData = jsonSheetData.data;

      //   const jsonContent = JSON.stringify(sheetData, null, 2);
      //   const blob = new Blob([jsonContent], { type: "application/json" });
      //   const fileName = `${sheetName}.json`;
      //   saveAs(blob, fileName);

      //   // // Add the downloaded file name to the downloadedFiles array
      //   downloadedFiles.push(fileName);
      // });
    });
  };

  reader.readAsArrayBuffer(file);
}

//////////////////////

function updateDownloadButton() {
  const checkboxes = document.querySelectorAll(
    "#sheetCheckboxContainer input[type='checkbox']"
  );
  const selectedSheets = Array.from(checkboxes).filter(
    (checkbox) => checkbox.checked
  );

  downloadBtn.disabled = selectedSheets.length === 0;
}

function hasMathFormula(value, allowedFormulas) {
  // const formulaRegex = /([A-Z]+)\(/g;
  const formulaRegex = /^=/; // Regex to match math formulas
  let match;
  while ((match = formulaRegex.exec(value))) {
    const formula = match[1];
    if (!allowedFormulas.includes(formula)) {
      return true;
    }
  }
  return false;
}
