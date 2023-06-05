import React, { useState, useRef } from "react";
import Dropzone from "react-dropzone";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";

import "./index.css";

const tableStyles = {
  borderCollapse: "collapse",
  width: "100%",
};

const cellStyles = {
  border: "1px solid #CCC",
  lineWidth: 0.1,
  padding: "4px",
};

const inputStyles = {
  margin: "8px",
  padding: "4px",
};

const tableContainerStyles = {
  margin: "20px", // Add margin around the table container
  display: "inline-block", // Add display: inline-block to allow the margin to take effect
};

const App = () => {
  const [excelData, setExcelData] = useState([]);
  const [startColumn, setStartColumn] = useState("");
  const [endColumn, setEndColumn] = useState("");
  const [startRow, setStartRow] = useState("");
  const [endRow, setEndRow] = useState("");
  const [screenshots, setScreenshots] = useState([]);

  const containerRef = useRef(null);

  const handleFileUpload = (files) => {
    const fileReader = new FileReader();
    fileReader.onload = (e) => {
      const arrayBuffer = e.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      setExcelData(data);
    };
    fileReader.readAsArrayBuffer(files[0]);
  };

  const handleScreenshotCapture = async () => {
    const selectedData = excelData
      .slice(startRow - 1, endRow)
      .map((row) =>
        row.slice(startColumn.charCodeAt(0) - 65, endColumn.charCodeAt(0) - 64)
      );

    const container = containerRef.current;
    container.style.display = "block"; // Set display to "block" to make it visible

    // Delay added to allow the content to render
    await new Promise((resolve) => setTimeout(resolve, 0));

    // Remove any existing table from the container
    container.innerHTML = "";

    // Create a new table to contain only the selected data
    const table = document.createElement("table");
    table.style.borderCollapse = "collapse"; // Add border-collapse style to collapse borders
    const thead = document.createElement("thead");
    const tbody = document.createElement("tbody");

    selectedData.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      row.forEach((cell, columnIndex) => {
        const cellElement =
          rowIndex === 0
            ? document.createElement("th")
            : document.createElement("td");
        const cellStyle =
          rowIndex === 0
            ? {
                ...cellStyles,
                fontWeight: "bold",
                fontSize: "9px",
              }
            : { ...cellStyles, fontSize: "8px" };
        Object.assign(cellElement.style, cellStyle);
        cellElement.innerText = cell;
        tr.appendChild(cellElement);
      });

      if (rowIndex === 0) {
        thead.appendChild(tr);
      } else {
        tbody.appendChild(tr);
      }
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    container.appendChild(table);

    html2canvas(table, {
      scrollX: 0,
      scrollY: -window.scrollY,
      useCORS: true,
      allowTaint: true,
      scale: 2, // Adjust the scale to improve image quality
    }).then((canvas) => {
      // Apply styles to the captured canvas
      const context = canvas.getContext("2d");
      context.lineWidth = 0.01; // Adjust the line width to reduce the thickness of the borders

      const rows = Array.from(table.getElementsByTagName("tr"));
      const cells = rows.flatMap((row) =>
        Array.from(row.getElementsByTagName("td"))
      );

      cells.forEach((cell) => {
        const { left, top, width, height } = cell.getBoundingClientRect();
        context.strokeRect(left, top, width, height);
      });

      setScreenshots([canvas.toDataURL("image/jpeg")]);
      container.style.display = "none"; // Reset display to "none" after capturing the screenshot
      container.removeChild(table); // Remove the temporary table from the container
    });
  };

  const handleSaveScreenshots = () => {
    screenshots.forEach((screenshot, index) => {
      const link = document.createElement("a");
      link.href = screenshot;
      link.download = `screenshot-${index + 1}.jpeg`;
      link.click();
    });
  };

  return (
    <div className="App">
      <h1>Excel ScreenShot App</h1>
      <Dropzone onDrop={handleFileUpload}>
        {({ getRootProps, getInputProps }) => (
          <div {...getRootProps()} className="upload-button">
            <input {...getInputProps()} />
            <p className="upload-button-text">Upload File</p>
          </div>
        )}
      </Dropzone>
      {excelData.length > 0 && (
        <div>
          <label htmlFor="startColumn">Start Column:</label>
          <input
            type="text"
            id="startColumn"
            value={startColumn}
            onChange={(e) => setStartColumn(e.target.value.toUpperCase())}
            style={inputStyles}
          />
          <label htmlFor="endColumn">End Column:</label>
          <input
            type="text"
            id="endColumn"
            value={endColumn}
            onChange={(e) => setEndColumn(e.target.value.toUpperCase())}
            style={inputStyles}
          />
          <label htmlFor="startRow">Start Row:</label>
          <input
            type="number"
            id="startRow"
            value={startRow}
            onChange={(e) => setStartRow(e.target.value)}
            style={inputStyles}
          />
          <label htmlFor="endRow">End Row:</label>
          <input
            type="number"
            id="endRow"
            value={endRow}
            onChange={(e) => setEndRow(e.target.value)}
            style={inputStyles}
          />
          <button onClick={handleScreenshotCapture} style={inputStyles}>
            Take a Screenshot
          </button>
          {screenshots.length > 0 && (
            <div>
              {screenshots.map((screenshot, index) => (
                <div key={index}>
                  <img src={screenshot} alt={`Screenshot ${index + 1}`} />
                </div>
              ))}
              <button onClick={handleSaveScreenshots} style={inputStyles}>
                Download Image
              </button>
            </div>
          )}
        </div>
      )}
      <div style={tableContainerStyles}>
        <div ref={containerRef}>
          <table style={{ ...tableStyles, display: "none" }} ref={containerRef}>
            <tbody>
              {excelData.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, columnIndex) => (
                    <td
                      key={`${rowIndex}-${columnIndex}`}
                      style={{
                        ...cellStyles,
                        borderTop: rowIndex === 0 ? "0.1px solid #000" : "none", // Only add top border for the first row
                      }}
                    >
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default App;
