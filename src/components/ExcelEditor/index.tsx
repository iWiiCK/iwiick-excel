import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

interface ExcelDataRow {
  [key: string]: string | number | boolean | Date;
}

interface Props {
  url: string;
}

const ExcelEditor: React.FC<Props> = ({ url }: Props) => {
  const [excelData, setExcelData] = useState<ExcelDataRow[]>([]);

  useEffect(() => {
    const fetchExcelFile = async () => {
      try {
        // Fetch the file
        const response = await fetch(url);
        if (!response.ok) throw new Error("Error fetching the file");
        const blob = await response.blob();

        // Read the file as an ArrayBuffer
        const reader = new FileReader();
        reader.onload = (e: ProgressEvent<FileReader>) => {
          const result = e.target?.result;

          if (result) {
            let data: Uint8Array;

            if (result instanceof ArrayBuffer) {
              data = new Uint8Array(result); // If it's an ArrayBuffer, use it directly
            } else if (typeof result === "string") {
              // If it's a string, encode it to a Uint8Array
              data = new Uint8Array(new TextEncoder().encode(result));
            } else {
              // Default to an empty ArrayBuffer if result is null
              data = new Uint8Array(new ArrayBuffer(0));
            }

            const workbook = XLSX.read(data, { type: "array" });

            // Convert the first sheet to JSON
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            setExcelData(jsonData as ExcelDataRow[]);
          }
        };
        reader.readAsArrayBuffer(blob);
      } catch (error) {
        console.error("Error loading Excel file:", error);
      }
    };

    fetchExcelFile();
  }, [url]);

  const renderCell = (cell: string | number | boolean | Date) => {
    // Convert Date to string if it's a Date object
    if (cell instanceof Date) {
      return cell.toLocaleString(); // or use cell.toISOString() depending on your format preference
    }
    return cell;
  };

  return (
    <div style={{ padding: "20px" }}>
      <h1>Excel Viewer</h1>
      {excelData.length > 0 ? (
        <table border={1} style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead>
            <tr>
              {Object.keys(excelData[0]).map((key) => (
                <th
                  key={key}
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    backgroundColor: "#f2f2f2",
                    color: "black",
                  }}
                >
                  {key}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {excelData.map((row, index) => (
              <tr key={index}>
                {Object.values(row).map((cell, cellIndex) => (
                  <td key={cellIndex} style={{ padding: "8px" }}>
                    {renderCell(cell)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        <p>Loading Excel data...</p>
      )}
    </div>
  );
};

export default ExcelEditor;
