import React from "react";

// Use xlsx from CDN (no import)
const XLSX = window.XLSX;

export default function Test() {
  const handleExport = () => {
    if (!XLSX) {
      alert("XLSX library not loaded. Check your index.html script.");
      return;
    }
    const data = [{ Name: "Test", Value: 123 }];
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "test.xlsx");
  };

  return (
    <div>
      <h1>XLSX Test</h1>
      <button onClick={handleExport}>Export Excel</button>
    </div>
  );
}