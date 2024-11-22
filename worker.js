importScripts("https://cdn.sheetjs.com/xlsx-0.19.1/package/dist/xlsx.full.min.js");

onmessage = async function(event) {
  const fileData = event.data;
  const sheetNames = ["Company Details", "Financial Info", "Executive"];
  const combinedSheets = {
    "Company Details": [],
    "Financial Info": [],
    "Executive": [],
  };

  const totalFiles = fileData.length;
  for (let i = 0; i < totalFiles; i++) {
    const file = fileData[i];
    try {
      const fileBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(fileBuffer, { type: "array" });
      sheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        if (sheet) {
          const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          if (sheetData.length > 0) {
            if (combinedSheets[sheetName].length === 0) {
              combinedSheets[sheetName].push(sheetData[0]); // Add header row
            }
            const dataRows = sheetData.slice(1).filter(row => row.some(cell => cell !== undefined && cell !== null));
            combinedSheets[sheetName].push(...dataRows); // Add non-empty data rows
          }
        }
      });
      // Send progress update
      const percentage = Math.round(((i + 1) / totalFiles) * 100);
      postMessage({ type: 'progress', message: `Processed ${file.name} (${percentage}%)` });
    } catch (error) {
      postMessage({ type: 'error', file: file.name, error: error.message });
    }
  }
  
  postMessage({ type: 'complete', combinedSheets });
};