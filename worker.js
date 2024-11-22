
importScripts("https://cdn.sheetjs.com/xlsx-0.19.1/package/dist/xlsx.full.min.js");

onmessage = async function(event) {
  const sheetNames = ["Company Details", "Financial Info", "Executive"];
  const combinedSheets = {
    "Company Details": [],
    "Financial Info": [],
    "Executive": [],
  };

  const files = event.data;

  const processFile = async (file) => {
    try {
      const fileBuffer = await file.buffer;
      const workbook = XLSX.read(fileBuffer, { type: "array" });

      sheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (sheet) {
          const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          if(sheetData.length) {
            if (combinedSheets[sheetName].length === 0) {
              combinedSheets[sheetName].push(sheetData[0]);
            }
            const dataRows = sheetData.slice(1).filter(row => row.some(cell => cell !== undefined && cell !== null));
            combinedSheets[sheetName].push(...dataRows);
          }
        }
      });

      postMessage({ type: 'progress', message: `Processed ${file.name}` });

    } catch (error) {
      postMessage({ type: 'error', file: file.name, error: error.message });
    }
  };

  // Process each file in sequence
  for (let i = 0; i < files.length; i++) {
    await processFile(files[i]);
  }
  
  postMessage({ type: 'complete', combinedSheets });
};