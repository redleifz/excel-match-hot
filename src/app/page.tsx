'use client';

import { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver'; // Make sure to install file-saver
import moment from 'moment';

export default function Home() {
  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);

  // useState to store parsed JSON data
  const [file1Data, setFile1Data] = useState<unknown[] | null>(null);
  const [file2Data, setFile2Data] = useState<unknown[] | null>(null);

  const [exportData, setExportData] = useState<unknown[] | null>(null);

  // Handle file upload for the first input
  const handleFile1Upload = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const uploadedFile = event.target.files ? event.target.files[0] : null;
    setFile1(uploadedFile);
  };

  // Handle file upload for the second input
  const handleFile2Upload = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const uploadedFile = event.target.files ? event.target.files[0] : null;
    setFile2(uploadedFile);
  };

  // Function to read and process an Excel file and store the result in the corresponding state
  const processFile = (file: File | null, setFileData: React.Dispatch<React.SetStateAction<unknown[] | null>>): void => {
    if (file) {
      const reader = new FileReader();

      reader.onload = (event: ProgressEvent<FileReader>): void => {
        const result = event.target?.result;
        if (!result) return;

        const data = new Uint8Array(result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // Combine all sheets into one array of data
        let allSheetData: unknown[] = [];
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
          const jsonData: unknown[] = XLSX.utils.sheet_to_json(worksheet);
          allSheetData = [...allSheetData, ...jsonData];
        });

        // Set the parsed data in the state
        setFileData(allSheetData);
        // console.log(`Processed data for ${file.name}:`, allSheetData);
      };

      reader.readAsArrayBuffer(file);
    }
  };

  // Handle submit to read both files and store data in state
  const handleSubmit = (): void => {
    if (file1) processFile(file1, setFile1Data);
    if (file2) processFile(file2, setFile2Data);

    // console.log('File 1 Data:', file1Data);
    // console.log('File 2 Data:', file2Data);
  };

  const excelSerialToJSDate = (serial: number): Date => {
    // Excel serial date starts on 1st January 1900, but it's offset by 1
    const daysSince1900 = serial - 25569; // Serial number days since 1970 (Unix epoch)
    const msPerDay = 86400000; // Number of milliseconds in a day
    return new Date(daysSince1900 * msPerDay); // Convert to Date object
  };

  const exportSubmit = (): void => {
    if (file1Data && file2Data) {
      const foundData: any[] = [];
      const notFoundData: any[] = [];
      const foundIDs = new Set<number | string>(); // Use a Set to track unique IDs
  
      // Process file2Data to separate into found and not found
      file2Data.forEach((file2Item) => {
        const matchingFile1Item = file1Data.find(
          (file1Item) => String(file1Item.ID) === String(file2Item.icode)
        );
  
        const formattedRxDate = moment(excelSerialToJSDate(file2Item.rxdate)).format('DD/MM/YYYY');
  
        if (matchingFile1Item) {
          const remain = Number(matchingFile1Item.Maximum) - Number(file2Item.SumOfSumOfqty) || `ไม่มีข้อมูล`;
  
          // Check if the ID is already in foundIDs Set
          if (!foundIDs.has(matchingFile1Item.ID)) {
            foundData.push({
              ID: matchingFile1Item.ID,
              NAME: matchingFile1Item.NAME,
              location: matchingFile1Item.location,
              Maximum: matchingFile1Item.Maximum ? Number(matchingFile1Item.Maximum) : `ไม่มีข้อมูล`,
              Minimum: matchingFile1Item.Minimum ? Number(matchingFile1Item.Minimum) : `ไม่มีข้อมูล`,
              SumOfSumOfqty: Number(file2Item.SumOfSumOfqty),
              remark: `${remain <= matchingFile1Item.Minimum ? `${remain}` : `เกิน ${remain}`}`,
            });
  
            // Add the ID to the foundIDs Set
            foundIDs.add(matchingFile1Item.ID);
          }
        } else {
          notFoundData.push({
            ...file2Item,
            rxdate: formattedRxDate, // Add formatted rxdate for not found items
          });
        }
      });
  
      // Create worksheets for found and not found data
      const foundWorksheet = XLSX.utils.json_to_sheet(foundData);
  
      // Create a new workbook and append the worksheets
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, foundWorksheet, 'Found');
  
      // Generate Excel file and trigger download
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const fileName = 'exported_data.xlsx';
  
      // Save the file using FileSaver
      const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
      saveAs(blob, fileName);
  
      setExportData(foundData); // Optional: set export data to include both found and not found
    } else {
      console.error('file1Data or file2Data is missing.');
    }
  };
  
  

  useEffect(() => {
    console.log(`data found in file1Data:`,exportData);
  },[exportData])

  return (
    <div className="flex h-screen items-center justify-center p-10">
      <div className="flex flex-col gap-2">
        <div className="text-xl text-center mb-10">Upload Excel Files</div>

        {/* First input for single file upload */}
        <div className="flex justify-center items-center gap-2">
          <span>Master Data:</span>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFile1Upload}
          />
        </div>

        {/* Second input for single file upload */}
        <div className="flex justify-center items-center gap-2">
          <span>HotXp Data:</span>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFile2Upload}
          />
        </div>

        {/* Submit button to process both files */}
        <button
          onClick={handleSubmit}
          className="mt-4 p-2 bg-blue-500 text-white rounded"
        >
          Read Files
        </button>

        {/* Display the number of rows in each file */}
        <div className="mt-4">
          <p>File 1 Data Rows: {file1Data?.length || 0}</p>
          <p>File 2 Data Rows: {file2Data?.length || 0}</p>
        </div>
        <button
          onClick={exportSubmit}
          className="mt-4 p-2 bg-green-500 text-white rounded"
        >
          Export Files
        </button>

      </div>
    </div>
  );
}
