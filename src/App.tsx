import React, { useState, useRef } from "react";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.mjs?url";
import * as ExcelJS from "exceljs";
import Fuse from "fuse.js";
import { UploadCloud, FileSpreadsheet, FileText, CheckCircle, RefreshCw, Download } from "lucide-react";
import "./App.css";

// Polyfill for Safari/WebKit missing async iteration on Streams
if (typeof ReadableStream !== "undefined" && !ReadableStream.prototype[Symbol.asyncIterator]) {
  (ReadableStream.prototype as any)[Symbol.asyncIterator] = async function* () {
    const reader = this.getReader();
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) return;
        yield value;
      }
    } finally {
      reader.releaseLock();
    }
  };
}

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const GRADE_REGEX = /\b(1\.00|1\.25|1\.50|1\.75|2\.00|2\.25|2\.50|2\.75|3\.00|4\.00|5\.00|1\.0|1\.5|2\.0|2\.5|3\.0|INC|DRP|AW|P|PASS|FAIL|S|U)\b/i;

function App() {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [pdfFiles, setPdfFiles] = useState<File[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState<string[]>([]);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const templateInputRef = useRef<HTMLInputElement>(null);
  const pdfsInputRef = useRef<HTMLInputElement>(null);

  const addLog = (msg: string) => setLogs((p) => [...p, msg]);

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setTemplateFile(e.target.files[0]);
      addLog(`Loaded template: ${e.target.files[0].name}`);
    }
  };

  const handlePdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const files = Array.from(e.target.files);
      setPdfFiles((prev) => [...prev, ...files]);
      addLog(`Loaded ${files.length} PDF file(s).`);
    }
  };

  const copyWorksheet = (workbook: ExcelJS.Workbook, originalSheet: ExcelJS.Worksheet, newSheetName: string) => {
    // Using deep clone of the model guarantees all formulas, styles, and merges are copied perfectly
    // without dropping calculation references or complex cells.
    const sheetModel = JSON.parse(JSON.stringify((originalSheet as any).model));
    sheetModel.name = newSheetName;
    
    const newSheet = workbook.addWorksheet(newSheetName);
    (newSheet as any).model = sheetModel;
    
    return newSheet;
  };

  const extractLinesFromPdf = async (file: File): Promise<{ lines: string[], extractedName: string | null }> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
    let extractedLines: string[] = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        // Items must be sorted by Y and X coordinate to construct coherent lines
        const items = textContent.items as any[];
        items.sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5];
            if (Math.abs(yDiff) < 5) return a.transform[4] - b.transform[4];
            return yDiff;
        });

        let currentY: number | null = null;
        let currentLine: string[] = [];
        
        for (let item of items) {
            const str = item.str.trim();
            if (!str) continue;

            if (currentY === null || Math.abs(item.transform[5] - currentY) > 5) {
                if (currentLine.length > 0) extractedLines.push(currentLine.join(" "));
                currentLine = [str];
                currentY = item.transform[5];
            } else {
                currentLine.push(str);
            }
        }
        if (currentLine.length > 0) extractedLines.push(currentLine.join(" "));
    }

    const fullText = extractedLines.join(" ");
    let extractedName = null;
    const nameMatch = fullText.match(/Name\s+(.*?)\s+Nationality/i);
    if (nameMatch && nameMatch[1]) {
        extractedName = nameMatch[1].trim();
    }

    return { lines: extractedLines, extractedName };
  };

  const startEncoding = async () => {
    if (!templateFile || pdfFiles.length === 0) return;
    setIsProcessing(true);
    setProgress(0);
    setLogs([]);
    setDownloadUrl(null);

    try {
      addLog("Starting extraction of PDFs...");
      const parsedPdfs = [];
      let currentProgress = 0;
      
      const progressPerPdf = 40 / pdfFiles.length; // 40% of time extracting

      for (let file of pdfFiles) {
        addLog(`Extracting text from: ${file.name}`);
        const { lines, extractedName } = await extractLinesFromPdf(file);
        parsedPdfs.push({
          file,
          lines,
          fullText: lines.join(" ").toLowerCase(),
          extractedName
        });
        currentProgress += progressPerPdf;
        setProgress(Math.round(currentProgress));
      }

      addLog("Parsing Excel template...");
      if (!templateFile.name.toLowerCase().endsWith(".xlsx")) {
        throw new Error("Invalid Excel format. Please ensure your template is an '.xlsx' file and not an older '.xls' file.");
      }
      const templateBuffer = await templateFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(templateBuffer);
      setProgress(50);

      const templateSheet = workbook.worksheets[0];
      if (!templateSheet) throw new Error("Template workbook has no sheets!");

      // Detect the dynamic 'Final Grade' column
      let gradeColIndex = -1;
      
      templateSheet.eachRow((row, rowNumber) => {
          row.eachCell((cell, colNumber) => {
              if (cell.type === ExcelJS.ValueType.String) {
                  const val = cell.value?.toString().toLowerCase().trim().replace(/\s+/g, " ");
                  if (val === "final grade" || val === "grade") {
                      gradeColIndex = colNumber;
                  }
              }
          });
      });

      if (gradeColIndex === -1) {
          addLog("⚠️ 'Final Grade' header not found strictly. Falling back to adjacent mapping, which may overwrite other fields.");
      } else {
          addLog(`Detected 'Final Grade' at column ${gradeColIndex}`);
      }

      const progressPerSheet = 40 / Math.max(1, parsedPdfs.length);

      for (let pdfData of parsedPdfs) {
        // Attempt to establish a clean tab name, limit Excel tab name to 31 chars
        let tabName = pdfData.extractedName || pdfData.file.name.replace(/\.pdf$/i, "");
        if (tabName.length > 31) tabName = tabName.substring(0, 31);
        
        addLog(`Duplicating template for student: ${tabName}`);
        
        const newSheet = copyWorksheet(workbook, templateSheet, tabName);

        // Put the extracted name into the Name field of the template explicitly replacing B8
        if (pdfData.extractedName) {
            newSheet.getCell('B8').value = pdfData.extractedName.toUpperCase();
        }

        const fuse = new Fuse(pdfData.lines, {
            includeScore: true,
            threshold: 0.35, 
            distance: 100,
        });

        let encodingsInTab = 0;

        newSheet.eachRow((row, _rowNumber) => {
            // Prevent automated scanning from breaking any headers by skipping rows 1-10
            if (_rowNumber <= 10) return;

            row.eachCell((cell, colNumber) => {
                // Ensure we only look at course codes which are on the left side (cols 1 to 5)
                if (colNumber > 5) return;

                if (cell.type === ExcelJS.ValueType.String) {
                    const val = cell.value?.toString().trim();
                    if (val && val.length > 2 && val.length < 50) { 
                        const normalizedval = val.toLowerCase().replace(/\s+/g, " ");
                        const ignoreList = ["final grade", "course code", "course title", "remarks", "credit", "contact hours", "pre requisite", "instructor", "lec", "lab"];
                        if (ignoreList.includes(normalizedval)) return;

                        // Check if the pdf fully includes the item first to bypass fuse score degrading on long lines
                        let matchedLine = pdfData.lines.find(l => l.toLowerCase().includes(val.toLowerCase()));
                        
                        if (!matchedLine) {
                            const results = fuse.search(val);
                            if (results.length > 0 && results[0].score !== undefined && results[0].score < 0.3) {
                                matchedLine = results[0].item;
                            }
                        }

                        if (matchedLine) {
                            const gradeMatch = matchedLine.match(GRADE_REGEX);
                            
                            if (gradeMatch) {
                                const targetCell = gradeColIndex !== -1 ? row.getCell(gradeColIndex) : row.getCell(colNumber + 1);
                                
                                if (!targetCell.value || targetCell.value === "") {
                                    targetCell.value = gradeMatch[0]; 
                                    encodingsInTab++;
                                }
                            }
                        }
                    }
                }
            });
        });
        
        addLog(`↳ Encoded ${encodingsInTab} grades for ${tabName}.`);
        
        currentProgress += progressPerSheet;
        setProgress(Math.round(currentProgress));
      }

      // Cleanup: remove the original empty template tab so it's not downloaded
      workbook.removeWorksheet(templateSheet.id);

      addLog("Generating output file...");
      const outputBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([outputBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setProgress(100);
      addLog("🎉 Encoding complete! File ready for download.");
    } catch (e: any) {
      console.error(e);
      addLog(`❌ ERROR: ${e.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="app-container">
      <header>
        <h1>Grade Encoder</h1>
        <p className="subtitle">Automated PDF to Excel grades processing</p>
      </header>

      <div className="dashboard">
        <div className="card">
          <h2><FileSpreadsheet className="icon" /> 1. Upload Template</h2>
          <div 
            className="dropzone" 
            onClick={() => templateInputRef.current?.click()}
          >
            <input 
              type="file" 
              accept=".xlsx" 
              ref={templateInputRef} 
              style={{ display: "none" }} 
              onChange={handleTemplateUpload}
            />
            <UploadCloud size={40} className="icon" />
            <p>{templateFile ? templateFile.name : "Click to select your Excel template"}</p>
          </div>
        </div>

        <div className="card">
          <h2><FileText className="icon" /> 2. Upload PDFs</h2>
          <div 
            className="dropzone" 
            onClick={() => pdfsInputRef.current?.click()}
          >
            <input 
              type="file" 
              accept=".pdf" 
              multiple 
              ref={pdfsInputRef} 
              style={{ display: "none" }} 
              onChange={handlePdfUpload}
            />
            <UploadCloud size={40} className="icon" />
            <p>Click to select one or more PDF files</p>
          </div>
          {pdfFiles.length > 0 && (
            <div className="file-list" style={{ maxHeight: "100px", overflowY: "auto" }}>
              {pdfFiles.map((f, i) => (
                <div key={i} className="file-item">
                  <span className="file-name">{f.name}</span>
                  <CheckCircle size={16} color="var(--secondary)" />
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="card full-width">
          <h2><RefreshCw className="icon" /> 3. Process</h2>
          
          <button 
            className="btn" 
            onClick={startEncoding} 
            disabled={isProcessing || !templateFile || pdfFiles.length === 0}
          >
            {isProcessing ? "Processing..." : "Start Encoding"}
          </button>

          {(progress > 0 || logs.length > 0) && (
            <div className="progress-container">
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.9rem" }}>
                <span>Status</span>
                <span>{progress}%</span>
              </div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${progress}%` }}></div>
              </div>
              
              <div className="status-log">
                {logs.map((log, i) => (
                  <p key={i}>{log}</p>
                ))}
              </div>
            </div>
          )}

          {downloadUrl && (
            <a 
              href={downloadUrl} 
              download="Encoded_Grades_Output.xlsx" 
              style={{ textDecoration: 'none' }}
            >
              <button className="btn" style={{ marginTop: "1rem", backgroundColor: "var(--secondary)" }}>
                <Download size={20} /> Download Encoded Excel
              </button>
            </a>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
