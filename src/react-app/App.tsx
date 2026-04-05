import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import "./App.css";

interface KPIData {
  labor: number;
  laborMarkup: number;
  materialsMarkup: number;
  landscaping: number;
  annualServices: number;
}

interface OperatorHours {
  [operator: string]: number;
}

interface ReportData {
  kpis: KPIData;
  operatorHours: OperatorHours;
  unassignedHours: number;
  totalWorkOrders: number;
}

function App() {
  const [reportData, setReportData] = useState<ReportData | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const processExcelFile = useCallback((file: File) => {
    setError(null);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: "array" });
        
        // Find the sheet with actual data (skip sheets with only 1 cell like "Generated: ...")
        let worksheet = null;
        let sheetName = "";
        for (const name of workbook.SheetNames) {
          const sheet = workbook.Sheets[name];
          const range = sheet["!ref"];
          if (range) {
            const decoded = XLSX.utils.decode_range(range);
            // Use sheet if it has more than 1 column
            if (decoded.e.c > 1) {
              worksheet = sheet;
              sheetName = name;
              break;
            }
          }
        }
        
        if (!worksheet) {
          worksheet = workbook.Sheets[workbook.SheetNames[0]];
          sheetName = workbook.SheetNames[0];
        }

        console.log("Using sheet:", sheetName);

        // Convert to JSON with header row
        const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
          defval: "",
        });

        console.log("Total rows:", jsonData.length);
        if (jsonData.length > 0) {
          console.log("Sample row keys:", Object.keys(jsonData[0]));
          console.log("Sample row:", jsonData[0]);
        }

        // Initialize KPIs
        const kpis: KPIData = {
          labor: 0,
          laborMarkup: 0,
          materialsMarkup: 0,
          landscaping: 0,
          annualServices: 0,
        };

        const operatorHours: OperatorHours = {};
        let unassignedHours = 0;

        // Helper to parse currency/number values
        const parseNumber = (val: unknown): number => {
          if (val === null || val === undefined || val === "" || val === " ") return 0;
          const str = String(val).replace(/[$,]/g, "").trim();
          if (str === "" || str === " ") return 0;
          const num = parseFloat(str);
          return isNaN(num) ? 0 : num;
        };

        // Helper to parse hours in "HH:MM" format or decimal
        const parseHours = (val: unknown): number => {
          if (val === null || val === undefined || val === "" || val === " ") return 0;
          const str = String(val).trim();
          if (str === "" || str === " ") return 0;
          
          // Check if it's in HH:MM format
          if (str.includes(":")) {
            const parts = str.split(":");
            const hours = parseInt(parts[0], 10) || 0;
            const minutes = parseInt(parts[1], 10) || 0;
            return hours + minutes / 60;
          }
          
          // Otherwise try to parse as decimal
          const num = parseFloat(str);
          return isNaN(num) ? 0 : num;
        };

        // Process each row
        jsonData.forEach((row) => {
          // Based on actual Excel structure:
          // Column A (0): Assigned Operator
          // Column R (17): Hours (format "HH:MM")
          // Column X (23): Labor
          // Column Y (24): Labor Markup
          // Column AA (26): Materials Markup
          // Column AD (29): Landscaping
          // Column AE (30): Spring/Winter Services (Annual Services)

          const laborValue = parseNumber(row["Labor"]);
          const laborMarkupValue = parseNumber(row["Labor Markup"]);
          const materialsMarkupValue = parseNumber(row["Materials Markup"]);
          const landscapingValue = parseNumber(row["Landscaping"]);
          const annualServicesValue = parseNumber(row["Spring/Winter Services"]);
          
          const hours = parseHours(row["Hours"]);
          const operator = String(row["Assigned Operator"] ?? "").trim();

          kpis.labor += laborValue;
          kpis.laborMarkup += laborMarkupValue;
          kpis.materialsMarkup += materialsMarkupValue;
          kpis.landscaping += landscapingValue;
          kpis.annualServices += annualServicesValue;

          if (operator && operator !== "0" && operator !== "undefined" && operator !== "") {
            operatorHours[operator] = (operatorHours[operator] || 0) + hours;
          } else {
            unassignedHours += hours;
          }
        });

        console.log("KPIs:", kpis);
        console.log("Operator Hours:", operatorHours);

        setReportData({
          kpis,
          operatorHours,
          unassignedHours,
          totalWorkOrders: jsonData.length,
        });
      } catch (err) {
        console.error("Error processing file:", err);
        setError("Error processing file. Please ensure it's a valid Excel file.");
      }
    };

    reader.onerror = () => {
      setError("Error reading file.");
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setIsDragging(false);

      const files = e.dataTransfer.files;
      if (files.length > 0) {
        const file = files[0];
        if (
          file.name.endsWith(".xls") ||
          file.name.endsWith(".xlsx") ||
          file.name.endsWith(".csv")
        ) {
          processExcelFile(file);
        } else {
          setError("Please upload an Excel file (.xls, .xlsx) or CSV file.");
        }
      }
    },
    [processExcelFile]
  );

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleFileSelect = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (files && files.length > 0) {
        processExcelFile(files[0]);
      }
    },
    [processExcelFile]
  );

  const formatCurrency = (value: number): string => {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
    }).format(value);
  };

  const formatHours = (value: number): string => {
    return value.toFixed(2);
  };

  const resetReport = () => {
    setReportData(null);
    setFileName(null);
    setError(null);
  };

  // Sort operators by hours (descending)
  const sortedOperators = reportData
    ? Object.entries(reportData.operatorHours).sort(([, a], [, b]) => b - a)
    : [];

  const totalOperatorHours = sortedOperators.reduce((sum, [, hours]) => sum + hours, 0);

  return (
    <div className="app-container">
      <header className="header">
        <h1>🔧 Maintenance Work Order Report</h1>
        <p className="subtitle">Drop your work order report to analyze KPIs</p>
      </header>

      {!reportData ? (
        <div
          className={`drop-zone ${isDragging ? "dragging" : ""}`}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
        >
          <div className="drop-content">
            <div className="drop-icon">📁</div>
            <h2>Drop Work Order Report Here</h2>
            <p>or click to select a file</p>
            <input
              type="file"
              accept=".xls,.xlsx,.csv"
              onChange={handleFileSelect}
              className="file-input"
            />
          </div>
        </div>
      ) : (
        <div className="results-container">
          <div className="file-info">
            <span className="file-name">📄 {fileName}</span>
            <button onClick={resetReport} className="reset-button">
              Upload New Report
            </button>
          </div>

          <div className="stats-summary">
            <span>Total Work Orders: <strong>{reportData.totalWorkOrders}</strong></span>
          </div>

          <section className="kpi-section">
            <h2>💰 Financial KPIs</h2>
            <div className="kpi-grid">
              <div className="kpi-card labor">
                <div className="kpi-icon">👷</div>
                <div className="kpi-label">Labor</div>
                <div className="kpi-value">{formatCurrency(reportData.kpis.labor)}</div>
              </div>
              <div className="kpi-card labor-markup">
                <div className="kpi-icon">📈</div>
                <div className="kpi-label">Labor Markup</div>
                <div className="kpi-value">{formatCurrency(reportData.kpis.laborMarkup)}</div>
              </div>
              <div className="kpi-card materials">
                <div className="kpi-icon">🔩</div>
                <div className="kpi-label">Materials Markup</div>
                <div className="kpi-value">{formatCurrency(reportData.kpis.materialsMarkup)}</div>
              </div>
              <div className="kpi-card landscaping">
                <div className="kpi-icon">🌿</div>
                <div className="kpi-label">Landscaping</div>
                <div className="kpi-value">{formatCurrency(reportData.kpis.landscaping)}</div>
              </div>
              <div className="kpi-card annual">
                <div className="kpi-icon">📅</div>
                <div className="kpi-label">Spring/Winter Services</div>
                <div className="kpi-value">{formatCurrency(reportData.kpis.annualServices)}</div>
              </div>
            </div>
          </section>

          <section className="operator-section">
            <h2>⏱️ Hours by Operator</h2>
            <div className="operator-stats">
              <span>Total Tracked Hours: <strong>{formatHours(totalOperatorHours + reportData.unassignedHours)}</strong></span>
            </div>
            
            <div className="operator-grid">
              {sortedOperators.map(([operator, hours]) => (
                <div key={operator} className="operator-card">
                  <div className="operator-name">{operator}</div>
                  <div className="operator-hours">{formatHours(hours)} hrs</div>
                  <div className="operator-bar">
                    <div 
                      className="operator-bar-fill" 
                      style={{ 
                        width: `${totalOperatorHours > 0 ? (hours / totalOperatorHours) * 100 : 0}%` 
                      }}
                    />
                  </div>
                </div>
              ))}
              
              {reportData.unassignedHours > 0 && (
                <div className="operator-card unassigned">
                  <div className="operator-name">⚠️ Unassigned</div>
                  <div className="operator-hours">{formatHours(reportData.unassignedHours)} hrs</div>
                  <div className="operator-bar">
                    <div 
                      className="operator-bar-fill unassigned-bar" 
                      style={{ 
                        width: `${totalOperatorHours + reportData.unassignedHours > 0 
                          ? (reportData.unassignedHours / (totalOperatorHours + reportData.unassignedHours)) * 100 
                          : 0}%` 
                      }}
                    />
                  </div>
                </div>
              )}

              {sortedOperators.length === 0 && reportData.unassignedHours === 0 && (
                <div className="no-data">No operator hours data found in the report.</div>
              )}
            </div>
          </section>
        </div>
      )}

      {error && (
        <div className="error-message">
          <span>⚠️ {error}</span>
          <button onClick={() => setError(null)}>✕</button>
        </div>
      )}
    </div>
  );
}

export default App;
