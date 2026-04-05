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

interface OperatorTimeData {
  trackedHours: number;
  untrackedHours: number; // Minimum 15 min for blank time entries
}

interface OperatorHoursMap {
  [operator: string]: OperatorTimeData;
}

interface ReportData {
  kpis: KPIData;
  operatorHours: OperatorHoursMap;
  unassignedTrackedHours: number;
  unassignedUntrackedHours: number;
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

        const operatorHours: OperatorHoursMap = {};
        let unassignedTrackedHours = 0;
        let unassignedUntrackedHours = 0;

        // Minimum hours for blank time entries (15 minutes)
        const MIN_HOURS = 0.25;

        // Helper to parse currency/number values
        const parseNumber = (val: unknown): number => {
          if (val === null || val === undefined || val === "" || val === " ") return 0;
          const str = String(val).replace(/[$,]/g, "").trim();
          if (str === "" || str === " ") return 0;
          const num = parseFloat(str);
          return isNaN(num) ? 0 : num;
        };

        // Helper to parse hours in "HH:MM" format or decimal
        // Returns { hours: number, isBlank: boolean }
        const parseHours = (val: unknown): { hours: number; isBlank: boolean } => {
          if (val === null || val === undefined || val === "" || val === " ") {
            return { hours: 0, isBlank: true };
          }
          const str = String(val).trim();
          if (str === "" || str === " ") {
            return { hours: 0, isBlank: true };
          }
          
          // Check if it's in HH:MM format
          if (str.includes(":")) {
            const parts = str.split(":");
            const hours = parseInt(parts[0], 10) || 0;
            const minutes = parseInt(parts[1], 10) || 0;
            const totalHours = hours + minutes / 60;
            // If time is 00:00, treat as blank
            if (totalHours === 0) {
              return { hours: 0, isBlank: true };
            }
            return { hours: totalHours, isBlank: false };
          }
          
          // Otherwise try to parse as decimal
          const num = parseFloat(str);
          if (isNaN(num) || num === 0) {
            return { hours: 0, isBlank: true };
          }
          return { hours: num, isBlank: false };
        };

        // Helper to check if a row is the grand total row
        const isGrandTotalRow = (row: Record<string, unknown>): boolean => {
          const operator = String(row["Assigned Operator"] ?? "").trim().toLowerCase();
          // Check if the operator field contains "total" or "grand total"
          if (operator.includes("total") || operator.includes("grand")) {
            return true;
          }
          // Also check if this might be a summary row by checking if Work Order # is empty but has values
          const workOrderNum = row["Work Order #"];
          const hasLabor = parseNumber(row["Labor"]) > 0;
          const hasNoWorkOrder = !workOrderNum || workOrderNum === "" || workOrderNum === 0;
          // If no work order number but has labor value, and operator is empty - likely a total row
          if (hasNoWorkOrder && hasLabor && !operator) {
            return true;
          }
          return false;
        };

        // Process each row (skip the last row which is the grand total)
        const dataRows = jsonData.slice(0, -1); // Remove last row (grand total)
        
        dataRows.forEach((row) => {
          // Skip any row that looks like a grand total
          if (isGrandTotalRow(row)) {
            console.log("Skipping total row:", row);
            return;
          }

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
          
          const { hours, isBlank } = parseHours(row["Hours"]);
          const operator = String(row["Assigned Operator"] ?? "").trim();

          kpis.labor += laborValue;
          kpis.laborMarkup += laborMarkupValue;
          kpis.materialsMarkup += materialsMarkupValue;
          kpis.landscaping += landscapingValue;
          kpis.annualServices += annualServicesValue;

          if (operator && operator !== "0" && operator !== "undefined" && operator !== "") {
            // Initialize operator if not exists
            if (!operatorHours[operator]) {
              operatorHours[operator] = { trackedHours: 0, untrackedHours: 0 };
            }
            
            if (isBlank) {
              // Add minimum 15 minutes for blank time
              operatorHours[operator].untrackedHours += MIN_HOURS;
            } else {
              operatorHours[operator].trackedHours += hours;
            }
          } else {
            if (isBlank) {
              unassignedUntrackedHours += MIN_HOURS;
            } else {
              unassignedTrackedHours += hours;
            }
          }
        });

        console.log("KPIs:", kpis);
        console.log("Operator Hours:", operatorHours);
        console.log("Data rows processed:", dataRows.length);

        setReportData({
          kpis,
          operatorHours,
          unassignedTrackedHours,
          unassignedUntrackedHours,
          totalWorkOrders: dataRows.length,
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

  // Sort operators by total hours (tracked + untracked) descending
  const sortedOperators = reportData
    ? Object.entries(reportData.operatorHours).sort(
        ([, a], [, b]) => (b.trackedHours + b.untrackedHours) - (a.trackedHours + a.untrackedHours)
      )
    : [];

  // Calculate totals
  const totalTrackedHours = sortedOperators.reduce((sum, [, data]) => sum + data.trackedHours, 0);
  const totalUntrackedHours = sortedOperators.reduce((sum, [, data]) => sum + data.untrackedHours, 0);
  const grandTotalTracked = totalTrackedHours + (reportData?.unassignedTrackedHours ?? 0);
  const grandTotalUntracked = totalUntrackedHours + (reportData?.unassignedUntrackedHours ?? 0);

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
              <div className="stats-row">
                <span>Total Tracked Hours: <strong className="tracked">{formatHours(grandTotalTracked)}</strong></span>
                <span>Total Untracked (Blank): <strong className="untracked">{formatHours(grandTotalUntracked)}</strong></span>
                <span>Combined Total: <strong>{formatHours(grandTotalTracked + grandTotalUntracked)}</strong></span>
              </div>
            </div>
            
            <div className="operator-grid">
              {sortedOperators.map(([operator, data]) => {
                const totalHours = data.trackedHours + data.untrackedHours;
                const maxTotal = grandTotalTracked + grandTotalUntracked;
                return (
                  <div key={operator} className="operator-card">
                    <div className="operator-name">{operator}</div>
                    <div className="operator-hours-container">
                      <div className="hours-breakdown">
                        <span className="hours-tracked" title="Tracked hours">
                          ✓ {formatHours(data.trackedHours)} hrs
                        </span>
                        <span className="hours-untracked" title="Untracked (blank) hours - 15 min minimum">
                          ○ {formatHours(data.untrackedHours)} hrs
                        </span>
                      </div>
                      <div className="hours-total">
                        Total: {formatHours(totalHours)} hrs
                      </div>
                    </div>
                    <div className="operator-bar">
                      <div 
                        className="operator-bar-fill tracked-bar" 
                        style={{ 
                          width: `${maxTotal > 0 ? (data.trackedHours / maxTotal) * 100 : 0}%` 
                        }}
                      />
                      <div 
                        className="operator-bar-fill untracked-bar" 
                        style={{ 
                          width: `${maxTotal > 0 ? (data.untrackedHours / maxTotal) * 100 : 0}%`,
                          left: `${maxTotal > 0 ? (data.trackedHours / maxTotal) * 100 : 0}%`
                        }}
                      />
                    </div>
                  </div>
                );
              })}
              
              {(reportData.unassignedTrackedHours > 0 || reportData.unassignedUntrackedHours > 0) && (
                <div className="operator-card unassigned">
                  <div className="operator-name">⚠️ Unassigned</div>
                  <div className="operator-hours-container">
                    <div className="hours-breakdown">
                      <span className="hours-tracked" title="Tracked hours">
                        ✓ {formatHours(reportData.unassignedTrackedHours)} hrs
                      </span>
                      <span className="hours-untracked" title="Untracked (blank) hours - 15 min minimum">
                        ○ {formatHours(reportData.unassignedUntrackedHours)} hrs
                      </span>
                    </div>
                    <div className="hours-total">
                      Total: {formatHours(reportData.unassignedTrackedHours + reportData.unassignedUntrackedHours)} hrs
                    </div>
                  </div>
                  <div className="operator-bar">
                    <div 
                      className="operator-bar-fill unassigned-tracked-bar" 
                      style={{ 
                        width: `${grandTotalTracked + grandTotalUntracked > 0 
                          ? (reportData.unassignedTrackedHours / (grandTotalTracked + grandTotalUntracked)) * 100 
                          : 0}%` 
                      }}
                    />
                  </div>
                </div>
              )}

              {sortedOperators.length === 0 && reportData.unassignedTrackedHours === 0 && reportData.unassignedUntrackedHours === 0 && (
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
