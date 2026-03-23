import React, {useState} from "react";
import {Button, Loading} from "@apitable/components";
import {useDatasheet, useActiveViewId, useFields, useRecords} from "@apitable/widget-sdk";
import * as XLSX from "xlsx";

type ImportMode = "append" | "update";
type Step = "upload" | "config" | "preview" | "result";

export const ExcelImport: React.FC = () => {
  const viewId = useActiveViewId();
  const datasheet = useDatasheet();
  const fields = useFields(viewId);
  const records = useRecords(viewId);
  
  const [progressState, setProgressState] = useState<boolean>(false);
  const fileInput = React.createRef<HTMLInputElement>();

  const [step, setStep] = useState<Step>("upload");
  const [excelData, setExcelData] = useState<any[]>([]);
  const [excelHeader, setExcelHeader] = useState<string[]>([]);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);

  const [importMode, setImportMode] = useState<ImportMode>("append");
  const [allowIncrement, setAllowIncrement] = useState<boolean>(true);
  const [primaryKey, setPrimaryKey] = useState<string>("");

  const [planData, setPlanData] = useState<any[]>([]);
  const [resultSummary, setResultSummary] = useState({ add: 0, update: 0, skip: 0 });

  // Custom formatted dates
  function format(excelDate: any) {
    if (typeof excelDate === "number") {
      let step = new Date().getTimezoneOffset() <= 0 ? 25567 + 2 : 25567 + 1;
      let utc_days = Math.floor(excelDate - step);
      let utc_value = utc_days * 86400;
      let date_info = new Date(utc_value * 1000);
      let fractional_day = excelDate - Math.floor(excelDate) + 0.0000001;
      let total_seconds = Math.floor(86400 * fractional_day);
      let seconds = total_seconds % 60;
      total_seconds -= seconds;
      let hours = Math.floor(total_seconds / (60 * 60));
      let minutes = Math.floor(total_seconds / 60) % 60;
      return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds).getTime();
    } else if (typeof excelDate === "string") {
      excelDate = excelDate.substring(0, 19);
      excelDate = excelDate.replace(/-/g, "/");
      const timestamp = new Date(excelDate).getTime();
      return timestamp;
    } else {
      return null;
    }
  }

  const parseSheetData = (wb: XLSX.WorkBook, sheetName: string) => {
    const ws = wb.Sheets[sheetName];
    const data: any[] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
    if (data.length === 0) return { header: [], validData: [] };
    
    const header: any[] = data[0];
    const validData: any[] = data.slice(1).filter((s: any[]) => s && s.length !== 0);
    return { header, validData };
  };

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheet = e.target.value;
    setSelectedSheet(sheet);
    if (workbook) {
      const { header, validData } = parseSheetData(workbook, sheet);
      setExcelHeader(header);
      setExcelData(validData);
    }
  };

  function selectFile() {
    fileInput.current?.click();
  }

  const onImportExcel = (file: any) => {
    const {files} = file.target;
    if (!files || files.length === 0) return;
    const fileReader = new FileReader();
    fileReader.readAsBinaryString(files[0]);
    setProgressState(true);

    fileReader.onload = (event) => {
      try {
        const target = event.target?.result;
        const wb = XLSX.read(target, { type: "binary", cellText: false, cellDates: false });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);
        
        const wsname = wb.SheetNames[0];
        setSelectedSheet(wsname);
        
        const { header, validData } = parseSheetData(wb, wsname);
        setExcelHeader(header);
        setExcelData(validData);
        
        const fieldNames = fields.map(f => f.name);
        const intersection = fieldNames.filter(name => header.includes(name));
        
        if (intersection.length === 0) {
          alert("导入文件与当前表的列名没有交集");
          setProgressState(false);
          return;
        }

        if (validData.length === 0) {
          alert("导入文件数据为空");
          setProgressState(false);
          return;
        }
        
        if (intersection.length > 0) {
            setPrimaryKey(intersection[0]);
        }

        setProgressState(false);
        setStep("config");

      } catch (e) {
        console.error(e);
        alert(e);
        setProgressState(false);
      }
    };
    fileReader.onloadend = () => {
      if (fileInput.current) fileInput.current.value = "";
    };
  };

  const buildValuesMap = (recordRaw: any[]) => {
    const handleDateTimeType = (data: any) => format(data);
    const handleCheckboxType = (data: any) => (data === 1 || data === true || data === "true" || data === "yes" || data === "是");
    const handleMultiSelectType = (data: any) => String(data).split(",");
    const handleCurrencyType = (data: any) => typeof data != "number" ? Number(String(data).match(/-?[0-9]+(\.[0-9]+)?/)?.[0] || 0) : data;
    const handlePercentType = (data: any) => {
      if (typeof data === "number") return data;
      const match = String(data).match(/-?[0-9]+(\.[0-9]+)?/);
      if (!match) return 0;
      return String(data).includes("%") ? Number(match[0]) : Number(match[0]) * 100;
    };
    const handleNumberType = (data: any) => typeof data === "number" ? data : Number(String(data).match(/-?[0-9]+(\.[0-9]+)?/)?.[0] || 0);

    let fieldHandle: any = {
      DateTime: handleDateTimeType,
      Number: handleNumberType,
      Checkbox: handleCheckboxType,
      MultiSelect: handleMultiSelectType,
      Currency: handleCurrencyType,
      Percent: handlePercentType,
      Rating: handleNumberType,
    };

    const valuesMap: any = {};
    fields.forEach((field) => {
      const specialType = ["Attachment", "Member", "MagicLink", "CreatedTime", "LastModifiedTime", "CreatedBy", "LastModifiedBy", "AutoNumber"];
      if (field.isComputed || specialType.includes(field.type)) return;

      const index = excelHeader.indexOf(field.name);
      if (index === -1) return;

      try {
        let handleType = fieldHandle[field.type];
        let parseData = (recordRaw[index] === undefined || recordRaw[index] === null || String(recordRaw[index]) === "") 
          ? null 
          : (field.type in fieldHandle ? handleType(recordRaw[index]) : String(recordRaw[index]));
        
        if (typeof parseData === "number" && isNaN(parseData)) parseData = null;
        if (parseData !== null) {
          valuesMap[field.id] = parseData;
        }
      } catch (error) {
        valuesMap[field.id] = null;
      }
    });
    return valuesMap;
  };

  const calculatePlan = () => {
    if (importMode === "update" && !primaryKey) {
      alert("更新导入模式下，需选择匹配主键");
      return;
    }

    const primaryKeyField = fields.find(f => f.name === primaryKey);
    const primaryKeyId = primaryKeyField?.id;
    const primaryKeyIndex = excelHeader.indexOf(primaryKey);

    const planData = excelData.map((row) => {
      if (importMode === "append") {
        return { action: "add", row };
      } else {
        let matchId: string | undefined = undefined;
        if (primaryKeyId && primaryKeyIndex !== -1) {
          const excelKeyVal = String(row[primaryKeyIndex] ?? "").trim();
          const matchedRecord = records.find((r) => {
            const val = r.getCellValueString(primaryKeyId);
            return val !== null && String(val).trim() === excelKeyVal;
          });
          if (matchedRecord) {
            matchId = matchedRecord.id;
          }
        }
        
        if (matchId) {
          return { action: "update", row, matchId };
        } else {
          if (allowIncrement) {
             return { action: "add", row };
          } else {
             return { action: "skip", row };
          }
        }
      }
    });

    setPlanData(planData);
    setStep("preview");
  };

  const executeImport = () => {
    setProgressState(true);
    setStep("result");
    
    setTimeout(async () => {
        const adds = planData.filter(p => p.action === "add").map(p => ({ valuesMap: buildValuesMap(p.row) }));
        const updates = planData.filter(p => p.action === "update").map(p => ({ id: p.matchId, valuesMap: buildValuesMap(p.row) }));
        const skips = planData.filter(p => p.action === "skip");

        setResultSummary({ add: adds.length, update: updates.length, skip: skips.length });

        const chunk = (arr: any[], num: number) => {
           let ret: any[] = [];
           arr.forEach((item, i) => {
             if (i % num === 0) ret.push([]);
             ret[ret.length - 1].push(item);
           });
           return ret;
        };

        if (!datasheet) {
           setProgressState(false);
           return;
        }

        try {
            // Process Additions
            if (adds.length > 0) {
                const addChunks = chunk(adds, 1000);
                for (let i = 0; i < addChunks.length; i++) {
                   if (i > 0) await new Promise(res => setTimeout(res, 5000));
                   await datasheet.addRecords(addChunks[i]);
                }
            }

            // Process Updates
            if (updates.length > 0) {
                const updateChunks = chunk(updates, 1000);
                for (let i = 0; i < updateChunks.length; i++) {
                   if (i > 0) await new Promise(res => setTimeout(res, 5000));
                   // use updateRecordsAsync or modify the records
                   await datasheet.setRecords(updateChunks[i]);
                }
            }
        } catch(e) {
            console.error(e);
            alert("Error during execution: " + String(e));
        }

        setProgressState(false);
    }, 100);
  };

  const renderUpload = () => (
    <div>
      <div role="upload" onClick={selectFile}>
        <input type="file" ref={fileInput} accept=".xlsx, .xls, .csv" style={{display: "none"}} id="inputfile" onChange={onImportExcel} />
        <Button color="primary">选择Excel文件并上传</Button>
      </div>
      <p style={{paddingTop: "10px", fontSize: "12px", textAlign: "center", color: "GrayText"}}>仅支持 .xlsx .xls .csv</p>
    </div>
  );

  const containerStyle: React.CSSProperties = {
     display: "flex", flexDirection: "column", gap: "16px", padding: "12px"
  };
  
  const cardStyle: React.CSSProperties = {
     background: "#f9fafb", border: "1px solid #e5e7eb", borderRadius: "8px", padding: "16px"
  };

  const renderConfig = () => {
    return (
      <div style={containerStyle}>
        <div style={cardStyle}>
            <h3 style={{margin: "0 0 12px 0", fontSize: "14px"}}>配置导入方案</h3>
            
            {sheetNames.length > 1 && (
                <div style={{marginBottom: "12px"}}>
                    <label style={{display: "block", fontSize: "12px", marginBottom: "4px"}}>选择工作表(Sheet)</label>
                    <select value={selectedSheet} onChange={handleSheetChange} style={{width: "100%", padding: "6px", borderRadius: "4px", border: "1px solid #d1d5db"}}>
                        {sheetNames.map(name => <option key={name} value={name}>{name}</option>)}
                    </select>
                </div>
            )}

            <div style={{marginBottom: "12px"}}>
                <label style={{display: "block", fontSize: "12px", marginBottom: "4px"}}>导入模式</label>
                <div style={{display: "flex", gap: "16px", fontSize: "14px"}}>
                    <label style={{display: "flex", alignItems: "center", gap: "4px", cursor: "pointer"}}>
                        <input type="radio" checked={importMode === 'append'} onChange={() => setImportMode('append')} />
                        全量导入 (直接新增)
                    </label>
                    <label style={{display: "flex", alignItems: "center", gap: "4px", cursor: "pointer"}}>
                        <input type="radio" checked={importMode === 'update'} onChange={() => setImportMode('update')} />
                        更新导入 (匹配再更新)
                    </label>
                </div>
            </div>

            {importMode === 'update' && (
                <>
                    <div style={{marginBottom: "12px"}}>
                        <label style={{display: "block", fontSize: "12px", marginBottom: "4px"}}>匹配主键 (选择表格中和Excel中名称一致的列)</label>
                        <select value={primaryKey} onChange={e => setPrimaryKey(e.target.value)} style={{width: "100%", padding: "6px", borderRadius: "4px", border: "1px solid #d1d5db"}}>
                            {fields.filter(f => excelHeader.includes(f.name)).map(f => (
                                <option key={f.id} value={f.name}>{f.name}</option>
                            ))}
                        </select>
                    </div>
                    <div style={{marginBottom: "12px"}}>
                        <label style={{display: "flex", alignItems: "center", gap: "8px", fontSize: "14px", cursor: "pointer"}}>
                            <input type="checkbox" checked={allowIncrement} onChange={e => setAllowIncrement(e.target.checked)} />
                            是否允许增量添加条目 (遇到原数据表中没有的条目则新增)
                        </label>
                    </div>
                </>
            )}
            <Button color="primary" onClick={calculatePlan} style={{marginTop: "8px"}}>数据预处理</Button>
        </div>

        <div style={{...cardStyle, flex: 1, overflow: "hidden", display: "flex", flexDirection: "column"}}>
            <h3 style={{margin: "0 0 12px 0", fontSize: "14px"}}>Excel 数据预览 (前100条)</h3>
            <div style={{overflow: "auto", flex: 1, maxHeight: "300px", border: "1px solid #e5e7eb", borderRadius: "4px"}}>
                <table style={{width: "100%", borderCollapse: "collapse", fontSize: "12px"}}>
                    <thead style={{background: "#f3f4f6", position: "sticky", top: 0}}>
                        <tr>
                            {excelHeader.map(h => <th key={h} style={{padding: "8px", borderBottom: "1px solid #e5e7eb", textAlign: "left"}}>{h}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        {excelData.slice(0, 100).map((row, i) => (
                            <tr key={i} style={{borderBottom: "1px solid #f3f4f6"}}>
                                {excelHeader.map(h => <td key={h} style={{padding: "8px"}}>{String(row[excelHeader.indexOf(h)] ?? "")}</td>)}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
      </div>
    );
  };

  const renderPreview = () => {
    const getActionColor = (action: string) => {
        if (action === "add") return "#10b981"; // green
        if (action === "update") return "#3b82f6"; // blue
        if (action === "skip") return "#9ca3af"; // gray
        return "";
    };
    const getActionText = (action: string) => {
        if (action === "add") return "新增";
        if (action === "update") return "更新";
        if (action === "skip") return "跳过";
        return action;
    };

    return (
        <div style={containerStyle}>
             <div style={cardStyle}>
                 <h3 style={{margin: "0 0 12px 0", fontSize: "14px"}}>预处理执行方案</h3>
                 <p style={{fontSize: "12px", color: "GrayText", marginBottom: "16px"}}>
                    总览: <span style={{color: getActionColor('add')}}>新增 {planData.filter(p=>p.action==='add').length}</span> |&nbsp;
                    <span style={{color: getActionColor('update')}}>更新 {planData.filter(p=>p.action==='update').length}</span> |&nbsp;
                    <span style={{color: getActionColor('skip')}}>跳过 {planData.filter(p=>p.action==='skip').length}</span>
                 </p>
                 <div style={{display: "flex", gap: "12px"}}>
                    <Button onClick={() => setStep("config")}>返回配置</Button>
                    <Button color="primary" onClick={executeImport}>执行操作</Button>
                 </div>
             </div>

             <div style={{...cardStyle, flex: 1, overflow: "hidden", display: "flex", flexDirection: "column"}}>
                <h3 style={{margin: "0 0 12px 0", fontSize: "14px"}}>数据状态预览 (前100条)</h3>
                <div style={{overflow: "auto", flex: 1, maxHeight: "400px", border: "1px solid #e5e7eb", borderRadius: "4px"}}>
                    <table style={{width: "100%", borderCollapse: "collapse", fontSize: "12px"}}>
                        <thead style={{background: "#f3f4f6", position: "sticky", top: 0, zIndex: 1}}>
                            <tr>
                                <th style={{padding: "8px", borderBottom: "1px solid #e5e7eb", textAlign: "center", width: "60px"}}>状态</th>
                                {excelHeader.map(h => <th key={h} style={{padding: "8px", borderBottom: "1px solid #e5e7eb", textAlign: "left"}}>{h}</th>)}
                            </tr>
                        </thead>
                        <tbody>
                            {planData.slice(0, 100).map((plan, i) => (
                                <tr key={i} style={{borderBottom: "1px solid #f3f4f6", background: plan.action === "update" ? "#eff6ff" : "white"}}>
                                    <td style={{padding: "8px", textAlign: "center", color: getActionColor(plan.action), fontWeight: "bold"}}>
                                        {getActionText(plan.action)}
                                    </td>
                                    {excelHeader.map(h => {
                                        const isHighlight = plan.action === "update";
                                        return (
                                            <td key={h} style={{padding: "8px", color: isHighlight ? "#1e40af" : "inherit"}}>
                                                {String(plan.row[excelHeader.indexOf(h)] ?? "")}
                                            </td>
                                        );
                                    })}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
  };

  const renderResult = () => (
      <div style={containerStyle}>
          <div style={{textAlign: "center", padding: "32px", background: "#f9fafb", borderRadius: "8px", border: "1px solid #e5e7eb"}}>
             <div style={{fontSize: "48px", marginBottom: "16px"}}>🎉</div>
             <h2 style={{fontSize: "18px", margin: "0 0 16px 0"}}>导入完成！</h2>
             <div style={{fontSize: "14px", color: "#4b5563", marginBottom: "24px", lineHeight: "1.6"}}>
                 本次共执行：
                 <br />
                 <span style={{color: "#10b981", fontWeight: "bold"}}>新增记录: {resultSummary.add} 条</span>
                 <br />
                 <span style={{color: "#3b82f6", fontWeight: "bold"}}>更新记录: {resultSummary.update} 条</span>
                 <br />
                 <span style={{color: "#9ca3af", fontWeight: "bold"}}>跳过记录: {resultSummary.skip} 条</span>
             </div>
             <Button color="primary" onClick={() => {
                 setStep("upload");
                 setExcelData([]);
                 setExcelHeader([]);
                 setWorkbook(null);
             }}>再次导入</Button>
          </div>
      </div>
  );

  if (progressState) return <Loading />;
  
  return (
    <div style={{height: "100%", width: "100%", overflow: "hidden", display: "flex", flexDirection: "column"}}>
        {step === "upload" && renderUpload()}
        {step === "config" && renderConfig()}
        {step === "preview" && renderPreview()}
        {step === "result" && renderResult()}
    </div>
  );
};
