import React, { useState, useMemo, useCallback, useEffect } from 'react';
import { 
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine, ReferenceArea,
  AreaChart, Area, Brush
} from 'recharts';
import { 
  Activity, Upload, Info, Crosshair, 
  BarChart, ShieldCheck, Layers, Waves,
  Home, ArrowLeft, ArrowUpRight, ArrowRight, Music, 
  ListMusic, Database, FolderOpen, PlaySquare,
  Save, Edit2, Trash2, Download, FileSpreadsheet,
  Settings2, Link, Eye, Zap, Target
} from 'lucide-react';

// --- 全域常數與設定 ---
const MUSCLE_LIST = [
  'UT', 'LT', 'SA', 'PM', 'LD', 'BB', 'TB'
];

// --- 動態載入 Excel (SheetJS) 函式庫 ---
const loadXLSX = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('無法載入 Excel 匯出模組，請檢查網路連線'));
    document.head.appendChild(script);
  });
};

// --- 數位信號處理 (DSP) 與數學工具函數 (共用) ---
const calcMean = (arr) => {
  if (!arr || arr.length === 0) return 0;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
};

const calcSD = (arr, mean) => {
  if (!arr || arr.length < 2) return 0;
  const variance = arr.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (arr.length - 1);
  return Math.sqrt(variance);
};

// 實作 2階 Butterworth Biquad 濾波器 (單向)
const biquadFilter = (data, type, cutoff, sampleRate) => {
  const omega = 2 * Math.PI * cutoff / sampleRate;
  const alpha = Math.sin(omega) / (2 * 0.7071); 
  const cosW = Math.cos(omega);

  let b0, b1, b2, a0, a1, a2;

  if (type === 'lowpass') {
    b0 = (1 - cosW) / 2;
    b1 = 1 - cosW;
    b2 = (1 - cosW) / 2;
    a0 = 1 + alpha;
    a1 = -2 * cosW;
    a2 = 1 - alpha;
  } else if (type === 'highpass') {
    b0 = (1 + cosW) / 2;
    b1 = -(1 + cosW);
    b2 = (1 + cosW) / 2;
    a0 = 1 + alpha;
    a1 = -2 * cosW;
    a2 = 1 - alpha;
  } else {
    return data;
  }

  b0 /= a0; b1 /= a0; b2 /= a0; a1 /= a0; a2 /= a0;

  const output = new Float64Array(data.length);
  let x1 = 0, x2 = 0, y1 = 0, y2 = 0;

  for (let i = 0; i < data.length; i++) {
    const x0 = data[i];
    const y0 = b0 * x0 + b1 * x1 + b2 * x2 - a1 * y1 - a2 * y2;
    output[i] = y0;
    x2 = x1; x1 = x0;
    y2 = y1; y1 = y0;
  }
  return output;
};

// 零相位位移濾波器 (Zero-Lag / Forward-Backward Filter)
const zeroLagBiquadFilter = (data, type, cutoff, sampleRate) => {
  const forward = biquadFilter(data, type, cutoff, sampleRate);
  const reversed = new Float64Array(forward.length);
  for (let i = 0; i < forward.length; i++) {
    reversed[i] = forward[forward.length - 1 - i];
  }
  const backward = biquadFilter(reversed, type, cutoff, sampleRate);
  const finalOut = new Float64Array(backward.length);
  for (let i = 0; i < backward.length; i++) {
    finalOut[i] = backward[backward.length - 1 - i];
  }
  return finalOut;
};

// 帶通濾波器
const bandpassFilter = (data, hpCutoff = 20, lpCutoff = 450, sampleRate = 1000) => {
  const hpFiltered = zeroLagBiquadFilter(data, 'highpass', hpCutoff, sampleRate);
  return zeroLagBiquadFilter(hpFiltered, 'lowpass', lpCutoff, sampleRate);
};

const calculateRMS = (data, windowSizeSamples) => {
  const n = data.length;
  const rmsData = new Float64Array(n);
  const halfWindow = Math.floor(windowSizeSamples / 2);

  let currentSumSq = 0;
  let count = 0;

  for (let i = 0; i <= halfWindow && i < n; i++) {
    currentSumSq += data[i] * data[i];
    count++;
  }
  if (count > 0) rmsData[0] = Math.sqrt(Math.max(0, currentSumSq) / count);

  for (let i = 1; i < n; i++) {
    const addedIndex = i + halfWindow;
    const removedIndex = i - halfWindow - 1;

    if (addedIndex < n) {
      currentSumSq += data[addedIndex] * data[addedIndex];
      count++;
    }
    if (removedIndex >= 0) {
      currentSumSq -= data[removedIndex] * data[removedIndex];
      count--;
    }
    rmsData[i] = Math.sqrt(Math.max(0, currentSumSq) / Math.max(1, count));
  }
  return rmsData;
};

const isNumericToken = (t) => {
  const v = parseFloat(t);
  return !Number.isNaN(v) && Number.isFinite(v);
};

// 線性插值函數 (處理遺失數據)
const linearInterpolate = (arr) => {
  const n = arr.length;
  let firstValidIdx = -1;
  for (let i = 0; i < n; i++) {
    if (!Number.isNaN(arr[i])) {
      firstValidIdx = i;
      break;
    }
  }

  if (firstValidIdx === -1) {
    for(let i = 0; i < n; i++) arr[i] = 0;
    return { interpolatedArr: arr, nanCount: n };
  }

  let nanCount = 0;
  for (let i = 0; i < firstValidIdx; i++) {
    arr[i] = arr[firstValidIdx];
    nanCount++;
  }

  let lastValidIdx = firstValidIdx;
  for (let i = firstValidIdx + 1; i < n; i++) {
    if (Number.isNaN(arr[i])) {
      let nextValidIdx = -1;
      for (let j = i + 1; j < n; j++) {
        if (!Number.isNaN(arr[j])) {
          nextValidIdx = j;
          break;
        }
      }
      
      if (nextValidIdx !== -1) {
        const startVal = arr[lastValidIdx];
        const endVal = arr[nextValidIdx];
        const steps = nextValidIdx - lastValidIdx;
        const delta = (endVal - startVal) / steps;
        for (let k = i; k < nextValidIdx; k++) {
          arr[k] = startVal + delta * (k - lastValidIdx);
          nanCount++;
        }
        i = nextValidIdx - 1; 
      } else {
        for (let k = i; k < n; k++) {
          arr[k] = arr[lastValidIdx];
          nanCount++;
        }
        break;
      }
    } else {
      lastValidIdx = i;
    }
  }
  return { interpolatedArr: arr, nanCount };
};

// 突波濾波器 (Spike Removal Filter) - 偵測異常跳動並套用線性插值補齊
const removeSpikes = (arr, threshold = 50) => {
  const n = arr.length;
  if (n === 0) return arr;
  const result = new Float64Array(n);
  
  let initialValid = arr[0];
  for (let i = 0; i < Math.min(20, n - 1); i++) {
    if (!Number.isNaN(arr[i]) && !Number.isNaN(arr[i+1]) && Math.abs(arr[i+1] - arr[i]) < threshold) {
      initialValid = arr[i];
      break;
    }
  }

  let lastValid = initialValid;
  let spikeCount = 0;

  for (let i = 0; i < n; i++) {
    if (Number.isNaN(arr[i])) {
      result[i] = NaN;
    } else if (Math.abs(arr[i] - lastValid) > threshold) {
      result[i] = NaN; 
      spikeCount++;
    } else {
      result[i] = arr[i];
      lastValid = arr[i]; 
    }
  }
  
  if (spikeCount > 0) {
      return linearInterpolate(result).interpolatedArr;
  }
  return result;
};


// 核心優化：嚴格的數據起點尋找
const findHeaderAndDataStart = (lines, splitLine) => {
  let dataStartIndex = -1;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const tokens = splitLine(line).filter(t => t !== '');
    if (tokens.length < 2) continue;

    const numericCount = tokens.filter(isNumericToken).length;
    
    if (numericCount >= tokens.length * 0.8) {
      dataStartIndex = i;
      break;
    }
  }
  let headerIndex = dataStartIndex > 0 ? dataStartIndex - 1 : -1;
  return { headerIndex, dataStartIndex };
};

const guessDelimiter = (lines) => {
  const candidates = ['\t', ',', ';', '|']; 
  let bestDelimiter = '\t';
  let maxValidScore = 0;

  candidates.forEach(delim => {
    let score = 0;
    for (let i = 0; i < Math.min(lines.length, 50); i++) {
      let parts = [];
      if (delim === ',') {
        parts = lines[i].split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
      } else {
        parts = lines[i].split(delim);
      }
      parts = parts.map(s => s.trim()).filter(s => s !== '');
      const nums = parts.filter(isNumericToken).length;
      if (parts.length > 1 && nums >= parts.length / 2) {
        score += parts.length;
      }
    }
    if (score > maxValidScore) {
      maxValidScore = score;
      bestDelimiter = delim;
    }
  });

  if (maxValidScore === 0) return /\s+/; 
  return bestDelimiter;
};

// 共用解析檔案引擎
const parseDataContent = (text) => {
  const lines = text.split(/\r?\n/);
  if (!lines.length) throw new Error("檔案內容為空！");

  const delim = guessDelimiter(lines);

  const parseLine = (line) => {
    let parts = [];
    if (delim instanceof RegExp) {
      const matches = line.match(/(?:[^\s"]+|"[^"]*")+/g);
      parts = matches || line.split(/\s+/);
    } else if (delim === ',') {
      parts = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
    } else {
      parts = line.split(delim);
    }
    return parts.map(s => s.replace(/^"|"$/g, '').trim());
  };

  const { headerIndex, dataStartIndex } = findHeaderAndDataStart(lines, parseLine);

  if (dataStartIndex === -1) {
    throw new Error("無法在檔案中找到連續的數值矩陣！請確保數據段落正確。");
  }

  const firstDataTokens = parseLine(lines[dataStartIndex]).filter(t => t !== '');
  const expectedCols = firstDataTokens.length;
  
  let extractedHeaders = [];
  if (headerIndex !== -1) {
    extractedHeaders = parseLine(lines[headerIndex]).filter(t => t !== '');
  }

  const finalHeaders = Array.from({ length: expectedCols }, (_, i) => {
    const defaultName = `第 ${i + 1} 欄`;
    return extractedHeaders[i] ? `${defaultName} (${extractedHeaders[i]})` : defaultName;
  });

  const dataLength = lines.length - dataStartIndex;
  const columns = Array.from({ length: expectedCols }, () => new Float64Array(dataLength));
  
  let validRowCount = 0;
  for (let i = dataStartIndex; i < lines.length; i++) {
    const line = lines[i];
    if (!line || line.length < 2) continue;
    
    const tokens = delim instanceof RegExp ? line.trim().split(delim) : line.split(delim);
    
    let colIdx = 0;
    for (let j = 0; j < tokens.length && colIdx < expectedCols; j++) {
      const str = tokens[j].trim();
      if (delim instanceof RegExp && str === '') continue;

      if (str === '') {
        columns[colIdx][validRowCount] = NaN;
      } else {
        const val = parseFloat(str);
        columns[colIdx][validRowCount] = val === val ? val : NaN; 
      }
      colIdx++;
    }
    while (colIdx < expectedCols) {
      columns[colIdx][validRowCount] = NaN;
      colIdx++;
    }
    validRowCount++;
  }

  let totalInterpolated = 0;
  const trimmedColumns = columns.map(col => {
    const trimmed = col.slice(0, validRowCount);
    const { interpolatedArr, nanCount } = linearInterpolate(trimmed);
    totalInterpolated += nanCount;
    return interpolatedArr;
  });

  return { finalHeaders, trimmedColumns, validRowCount, interpolatedCount: totalInterpolated }; 
};


// --- MVIC 歷史數據庫模組 ---
const MvicDatabase = ({ mvicData, setMvicData, onBack }) => {
  const [modal, setModal] = useState({ isOpen: false, type: '', muscle: '', index: -1, value: '' });

  const handleEdit = (muscle, index) => {
    setModal({ isOpen: true, type: 'edit', muscle, index, value: String(mvicData[muscle][index]) });
  };
  const handleDelete = (muscle, index) => {
    setModal({ isOpen: true, type: 'delete', muscle, index, value: '' });
  };
  const handleClearMuscle = (muscle) => {
    if (mvicData[muscle].length > 0) {
      setModal({ isOpen: true, type: 'clear', muscle, index: -1, value: '' });
    }
  };

  const confirmModal = () => {
    const { type, muscle, index, value } = modal;
    const newData = { ...mvicData };
    if (type === 'edit') {
      const parsedVal = parseFloat(value);
      if (!isNaN(parsedVal)) {
        newData[muscle] = [...newData[muscle]];
        newData[muscle][index] = parsedVal;
        setMvicData(newData);
      }
    } else if (type === 'delete') {
      newData[muscle] = newData[muscle].filter((_, i) => i !== index);
      setMvicData(newData);
    } else if (type === 'clear') {
      newData[muscle] = [];
      setMvicData(newData);
    }
    setModal({ isOpen: false, type: '', muscle: '', index: -1, value: '' });
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500 relative">
      {modal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-6 w-full max-w-sm animate-in zoom-in-95 duration-200">
            <h3 className="text-xl font-bold text-slate-900 mb-2">
              {modal.type === 'edit' && '修改測試數據'}
              {modal.type === 'delete' && '刪除測試數據'}
              {modal.type === 'clear' && '清空肌肉數據'}
            </h3>
            <div className="text-sm text-slate-600 mb-6 mt-4">
              {modal.type === 'edit' && (
                <div>
                  <p className="mb-3 font-medium text-slate-700">請輸入 <span className="text-indigo-600 font-bold">{modal.muscle}</span> 第 {modal.index + 1} 次的新數值 (mV)：</p>
                  <input 
                    type="number" value={modal.value} onChange={(e) => setModal((prev) => ({ ...prev, value: e.target.value }))}
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono text-lg" autoFocus
                  />
                </div>
              )}
              {modal.type === 'delete' && <p>確定要刪除 <span className="text-indigo-600 font-bold">{modal.muscle}</span> 的第 {modal.index + 1} 次數據嗎？<br/><br/>此操作無法復原。</p>}
              {modal.type === 'clear' && <p>確定要清空 <span className="text-indigo-600 font-bold">{modal.muscle}</span> 的所有數據嗎？<br/><br/>此操作無法復原。</p>}
            </div>
            <div className="flex justify-end gap-3 mt-2">
              <button onClick={() => setModal({ isOpen: false, type: '', muscle: '', index: -1, value: '' })} className="px-5 py-2.5 rounded-xl text-slate-500 hover:bg-slate-100 font-bold transition-colors">取消</button>
              <button onClick={confirmModal} className={`px-5 py-2.5 rounded-xl text-white font-bold transition-colors shadow-sm ${modal.type === 'edit' ? 'bg-emerald-500 hover:bg-emerald-600' : 'bg-rose-500 hover:bg-rose-600'}`}>確定{modal.type === 'edit' ? '修改' : '刪除'}</button>
            </div>
          </div>
        </div>
      )}

      <header className="max-w-7xl mx-auto flex items-center gap-4 bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500 hover:text-slate-800">
          <ArrowLeft size={24} />
        </button>
        <div className="bg-emerald-500 p-3 rounded-2xl shadow-lg text-white">
          <FolderOpen className="w-6 h-6" />
        </div>
        <div>
          <h1 className="text-xl font-bold text-slate-900">MVIC 歷史數據庫</h1>
        </div>
      </header>

      <main className="max-w-7xl mx-auto bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold text-sm">
              <tr>
                <th className="p-5">目標肌肉</th>
                <th className="p-5">Trial 1 (mV)</th>
                <th className="p-5">Trial 2 (mV)</th>
                <th className="p-5">Trial 3 (mV)</th>
                <th className="p-5 bg-indigo-50 text-indigo-800">綜合平均 (Mean)</th>
                <th className="p-5 bg-indigo-50 text-indigo-800">標準差 (SD)</th>
                <th className="p-5 text-center">操作</th>
              </tr>
            </thead>
            <tbody className="text-sm">
              {MUSCLE_LIST.map((muscle) => {
                const trials = mvicData[muscle];
                const mean = calcMean(trials);
                const sd = calcSD(trials, mean);
                return (
                  <tr key={muscle} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                    <td className="p-5 font-bold text-slate-800 flex items-center gap-2">
                      <div className={`w-2 h-2 rounded-full ${trials.length === 3 ? 'bg-emerald-500' : (trials.length > 0 ? 'bg-amber-400' : 'bg-slate-300')}`}></div>
                      {muscle}
                    </td>
                    {[0, 1, 2].map((idx) => (
                      <td key={idx} className="p-5 font-mono">
                        {trials[idx] !== undefined ? (
                          <div className="flex items-center gap-3 group">
                            <span className="font-semibold text-slate-700">{trials[idx].toFixed(4)}</span>
                            <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => handleEdit(muscle, idx)} className="text-blue-500 hover:bg-blue-50 p-1 rounded"><Edit2 size={14}/></button>
                              <button onClick={() => handleDelete(muscle, idx)} className="text-rose-500 hover:bg-rose-50 p-1 rounded"><Trash2 size={14}/></button>
                            </div>
                          </div>
                        ) : ( <span className="text-slate-300">-</span> )}
                      </td>
                    ))}
                    <td className="p-5 font-mono font-bold text-indigo-700 bg-indigo-50/30">{trials.length > 0 ? mean.toFixed(4) : '-'}</td>
                    <td className="p-5 font-mono font-bold text-indigo-700 bg-indigo-50/30">{trials.length > 1 ? sd.toFixed(4) : (trials.length === 1 ? '0.0000' : '-')}</td>
                    <td className="p-5 text-center">
                      <button 
                        onClick={() => handleClearMuscle(muscle)} 
                        disabled={trials.length === 0}
                        className={`text-xs font-bold px-3 py-1.5 rounded-lg transition-colors ${trials.length > 0 ? 'text-rose-600 hover:bg-rose-50' : 'text-slate-300 cursor-not-allowed'}`}
                      >清除全部</button>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  );
};

// --- 任務數據總表 (Task Database) 模組 ---
const TaskDatabase = ({ 
  taskLiftEmgData, setTaskLiftEmgData, taskLiftAngleData, setTaskLiftAngleData,
  taskTennisServeData, setTaskTennisServeData, taskTennisServeAngleData, setTaskTennisServeAngleData,
  onBack 
}) => {
  const [activeTask, setActiveTask] = useState('lifting');
  const [activeTab, setActiveTab] = useState('emg');
  const [modal, setModal] = useState({ isOpen: false, target: '', type: '' });

  const tasks = {
    lifting: { id: 'lifting', name: '舉手任務', icon: <ArrowUpRight size={18} />, emg: taskLiftEmgData, angle: taskLiftAngleData, setEmg: setTaskLiftEmgData, setAngle: setTaskLiftAngleData },
    tennis_serve: { id: 'tennis_serve', name: '網球發球分析', icon: <Target size={18} />, emg: taskTennisServeData, angle: taskTennisServeAngleData, setEmg: setTaskTennisServeData, setAngle: setTaskTennisServeAngleData }
  };

  const handleClear = (target, type) => {
    setModal({ isOpen: true, target, type });
  };

  const confirmClear = () => {
    const currentTaskActions = tasks[activeTask];
    if (modal.type === 'emg') {
      currentTaskActions.setEmg(prev => { const n = {...prev}; delete n[modal.target]; return n; });
    } else {
      currentTaskActions.setAngle(prev => { const n = {...prev}; delete n[modal.target]; return n; });
    }
    setModal({ isOpen: false, target: '', type: '' });
  };

  const currentTaskData = tasks[activeTask];
  const currentData = activeTab === 'emg' ? currentTaskData.emg : currentTaskData.angle;
  const displayKeys = Object.keys(currentData).filter(k => currentData[k] && currentData[k].length > 0);

  const getMean = (trials, phase) => {
    const validVals = trials.map(t => t[phase]).filter(v => v !== undefined && v !== '');
    if (validVals.length === 0) return '-';
    return (validVals.reduce((a, b) => a + parseFloat(b), 0) / validVals.length).toFixed(4);
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500 relative">
      {modal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-6 w-full max-w-sm animate-in zoom-in-95 duration-200">
            <h3 className="text-xl font-bold text-slate-900 mb-2">清空任務數據</h3>
            <p className="text-sm text-slate-600 mb-6 mt-4">
              確定要清空 <span className="font-bold text-rose-600">{modal.target}</span> 的所有儲存數據嗎？此操作無法復原。
            </p>
            <div className="flex justify-end gap-3">
              <button onClick={() => setModal({ isOpen: false, target: '', type: '' })} className="px-5 py-2.5 rounded-xl text-slate-500 hover:bg-slate-100 font-bold transition-colors">取消</button>
              <button onClick={confirmClear} className="px-5 py-2.5 rounded-xl text-white font-bold transition-colors shadow-sm bg-rose-500 hover:bg-rose-600">確定刪除</button>
            </div>
          </div>
        </div>
      )}

      <header className="max-w-7xl mx-auto flex items-center gap-4 bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500 hover:text-slate-800">
          <ArrowLeft size={24} />
        </button>
        <div className="bg-blue-500 p-3 rounded-2xl shadow-lg text-white">
          <Database className="w-6 h-6" />
        </div>
        <div>
          <h1 className="text-xl font-bold text-slate-900">任務數據總表</h1>
        </div>
      </header>

      <main className="max-w-7xl mx-auto space-y-4">
        <div className="flex flex-wrap gap-3 pb-2">
          {Object.values(tasks).map(task => (
            <button
              key={task.id}
              onClick={() => { setActiveTask(task.id); setActiveTab('emg'); }}
              className={`px-5 py-2.5 rounded-2xl font-bold transition-all flex items-center gap-2 text-sm shadow-sm ${activeTask === task.id ? 'bg-slate-800 text-white' : 'bg-white text-slate-500 hover:bg-slate-50 border border-slate-200'}`}
            >
              {task.icon} {task.name}
            </button>
          ))}
        </div>

        <div className="flex flex-wrap gap-3 pb-2">
          <button 
            onClick={() => setActiveTab('emg')} 
            className={`px-6 py-2.5 rounded-2xl font-bold transition-all flex items-center gap-2 text-sm shadow-sm ${activeTab === 'emg' ? 'bg-indigo-600 text-white' : 'bg-white text-slate-500 hover:bg-indigo-50 border border-slate-200'}`}
          >
            <Activity size={18} /> EMG 肌肉活化數據
          </button>
          <button 
            onClick={() => setActiveTab('angle')} 
            className={`px-6 py-2.5 rounded-2xl font-bold transition-all flex items-center gap-2 text-sm shadow-sm ${activeTab === 'angle' ? 'bg-emerald-600 text-white' : 'bg-white text-slate-500 hover:bg-emerald-50 border border-slate-200'}`}
          >
            <Eye size={18} /> 觀察關節角度數據
          </button>
        </div>

        <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden min-h-[400px]">
          {displayKeys.length === 0 ? (
            <div className="p-20 text-center flex flex-col items-center justify-center h-full">
              <Database size={64} className="text-slate-200 mb-4" />
              <h3 className="text-lg font-bold text-slate-400">「{currentTaskData.name} - {activeTab === 'emg' ? 'EMG' : '觀察角度'}」尚無儲存數據</h3>
              <p className="text-sm text-slate-400 mt-2">請先前往對應的分析模組進行分析並寫入資料庫。</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold text-sm">
                  <tr>
                    <th className="p-5">{activeTab === 'emg' ? '目標' : '觀察通道'}</th>
                    <th className="p-5">動作階段 (Phase)</th>
                    <th className="p-5">Trial 1</th>
                    <th className="p-5">Trial 2</th>
                    <th className="p-5">Trial 3</th>
                    <th className={`p-5 text-white ${activeTab === 'emg' ? 'bg-indigo-500' : 'bg-emerald-500'}`}>階段平均 (Mean)</th>
                    <th className={`p-5 text-white ${activeTab === 'emg' ? 'bg-indigo-400' : 'bg-emerald-400'}`}>標準差 (SD)</th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-100">
                  {displayKeys.map(targetKey => {
                    const trials = currentData[targetKey];
                    let phases = ['Overall'];
                    if (activeTask === 'lifting') {
                      phases = activeTab === 'emg' 
                        ? ['Up_30-60', 'Up_60-90', 'Up_90-120', 'Down_120-90', 'Down_90-60', 'Down_60-30']
                        : ['Up_30', 'Up_60', 'Up_90', 'Down_90', 'Down_60', 'Down_30'];
                    } else if (activeTask === 'tennis_serve') {
                      phases = activeTab === 'emg'
                        ? ['Cocking', 'Acceleration', 'Deceleration']
                        : ['Start', 'MinPlane', 'Impact', 'MaxPlane'];
                    }
                    
                    return phases.map((phase, pIdx) => {
                      let t1, t2, t3, mean, sd;
                      if (activeTask === 'lifting' || activeTask === 'tennis_serve') {
                        t1 = trials[0]?.[phase];
                        t2 = trials[1]?.[phase];
                        t3 = trials[2]?.[phase];
                        mean = getMean(trials, phase);
                        const validVals = trials.map(t => t[phase]).filter(v => v !== undefined && v !== '').map(Number);
                        sd = validVals.length > 1 ? calcSD(validVals, parseFloat(mean)).toFixed(4) : '-';
                      } else {
                        t1 = trials[0];
                        t2 = trials[1];
                        t3 = trials[2];
                        const validVals = trials.filter(v => v !== undefined && v !== '').map(Number);
                        mean = validVals.length > 0 ? (validVals.reduce((a,b)=>a+b,0)/validVals.length).toFixed(4) : '-';
                        sd = validVals.length > 1 ? calcSD(validVals, parseFloat(mean)).toFixed(4) : '-';
                      }
                      
                      const isEmg = activeTab === 'emg';
                      const formattedPhase = phase.replace('_', ' ');

                      return (
                        <tr key={`${targetKey}-${phase}`} className="hover:bg-slate-50 transition-colors">
                          {pIdx === 0 && (
                            <td rowSpan={phases.length} className="p-5 align-top border-r border-slate-100 bg-white">
                              <div className="font-bold text-slate-800 text-base">{targetKey}</div>
                              <div className="mt-1.5 flex items-center gap-2">
                                <span className={`text-xs font-bold px-2 py-0.5 rounded-md ${isEmg ? 'bg-indigo-100 text-indigo-700' : 'bg-emerald-100 text-emerald-700'}`}>
                                  已存 {trials.length}/3 次
                                </span>
                              </div>
                              <button 
                                onClick={() => handleClear(targetKey, activeTab)} 
                                className="mt-4 flex items-center gap-1 text-xs font-bold text-rose-500 hover:text-rose-600 hover:underline transition-colors"
                              >
                                <Trash2 size={14} /> 清除此目標
                              </button>
                            </td>
                          )}
                          <td className="p-5 font-bold text-slate-600 border-r border-slate-50 bg-slate-50/30">
                            {formattedPhase}{activeTask === 'lifting' ? '°' : ''}
                          </td>
                          <td className="p-5 font-mono text-slate-700">{t1 !== undefined && t1 !== '' ? t1 : <span className="text-slate-300">-</span>}</td>
                          <td className="p-5 font-mono text-slate-700">{t2 !== undefined && t2 !== '' ? t2 : <span className="text-slate-300">-</span>}</td>
                          <td className="p-5 font-mono text-slate-700">{t3 !== undefined && t3 !== '' ? t3 : <span className="text-slate-300">-</span>}</td>
                          <td className={`p-5 font-mono font-bold ${isEmg ? 'text-indigo-700 bg-indigo-50/40' : 'text-emerald-700 bg-emerald-50/40'}`}>
                            {mean !== '-' ? mean : <span className="text-slate-300">-</span>}
                          </td>
                          <td className={`p-5 font-mono font-bold ${isEmg ? 'text-indigo-600 bg-indigo-50/20' : 'text-emerald-600 bg-emerald-50/20'}`}>
                            {sd !== '-' ? sd : <span className="text-slate-300">-</span>}
                          </td>
                        </tr>
                      );
                    });
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

// --- 舉手動作分析 (Lifting Task) 模組 ---
const LiftingAnalysis = ({ onBack, taskLiftEmgData, setTaskLiftEmgData, taskLiftAngleData, setTaskLiftAngleData }) => {
  const [emgFileResult, setEmgFileResult] = useState(null);
  const [emgHeaders, setEmgHeaders] = useState([]);
  const [kinFileResult, setKinFileResult] = useState(null);
  const [kinHeaders, setKinHeaders] = useState([]);

  const [errorMessage, setErrorMessage] = useState(null);
  const [toastMessage, setToastMessage] = useState(null);

  const [kinAngleColIdx, setKinAngleColIdx] = useState(1);
  const [kinTrigColIdx, setKinTrigColIdx] = useState(0);

  const [emgSR, setEmgSR] = useState(1500);
  const [kinSR, setKinSR] = useState(960); 
  const [trigThresh, setTrigThresh] = useState(3.0);
  
  const [kinOnsetConsecutive, setKinOnsetConsecutive] = useState(50);
  const [kinSpikeThresh, setKinSpikeThresh] = useState(50); 

  const [bpHigh, setBpHigh] = useState(20);
  const [bpLow, setBpLow] = useState(450);
  const [lpfCutoff, setLpfCutoff] = useState(20); 

  const [analysisResult, setAnalysisResult] = useState(null);
  const [selectedRepIdx, setSelectedRepIdx] = useState(0);
  const [visibleEmgMuscle, setVisibleEmgMuscle] = useState(MUSCLE_LIST[0]);

  const [draggingMarker, setDraggingMarker] = useState(null);

  const LIFT_KIN_LIST = [
    { key: 'RScapUpDownRotation', search: 'rscapupdown' },
    { key: 'RScapAntPosTilt', search: 'rscapantpos' },
    { key: 'RScapIntExtRotation', search: 'rscapinext' }
  ];

  const handleEmgUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const { finalHeaders, trimmedColumns, interpolatedCount } = parseDataContent(e.target.result);
        setEmgHeaders(finalHeaders);
        setEmgFileResult(trimmedColumns);

        setErrorMessage(null);
        if (interpolatedCount > 0) showToast(`⚠️ 偵測到 ${interpolatedCount} 筆 EMG 遺失數據，已自動線性插值補齊！`);
      } catch (err) { setErrorMessage(`EMG 解析失敗: ${err.message}`); }
    };
    reader.readAsText(file);
  };

  const handleKinUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const { finalHeaders, trimmedColumns, interpolatedCount } = parseDataContent(e.target.result);
        setKinHeaders(finalHeaders);
        setKinFileResult(trimmedColumns);
        
        const trigIdx = finalHeaders.findIndex(h => h.toLowerCase().includes('trigger') || h.toLowerCase().includes('trig'));
        if (trigIdx !== -1) setKinTrigColIdx(trigIdx);

        const rhtIdx = finalHeaders.findIndex(h => h.replace(/[^a-zA-Z]/g, '').toLowerCase().includes('rhtelevation'));
        if (rhtIdx !== -1) setKinAngleColIdx(rhtIdx);

        setErrorMessage(null);
        if (interpolatedCount > 0) showToast(`⚠️ 偵測到 ${interpolatedCount} 筆 KINEMATIC 遺失數據，已自動線性插值補齊！`);
      } catch (err) { setErrorMessage(`Kinematic 解析失敗: ${err.message}`); }
    };
    reader.readAsText(file);
  };

  const recalculateCycleIndices = (startIdx, peakIdx, endIdx, kinAngleData) => {
    const getUpIdx = (startI, endI, threshold) => {
       for(let i=startI; i<=endI; i++) if (kinAngleData[i] >= threshold) return i;
       return null;
    };
    const getDownIdx = (startI, endI, threshold) => {
       if (kinAngleData[startI] < threshold) return null; 
       for(let i=startI; i<=endI; i++) if (kinAngleData[i] <= threshold) return i;
       return null;
    };

    let i30_up = getUpIdx(startIdx, peakIdx, 30);
    let i60_up = getUpIdx(i30_up || startIdx, peakIdx, 60);
    let i90_up = getUpIdx(i60_up || startIdx, peakIdx, 90);
    let i120_up = getUpIdx(i90_up || startIdx, peakIdx, 120);

    let i120_down = getDownIdx(peakIdx, endIdx, 120);
    let i90_down = getDownIdx(i120_down || peakIdx, endIdx, 90);
    let i60_down = getDownIdx(i90_down || peakIdx, endIdx, 60);
    let i30_down = getDownIdx(i60_down || peakIdx, endIdx, 30);

    return { i30_up, i60_up, i90_up, i120_up, i120_down, i90_down, i60_down, i30_down };
  };

  const generateMetricsForCycle = (cycle, emgDataArrays, kinDataArrays, kinTrigIdx, emgTrigIdx, localKinSR, localEmgSR) => {
    
    const getEmgIdx = (t) => emgTrigIdx + Math.round(t * localEmgSR);

    const calcEmgSegment = (rmsArr, sIdx, eIdx) => {
      if (sIdx === null || eIdx === null || sIdx >= eIdx) return '';
      const tStart = (sIdx - kinTrigIdx) / localKinSR;
      const tEnd = (eIdx - kinTrigIdx) / localKinSR;
      
      const emgStart = Math.max(0, getEmgIdx(tStart));
      const emgEnd = Math.min(rmsArr.length - 1, getEmgIdx(tEnd));
      
      let sumSq = 0, countRms = 0;
      for(let i=emgStart; i<=emgEnd && i<rmsArr.length; i++) { 
        sumSq += Math.pow(rmsArr[i], 2); countRms++; 
      }
      return countRms > 0 ? +(Math.sqrt(sumSq / countRms)).toFixed(4) : '';
    };

    const emgMetrics = emgDataArrays.map(item => ({
      muscle: item.muscle,
      segments: {
        'Up_30-60': calcEmgSegment(item.rmsEnvelope, cycle.indices.i30_up, cycle.indices.i60_up),
        'Up_60-90': calcEmgSegment(item.rmsEnvelope, cycle.indices.i60_up, cycle.indices.i90_up),
        'Up_90-120': calcEmgSegment(item.rmsEnvelope, cycle.indices.i90_up, cycle.indices.i120_up),
        'Down_120-90': calcEmgSegment(item.rmsEnvelope, cycle.indices.i120_down, cycle.indices.i90_down),
        'Down_90-60': calcEmgSegment(item.rmsEnvelope, cycle.indices.i90_down, cycle.indices.i60_down),
        'Down_60-30': calcEmgSegment(item.rmsEnvelope, cycle.indices.i60_down, cycle.indices.i30_down),
        'Overall': calcEmgSegment(item.rmsEnvelope, cycle.startIdx, cycle.endIdx)
      }
    }));

    const safeGet = (arr, idx) => (arr && idx !== null && idx >= 0 && idx < arr.length) ? +(arr[idx]).toFixed(2) : '-';
    
    const kinMetrics = kinDataArrays.map(item => ({
      field: item.field,
      points: {
        'Up_30': safeGet(item.data, cycle.indices.i30_up),
        'Up_60': safeGet(item.data, cycle.indices.i60_up),
        'Up_90': safeGet(item.data, cycle.indices.i90_up),
        'Down_90': safeGet(item.data, cycle.indices.i90_down),
        'Down_60': safeGet(item.data, cycle.indices.i60_down),
        'Down_30': safeGet(item.data, cycle.indices.i30_down)
      }
    }));

    return { emgMetrics, kinMetrics };
  };

  const processLiftingTask = () => {
    if (!emgFileResult || !kinFileResult) { setErrorMessage("請先載入 EMG 與 KINEMATIC 兩個檔案！"); return; }
    setErrorMessage(null); setAnalysisResult(null); setSelectedRepIdx(0);

    const kinTriggerData = kinFileResult[kinTrigColIdx];
    const kinAngleDataRaw = kinFileResult[kinAngleColIdx];

    // 1. 尋找 KINEMATIC 檔案中的 Trigger 點 ( > 3V ) 作為 t=0
    let kinTrigIdx = -1;
    for (let i = 0; i < kinTriggerData.length; i++) {
      if (kinTriggerData[i] >= trigThresh) { kinTrigIdx = i; break; }
    }
    if (kinTrigIdx === -1) { setErrorMessage(`Kinematic 同步失敗：找不到大於 ${trigThresh} 的 Trigger 訊號。`); return; }

    // 2. 尋找 EMG 檔案中的 Time=0 點作為同步起點 ( 跨越 0 ) 作為 t=0
    let emgTrigIdx = -1;
    const emgTimeColIdx = emgHeaders.findIndex(h => h.toLowerCase().replace(/[^a-z]/g, '').includes('time'));
    if (emgTimeColIdx !== -1) {
        const emgTimeData = emgFileResult[emgTimeColIdx];
        for (let i = 0; i < emgTimeData.length; i++) {
            if (emgTimeData[i] >= 0) { 
                emgTrigIdx = i;
                break;
            }
        }
    }
    // 如果 EMG 檔案沒有 Time 欄位，退回相對頻率推算
    if (emgTrigIdx === -1 || emgTrigIdx === 0) {
        emgTrigIdx = Math.round((kinTrigIdx / kinSR) * emgSR);
    }

    const kinAngleData = removeSpikes(kinAngleDataRaw, kinSpikeThresh);

    const detectedCycles = [];
    let firstOnsetIdx = -1;
    let consecutiveUp = 0;
    
    for (let i = kinTrigIdx; i < kinAngleData.length; i++) {
      if (i === 0) continue;
      const delta = kinAngleData[i] - kinAngleData[i-1];
      if (delta > 0) {
        consecutiveUp++;
        if (consecutiveUp >= kinOnsetConsecutive) { firstOnsetIdx = i - kinOnsetConsecutive + 1; break; }
      } else { consecutiveUp = 0; }
    }

    if (firstOnsetIdx === -1) { setErrorMessage(`無法找到連續上升的起點，請調整參數。`); return; }

    let currentStartIdx = firstOnsetIdx;
    while (currentStartIdx < kinAngleData.length - 1 && detectedCycles.length < 3) {
      let peakIdx = currentStartIdx;
      let maxAngle = kinAngleData[currentStartIdx];
      for (let i = currentStartIdx + 1; i < kinAngleData.length; i++) {
        if (kinAngleData[i] > maxAngle) { maxAngle = kinAngleData[i]; peakIdx = i; } 
        else if (maxAngle - kinAngleData[i] > 10) break; 
      }

      let endIdx = peakIdx;
      let minAngle = kinAngleData[peakIdx];
      for (let i = peakIdx + 1; i < kinAngleData.length; i++) {
        if (kinAngleData[i] < minAngle) { minAngle = kinAngleData[i]; endIdx = i; } 
        else if (kinAngleData[i] - minAngle > 10) break; 
      }

      if (maxAngle - kinAngleData[currentStartIdx] >= 15) {
        detectedCycles.push({ 
          id: detectedCycles.length + 1,
          startIdx: currentStartIdx, 
          peakIdx: peakIdx, 
          endIdx: endIdx,
          tStart: +( (currentStartIdx - kinTrigIdx) / kinSR ).toFixed(3),
          tPeak: +( (peakIdx - kinTrigIdx) / kinSR ).toFixed(3),
          tEnd: +( (endIdx - kinTrigIdx) / kinSR ).toFixed(3),
        });
      } else break; 
      
      if (endIdx === peakIdx || endIdx >= kinAngleData.length - 1) break; 
      currentStartIdx = endIdx;
    }

    if (detectedCycles.length === 0) { setErrorMessage(`無法找出完整循環。`); return; }

    const emgDataArrays = MUSCLE_LIST.map((muscle, idx) => {
      let colIdx = emgHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(muscle.toLowerCase()));
      if (colIdx === -1) colIdx = Math.min(idx, emgFileResult.length - 1);

      const emgData = emgFileResult[colIdx];
      const filtered = bandpassFilter(emgData, bpHigh, bpLow, emgSR);
      const rectified = new Float64Array(filtered.length);
      for(let i = 0; i < filtered.length; i++) rectified[i] = Math.abs(filtered[i]);
      const rmsEnvelope = zeroLagBiquadFilter(rectified, 'lowpass', lpfCutoff, emgSR); 
      return { muscle, rmsEnvelope };
    });

    const kinDataArrays = LIFT_KIN_LIST.map(item => {
      const colIdx = kinHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(item.search));
      const dataRaw = colIdx !== -1 ? kinFileResult[colIdx] : null;
      const data = dataRaw ? removeSpikes(dataRaw, kinSpikeThresh) : null;
      return { field: item.key, data };
    });

    const processedCycles = detectedCycles.map(cycle => {
      cycle.indices = recalculateCycleIndices(cycle.startIdx, cycle.peakIdx, cycle.endIdx, kinAngleData);
      const metrics = generateMetricsForCycle(cycle, emgDataArrays, kinDataArrays, kinTrigIdx, emgTrigIdx, kinSR, emgSR);
      
      return {
        ...cycle,
        emgMetrics: metrics.emgMetrics,
        kinMetrics: metrics.kinMetrics,
        maxAngle: kinAngleData[cycle.peakIdx].toFixed(1),
        duration: +( (cycle.endIdx - cycle.startIdx) / kinSR ).toFixed(2)
      };
    });

    const chartData = [];
    for (let i = 0; i < kinAngleData.length; i++) {
      const relT = (i - kinTrigIdx) / kinSR;
      const matchingEmgIdx = emgTrigIdx + Math.round(relT * emgSR);
      
      const point = {
        time: Math.round(relT * 1000) / 1000,
        angleMain: Math.round(kinAngleData[i] * 100) / 100
      };

      emgDataArrays.forEach(m => {
        point[`emg_${m.muscle}`] = (matchingEmgIdx >= 0 && matchingEmgIdx < m.rmsEnvelope.length) 
          ? Math.round(m.rmsEnvelope[matchingEmgIdx] * 10000) / 10000 : null;
      });

      kinDataArrays.forEach(k => {
        point[`kin_${k.field}`] = k.data ? Math.round(k.data[i] * 100) / 100 : null;
      });

      chartData.push(point);
    }

    setAnalysisResult({ 
      chartData, 
      cycles: processedCycles, 
      kinTrigIdx,
      emgTrigIdx,
      emgDataArrays,
      kinDataArrays,
      kinAngleData 
    });
  };

  const showToast = (msg) => { setToastMessage(msg); setTimeout(() => setToastMessage(null), 3000); };

  const handleSaveAllData = () => {
    if (!analysisResult) return;

    const firstMuscle = MUSCLE_LIST[0];
    const currentDataCount = (taskLiftEmgData[firstMuscle] || []).length;
    if (currentDataCount >= 3) {
      showToast(`❌ 已達 3 次儲存上限！請先至資料庫刪除舊資料。`);
      return;
    }

    const cycle = analysisResult.cycles[selectedRepIdx];

    setTaskLiftEmgData(prev => {
      const newData = { ...prev };
      cycle.emgMetrics.forEach(item => {
        if (!newData[item.muscle]) newData[item.muscle] = [];
        newData[item.muscle] = [...newData[item.muscle], { ...item.segments }];
      });
      return newData;
    });

    setTaskLiftAngleData(prev => {
      const newData = { ...prev };
      cycle.kinMetrics.forEach(kin => {
        if (kin.field === '-' || !kin.field) return;
        if (!newData[kin.field]) newData[kin.field] = [];
        newData[kin.field] = [...newData[kin.field], { ...kin.points }];
      });
      return newData;
    });

    showToast(`✅ 成功一鍵寫入！已將 7 條 EMG 與 3 個肩胛角度寫入資料庫 (第 ${currentDataCount + 1} 次)`);
  };

  const handleNextRepetition = () => {
    if (analysisResult && selectedRepIdx < analysisResult.cycles.length - 1) setSelectedRepIdx(selectedRepIdx + 1);
    else showToast("⚠️ 後續沒有找到更多循環動作了！");
  };

  const handleChartMouseDown = useCallback((e) => {
    if (e && e.activePayload && analysisResult) {
      const time = e.activePayload[0].payload.time;
      const cycle = analysisResult.cycles[selectedRepIdx];
      const minD = Math.min(Math.abs(time - cycle.tStart), Math.abs(time - cycle.tPeak), Math.abs(time - cycle.tEnd));
      if (minD < 0.5) {
        if (minD === Math.abs(time - cycle.tStart)) setDraggingMarker('start');
        else if (minD === Math.abs(time - cycle.tPeak)) setDraggingMarker('peak');
        else if (minD === Math.abs(time - cycle.tEnd)) setDraggingMarker('end');
      }
    }
  }, [analysisResult, selectedRepIdx]);

  const handleChartMouseMove = useCallback((e) => {
    if (draggingMarker && e && e.activePayload && analysisResult) {
      const time = e.activePayload[0].payload.time;
      const newCycles = [...analysisResult.cycles];
      const cycle = { ...newCycles[selectedRepIdx] };
      
      const idx = Math.max(0, analysisResult.kinTrigIdx + Math.round(time * kinSR));
      const maxIdx = kinFileResult[kinAngleColIdx].length - 1;

      if (draggingMarker === 'start') { if (idx < cycle.peakIdx) { cycle.startIdx = Math.max(0, idx); cycle.tStart = time; } } 
      else if (draggingMarker === 'peak') { if (idx > cycle.startIdx && idx < cycle.endIdx) { cycle.peakIdx = idx; cycle.tPeak = time; } } 
      else if (draggingMarker === 'end') { if (idx > cycle.peakIdx) { cycle.endIdx = Math.min(maxIdx, idx); cycle.tEnd = time; } }
      
      cycle.indices = recalculateCycleIndices(cycle.startIdx, cycle.peakIdx, cycle.endIdx, analysisResult.kinAngleData);
      const metrics = generateMetricsForCycle(cycle, analysisResult.emgDataArrays, analysisResult.kinDataArrays, analysisResult.kinTrigIdx, analysisResult.emgTrigIdx, kinSR, emgSR);
      
      cycle.emgMetrics = metrics.emgMetrics;
      cycle.kinMetrics = metrics.kinMetrics;
      cycle.maxAngle = analysisResult.kinAngleData[cycle.peakIdx].toFixed(1);
      cycle.duration = +( (cycle.tEnd - cycle.tStart) ).toFixed(2);

      newCycles[selectedRepIdx] = cycle;
      setAnalysisResult({ ...analysisResult, cycles: newCycles });
    }
  }, [draggingMarker, analysisResult, selectedRepIdx, kinSR, emgSR, kinFileResult, kinAngleColIdx]);

  const handleChartMouseUp = useCallback(() => setDraggingMarker(null), []);

  const currentMetrics = analysisResult?.cycles[selectedRepIdx];
  const emgKeys = ['Up_30-60', 'Up_60-90', 'Up_90-120', 'Down_120-90', 'Down_90-60', 'Down_60-30'];
  const kinKeys = ['Up_30', 'Up_60', 'Up_90', 'Down_90', 'Down_60', 'Down_30'];

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500 relative" onMouseUp={handleChartMouseUp} onMouseLeave={handleChartMouseUp}>
      {toastMessage && (<div className="fixed top-8 left-1/2 transform -translate-x-1/2 z-50 bg-slate-800 text-white px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-3 animate-in slide-in-from-top-4 duration-300"><span className="font-bold text-sm">{toastMessage}</span></div>)}
      <header className="max-w-7xl mx-auto flex flex-col xl:flex-row justify-between items-start xl:items-center bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6 gap-4">
        <div className="flex items-center gap-4 shrink-0">
          <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500 hover:text-slate-800"><ArrowLeft size={24} /></button>
          <div className="bg-blue-500 p-3 rounded-2xl shadow-lg"><ArrowUpRight className="text-white w-6 h-6" /></div>
          <div><h1 className="text-xl font-bold text-slate-900">舉手動作分析 (Lifting)</h1></div>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
          <label className={`flex items-center gap-2 px-5 py-2.5 rounded-2xl transition-all shadow-sm cursor-pointer text-sm font-bold shrink-0 ${emgFileResult ? 'bg-indigo-100 text-indigo-700' : 'bg-indigo-600 hover:bg-indigo-700 text-white'}`}>
            <Upload size={18} /> {emgFileResult ? '已載入 EMG' : '載入 EMG 檔'}<input type="file" className="hidden" accept=".csv,.txt" onChange={handleEmgUpload} />
          </label>
          <label className={`flex items-center gap-2 px-5 py-2.5 rounded-2xl transition-all shadow-sm cursor-pointer text-sm font-bold shrink-0 ${kinFileResult ? 'bg-emerald-100 text-emerald-700' : 'bg-emerald-600 hover:bg-emerald-700 text-white'}`}>
            <Upload size={18} /> {kinFileResult ? '已載入 KINEMATIC' : '載入 KINEMATIC 檔'}<input type="file" className="hidden" accept=".csv,.txt" onChange={handleKinUpload} />
          </label>
        </div>
      </header>
      <main className="max-w-7xl mx-auto space-y-6">
        {errorMessage && ( <div className="bg-rose-50 border border-rose-200 text-rose-700 px-6 py-4 rounded-2xl font-bold flex items-center gap-3"><Info size={20} /> {errorMessage}</div> )}
        {emgFileResult && kinFileResult && (
          <div className="bg-white p-5 rounded-3xl shadow-sm border border-slate-100 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-[1fr_1fr_1.2fr_1fr_120px] gap-3">
              <div className="bg-indigo-50/50 p-3 rounded-2xl border border-indigo-100/50 flex flex-col justify-between">
                <div className="text-sm font-bold text-indigo-800 mb-2 uppercase tracking-wide flex items-center gap-2"><Activity size={14} className="shrink-0" /> 同步設定</div>
                <div className="grid grid-cols-2 gap-x-2 gap-y-1.5">
                  <div className="col-span-2"><span className="text-xs font-semibold text-slate-500 mb-1 block">KIN Trigger:</span><select value={kinTrigColIdx} onChange={e=>setKinTrigColIdx(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-slate-700 bg-white">{kinHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}</select></div>
                  <div><span className="text-xs font-semibold text-slate-500 mb-1 block">閥值(V):</span><input type="number" step="0.5" value={trigThresh} onChange={e=>setTrigThresh(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center text-rose-600 bg-white" /></div>
                  <div><span className="text-xs font-semibold text-slate-500 mb-1 block">Kin SR (Hz):</span><input type="number" value={kinSR} onChange={e=>setKinSR(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center bg-white" /></div>
                </div>
              </div>
              <div className="bg-slate-50 p-3 rounded-2xl border border-slate-200/50 flex flex-col justify-between">
                <div className="text-sm font-bold text-slate-800 mb-2 uppercase tracking-wide flex items-center gap-2"><Activity size={14} className="shrink-0" /> EMG 參數</div>
                <div className="space-y-2">
                  <div className="flex flex-col gap-1">
                    <span className="text-xs font-semibold text-slate-500 mb-1 block">自動全解析肌肉通道 (RMS):</span>
                    <div className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-600 bg-white shadow-sm flex items-center gap-2">
                      <Activity size={14} className="text-indigo-400" /> {MUSCLE_LIST.join(', ')}
                    </div>
                  </div>
                  <div><span className="text-xs font-semibold text-slate-500 mb-1 block">取樣頻率(Hz):</span><input type="number" value={emgSR} onChange={e=>setEmgSR(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center bg-white" /></div>
                </div>
              </div>
              <div className="bg-amber-50/50 p-3 rounded-2xl border border-amber-100/50 flex flex-col justify-between">
                <div className="text-sm font-bold text-amber-800 mb-2 uppercase tracking-wide flex items-center gap-2"><Crosshair size={14} className="shrink-0" /> 判定關節 & 演算法設定</div>
                <div className="grid grid-cols-2 gap-x-2 gap-y-1.5">
                  <div className="col-span-2"><span className="text-xs font-semibold text-slate-500 mb-1 block">主要角度通道:</span><select value={kinAngleColIdx} onChange={e=>setKinAngleColIdx(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-slate-700 bg-white">{kinHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}</select></div>
                  <div><span className="text-xs font-semibold text-slate-500 mb-1 block">Onset 上升筆數:</span><input type="number" value={kinOnsetConsecutive} onChange={e=>setKinOnsetConsecutive(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-black text-center bg-white" /></div>
                  <div><span className="text-xs font-semibold text-slate-500 mb-1 block">突波濾除 (deg):</span><input type="number" value={kinSpikeThresh} onChange={e=>setKinSpikeThresh(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-amber-300 text-sm font-black text-amber-700 text-center bg-white" /></div>
                </div>
              </div>
              <div className="bg-emerald-50/50 p-3 rounded-2xl border border-emerald-100/50 flex flex-col justify-between">
                <div className="text-sm font-bold text-emerald-800 mb-2 uppercase tracking-wide flex items-center gap-2"><Eye size={14} className="shrink-0" /> 觀察關節通道</div>
                <div className="space-y-2 flex-1">
                  <span className="text-[11px] font-semibold text-slate-500">系統將自動配對擷取以下肩胛骨數值：</span>
                  <div className="flex flex-col gap-1.5">
                    {LIFT_KIN_LIST.map(k => (
                      <div key={k.key} className="text-[10px] font-bold text-emerald-700 bg-emerald-100/50 px-2 py-1 rounded border border-emerald-200">{k.key}</div>
                    ))}
                  </div>
                </div>
              </div>
              <div className="flex items-end shrink-0">
                <button onClick={processLiftingTask} className="w-full bg-blue-600 hover:bg-blue-700 text-white p-3 h-full min-h-[84px] rounded-2xl font-bold transition-all shadow-lg active:scale-95 flex flex-col items-center justify-center gap-2 hover:shadow-blue-200">
                  <Activity size={24} /> <span className="text-sm">開始分析</span>
                </button>
              </div>
            </div>
          </div>
        )}
        {analysisResult && currentMetrics && (
          <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <MetricCard title="代表性肌肉 (UT) 平均 RMS" value={currentMetrics.emgMetrics.find(m=>m.muscle==='UT')?.segments['Overall'] || '-'} unit="mV" icon={<BarChart className="text-blue-500" />} />
              <MetricCard title="最高峰值角度 (Peak)" value={currentMetrics.maxAngle} unit="°" icon={<Layers className="text-amber-500" />} />
              <MetricCard title="循環總時長" value={currentMetrics.duration} unit="s" icon={<Info className="text-indigo-500" />} />
              <div className="flex items-center justify-between bg-slate-800 p-4 rounded-3xl border border-slate-700 shadow-sm text-white">
                <div className="flex flex-col justify-center">
                  <span className="text-xs font-bold text-slate-400 mb-1">目前檢視動作</span>
                  <span className="text-3xl font-black">{currentMetrics.id} <span className="text-sm font-medium text-slate-400">/ {analysisResult.cycles.length}</span></span>
                </div>
                <button onClick={handleNextRepetition} disabled={selectedRepIdx >= analysisResult.cycles.length - 1} className={`p-3 rounded-xl font-bold flex items-center justify-center transition-all ${selectedRepIdx < analysisResult.cycles.length - 1 ? 'bg-blue-600 hover:bg-blue-500 active:scale-95' : 'bg-slate-700 text-slate-500'}`}><ArrowRight size={20} /></button>
              </div>
            </div>

            <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
              <h3 className="text-md font-bold text-slate-800 flex items-center gap-2 mb-4">
                <Activity size={18} className="text-indigo-500" /> 
                EMG {MUSCLE_LIST.length} 條肌肉全活化分期指標 (Mean RMS)
              </h3>
              <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-inner max-h-[400px]">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-100 text-slate-600 font-bold sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="p-3 border-b border-slate-300 min-w-[80px]">肌肉</th>
                      {emgKeys.map(phase => (
                         <th key={phase} className="p-3 border-b border-slate-300 text-center whitespace-nowrap text-indigo-700">{phase.replace('_', ' ')}°</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {currentMetrics.emgMetrics.map((item) => (
                      <tr key={item.muscle} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 font-black text-slate-700 bg-slate-50/50 border-r border-slate-100">{item.muscle}</td>
                        {emgKeys.map(phase => (
                          <td key={phase} className="p-3 font-mono font-bold text-center text-slate-600">{item.segments[phase] || '-'}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
              <h3 className="text-md font-bold text-slate-800 flex items-center gap-2 mb-4">
                <Layers size={18} className="text-emerald-500" /> 
                3 項肩胛骨觀察關節角度變化
              </h3>
              <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-inner max-h-[300px]">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-100 text-slate-600 font-bold sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="p-3 border-b border-slate-300 min-w-[150px]">關節通道</th>
                      {kinKeys.map(phase => (
                         <th key={phase} className="p-3 border-b border-slate-300 text-center whitespace-nowrap text-emerald-700">{phase.replace('_', ' ')}°</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {currentMetrics.kinMetrics.map((item) => (
                      <tr key={item.field} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 font-bold text-slate-700 bg-slate-50/50 border-r border-slate-100">{item.field}</td>
                        {kinKeys.map(phase => (
                          <td key={phase} className="p-3 font-mono text-center text-slate-600">{item.points[phase] || '-'}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-blue-50 border border-blue-200 p-5 rounded-3xl flex flex-wrap items-center justify-between gap-4 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="bg-blue-100 p-2 rounded-xl text-blue-600"><Save size={20} /></div>
                <div>
                  <h3 className="font-bold text-blue-800">一鍵批量儲存任務數據</h3>
                  <p className="text-xs text-blue-600 font-medium">將上述表格中所有分析結果 (EMG 與 角度) 完整寫入總資料庫</p>
                </div>
              </div>
              <button onClick={handleSaveAllData} className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-2.5 rounded-xl font-bold transition-colors shadow-sm active:scale-95">
                一鍵寫入資料庫
              </button>
            </div>

            <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
              <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-3">
                <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Waves size={18} className="text-indigo-500" /> 同步分析波形圖
                </h3>
                <div className="flex items-center gap-3">
                  <span className="text-xs font-bold text-slate-500">選擇顯示的 EMG:</span>
                  <select 
                    value={visibleEmgMuscle} 
                    onChange={e => setVisibleEmgMuscle(e.target.value)}
                    className="text-xs font-bold border border-slate-200 rounded px-2 py-1 bg-slate-50 text-indigo-700 outline-none cursor-pointer"
                  >
                    {MUSCLE_LIST.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
              </div>

              <div className="space-y-6" style={{ cursor: draggingMarker ? 'col-resize' : 'default' }}>
                <div>
                  <div className="h-[220px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={analysisResult.chartData} syncId="liftingSync" onMouseDown={handleChartMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleChartMouseUp}>
                          <defs><linearGradient id="emgFill" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#4f46e5" stopOpacity={0.2}/><stop offset="95%" stopColor="#4f46e5" stopOpacity={0}/></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="time" type="number" domain={['dataMin', 'dataMax']} hide />
                          <YAxis tick={{fontSize: 10}} width={40} />
                          <Tooltip contentStyle={{fontSize:'12px'}} labelFormatter={(l)=>`Time: ${l}s`} />
                          {analysisResult.cycles.flatMap((cycle, idx) => {
                            const elements = [];
                            if (selectedRepIdx === idx) {
                              elements.push(<ReferenceLine key={`peak-emg-${idx}`} x={cycle.tPeak} stroke="#ef4444" strokeWidth={3} style={{ cursor: 'col-resize' }} label={{value:'Peak', fill:'#ef4444', fontSize:10}} />);
                              elements.push(<ReferenceLine key={`start-emg-${idx}`} x={cycle.tStart} stroke="#3b82f6" strokeWidth={3} style={{ cursor: 'col-resize' }} label={{value:'Start', fill:'#3b82f6', fontSize:10}} />);
                              elements.push(<ReferenceLine key={`end-emg-${idx}`} x={cycle.tEnd} stroke="#10b981" strokeWidth={3} style={{ cursor: 'col-resize' }} label={{value:'End', fill:'#10b981', fontSize:10}} />);
                            }
                            return elements;
                          })}
                          <Area type="monotone" dataKey={`emg_${visibleEmgMuscle}`} stroke="#4f46e5" fill="url(#emgFill)" strokeWidth={2} isAnimationActive={false} name={`EMG (${visibleEmgMuscle})`} />
                        </AreaChart>
                      </ResponsiveContainer>
                  </div>
                </div>
                <div>
                  <div className="h-[280px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={analysisResult.chartData} syncId="liftingSync" onMouseDown={handleChartMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleChartMouseUp}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="time" type="number" domain={['dataMin', 'dataMax']} tick={{fontSize: 10}} label={{value:'Time (s)', position:'insideBottom', offset:-5, fontSize:10}} />
                        <YAxis domain={['auto', 'auto']} tick={{fontSize: 10}} width={40} />
                        <Tooltip contentStyle={{fontSize:'12px'}} labelFormatter={(l)=>`Time: ${l}s`} />
                        <ReferenceLine x={0} stroke="#94a3b8" strokeWidth={1.5} label={{value:'Trigger (t=0)', position:'insideBottomLeft', fill:'#94a3b8', fontSize:10}} />
                        {analysisResult.cycles.flatMap((cycle, idx) => {
                          const elements = [];
                          if (selectedRepIdx === idx) {
                            elements.push(<ReferenceLine key={`peak-kin-${idx}`} x={cycle.tPeak} stroke="#ef4444" strokeWidth={3} style={{ cursor: 'col-resize' }} />);
                            elements.push(<ReferenceLine key={`start-kin-${idx}`} x={cycle.tStart} stroke="#3b82f6" strokeWidth={3} style={{ cursor: 'col-resize' }} />);
                            elements.push(<ReferenceLine key={`end-kin-${idx}`} x={cycle.tEnd} stroke="#10b981" strokeWidth={3} style={{ cursor: 'col-resize' }} />);
                          }
                          return elements;
                        })}
                        <Line type="monotone" dataKey="angleMain" name={`主要角度`} stroke="#f59e0b" strokeWidth={3} dot={false} isAnimationActive={false} />
                        <Line type="monotone" dataKey="kin_RScapUpDownRotation" name="UpDown Rot" stroke="#10b981" strokeWidth={1.5} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                        <Line type="monotone" dataKey="kin_RScapAntPosTilt" name="AntPos Tilt" stroke="#0ea5e9" strokeWidth={1.5} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                        <Line type="monotone" dataKey="kin_RScapIntExtRotation" name="IntExt Rot" stroke="#8b5cf6" strokeWidth={1.5} strokeDasharray="3 3" dot={false} isAnimationActive={false} />
                        <Legend wrapperStyle={{fontSize: '11px', fontWeight: 'bold'}} verticalAlign="top" height={36}/>
                        <Brush dataKey="time" height={30} stroke="#94a3b8" fill="#f8fafc" travellerWidth={10} tickFormatter={(v) => `${v}s`} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
};

// --- 網球發球分析 (Tennis Serve Task) 模組 ---
const TennisServeAnalysis = ({ onBack, taskTennisServeData, setTaskTennisServeData, taskTennisServeAngleData, setTaskTennisServeAngleData }) => {
  const [emgFileResult, setEmgFileResult] = useState(null);
  const [emgHeaders, setEmgHeaders] = useState([]);
  
  const [kinFileResult, setKinFileResult] = useState(null);
  const [kinHeaders, setKinHeaders] = useState([]);

  const [errorMessage, setErrorMessage] = useState(null);
  const [toastMessage, setToastMessage] = useState(null);

  // Settings: 欄位對應
  const [emgHandCourseColIdx, setEmgHandCourseColIdx] = useState(7); 
  const [kinHTElevColIdx, setKinHTElevColIdx] = useState(0); 
  const [kinHTPlaneColIdx, setKinHTPlaneColIdx] = useState(1); 
  const [kinTrigColIdx, setKinTrigColIdx] = useState(2); 
  const [trigThresh, setTrigThresh] = useState(3.0); 

  // Settings: DSP 參數
  const [emgSR, setEmgSR] = useState(1500);
  const [kinSR, setKinSR] = useState(960);
  const [bpHigh, setBpHigh] = useState(20);
  const [bpLow, setBpLow] = useState(450);
  const [rmsWindowMs, setRmsWindowMs] = useState(20);
  
  // Settings: 判定參數
  const [baselineFrames, setBaselineFrames] = useState(400); 
  const [sdMultiplier, setSdMultiplier] = useState(5); 
  const [kinSpikeThresh, setKinSpikeThresh] = useState(50); // 突波閥值設定

  const [visibleEmgMuscle, setVisibleEmgMuscle] = useState(MUSCLE_LIST[0]);
  const [analysisResult, setAnalysisResult] = useState(null);
  
  const [draggingMarker, setDraggingMarker] = useState(null);

  const KIN_METRICS_LIST = [
    'ScapIntExtRotation', 'ScapUpDownRotation', 'ScapAntPosTilt',
    'GHPlaneOfElev', 'GHElevation', 'GHAxialRotation',
    'HTPlaneOfElev', 'HTElevation', 'HTAxialRotation',
    'UpperTxFlexExt', 'UpperTxSidebending', 'UpperTxRotating',
    'LowerTxFlexExt', 'LowerTxSidebending', 'LowerTxRotating'
  ];

  const handleEmgUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const { finalHeaders, trimmedColumns, interpolatedCount } = parseDataContent(e.target.result);
        setEmgHeaders(finalHeaders);
        setEmgFileResult(trimmedColumns);
        
        const findCol = (keyword) => finalHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(keyword.toLowerCase().replace(/\s+/g, '')));
        const guessHandCourse = findCol('handcoursert');
        
        setEmgHandCourseColIdx(guessHandCourse !== -1 ? guessHandCourse : Math.min(7, finalHeaders.length - 1));

        setErrorMessage(null);
        if (interpolatedCount > 0) {
          setToastMessage(`⚠️ 偵測到 ${interpolatedCount} 筆 EMG 遺失數據，已自動線性插值補齊！`);
          setTimeout(() => setToastMessage(null), 4000);
        }
      } catch (err) {
        setErrorMessage(`EMG 解析失敗: ${err.message}`);
      }
    };
    reader.readAsText(file);
  };

  const handleKinUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const { finalHeaders, trimmedColumns, interpolatedCount } = parseDataContent(e.target.result);
        setKinHeaders(finalHeaders);
        setKinFileResult(trimmedColumns);
        
        const findCol = (keyword) => finalHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(keyword.toLowerCase()));
        const guessHTElev = findCol('htelevation');
        const guessHTPlane = findCol('htplaneofelev');
        const guessTrig = findCol('trigger');
        
        setKinHTElevColIdx(guessHTElev !== -1 ? guessHTElev : 0);
        setKinHTPlaneColIdx(guessHTPlane !== -1 ? guessHTPlane : Math.min(1, finalHeaders.length - 1));
        setKinTrigColIdx(guessTrig !== -1 ? guessTrig : Math.min(2, finalHeaders.length - 1));
        
        setErrorMessage(null);
        if (interpolatedCount > 0) {
          setToastMessage(`⚠️ 偵測到 ${interpolatedCount} 筆 KINEMATIC 遺失數據，已自動線性插值補齊！`);
          setTimeout(() => setToastMessage(null), 4000);
        }
      } catch (err) {
        setErrorMessage(`Kinematic 解析失敗: ${err.message}`);
      }
    };
    reader.readAsText(file);
  };

  const generateServeMetrics = (events, emgDataArrays, localKinFileResult, localKinHeaders, kinTrigIdx, localKinSR, localEmgSR, emgTrigIdx) => {
    
    const getEmgIdx = (t) => emgTrigIdx + Math.round(t * localEmgSR);

    const calcPhaseMetrics = (rmsArr, startT, endT) => {
      const sIdx = Math.max(0, getEmgIdx(startT));
      const eIdx = Math.min(rmsArr.length, getEmgIdx(endT));
      if (sIdx >= eIdx) return { mean: '-', peak: '-', peakTime: '-' };
      
      let sum = 0, peak = -Infinity, peakIdx = -1;
      for (let i = sIdx; i < eIdx; i++) {
        sum += rmsArr[i];
        if (rmsArr[i] > peak) { peak = rmsArr[i]; peakIdx = i; }
      }
      return {
        mean: (sum / (eIdx - sIdx)).toFixed(4),
        peak: peak.toFixed(4),
        peakTime: ((peakIdx - emgTrigIdx) / localEmgSR).toFixed(3)
      };
    };

    const emgMetricsList = emgDataArrays.map(item => ({
      muscle: item.muscle,
      rmsEnvelope: item.rmsEnvelope,
      metrics: {
        cocking: calcPhaseMetrics(item.rmsEnvelope, events.start, events.minPlane),
        acceleration: calcPhaseMetrics(item.rmsEnvelope, events.minPlane, events.impact),
        deceleration: calcPhaseMetrics(item.rmsEnvelope, events.impact, events.maxPlane)
      }
    }));

    const getKinIdx = (t) => kinTrigIdx + Math.round(t * localKinSR);

    const kinIndicesToExtract = {
      'Start': getKinIdx(events.start),
      'MinPlane': getKinIdx(events.minPlane),
      'Impact': getKinIdx(events.impact),
      'MaxPlane': getKinIdx(events.maxPlane)
    };

    const kinMetricsData = KIN_METRICS_LIST.map(field => {
      const colIdx = localKinHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(field.toLowerCase()));
      if (colIdx === -1) {
        return { field, start: '-', minPlane: '-', impact: '-', maxPlane: '-' };
      }
      const colData = localKinFileResult[colIdx];
      const safeGet = (idx) => (idx >= 0 && idx < colData.length) ? colData[idx].toFixed(2) : '-';
      return {
        field,
        start: safeGet(kinIndicesToExtract['Start']),
        minPlane: safeGet(kinIndicesToExtract['MinPlane']),
        impact: safeGet(kinIndicesToExtract['Impact']),
        maxPlane: safeGet(kinIndicesToExtract['MaxPlane'])
      };
    });

    return { emgMetricsList, kinMetricsData };
  };

  const processTennisServeTask = () => {
    if (!emgFileResult || !kinFileResult) {
      setErrorMessage("請先載入 EMG 與 KINEMATIC 兩個檔案！");
      return;
    }
    
    setErrorMessage(null);
    setAnalysisResult(null);

    try {
      const emgHandCourseData = emgFileResult[emgHandCourseColIdx];
      const kinHTElevDataRaw = kinFileResult[kinHTElevColIdx];
      const kinHTPlaneDataRaw = kinFileResult[kinHTPlaneColIdx];
      const kinTriggerData = kinFileResult[kinTrigColIdx];

      // 🟢 套用突波濾波器以去除異常訊號跳動 (例如萬向鎖)
      const kinHTElevData = removeSpikes(kinHTElevDataRaw, kinSpikeThresh);
      const kinHTPlaneData = removeSpikes(kinHTPlaneDataRaw, kinSpikeThresh);

      // 1. 尋找 KINEMATIC 系統絕對起點 (t = 0)
      let kinTrigIdx = -1;
      for (let i = 0; i < kinTriggerData.length; i++) {
        if (kinTriggerData[i] > trigThresh) {
          kinTrigIdx = i;
          break;
        }
      }
      if (kinTrigIdx === -1) {
        throw new Error(`KIN 同步失敗：找不到大於 ${trigThresh} 的 Trigger 訊號。`);
      }

      // 2. 尋找 EMG 系統絕對起點 (t = 0) - 尋找 Time, s 跨越 0 的列數
      let emgTrigIdx = -1;
      const emgTimeColIdx = emgHeaders.findIndex(h => h.toLowerCase().replace(/[^a-z]/g, '').includes('time'));
      if (emgTimeColIdx !== -1) {
          const emgTimeData = emgFileResult[emgTimeColIdx];
          for (let i = 0; i < emgTimeData.length; i++) {
              if (emgTimeData[i] >= 0) { 
                  emgTrigIdx = i;
                  break;
              }
          }
      }
      
      // 若無 Time 欄位，才退回頻率推算
      if (emgTrigIdx === -1 || emgTrigIdx === 0) {
          emgTrigIdx = Math.round((kinTrigIdx / kinSR) * emgSR);
      }

      const baseStart = kinTrigIdx;
      const baseEnd = Math.min(baseStart + baselineFrames, kinHTElevData.length);
      if (baseEnd <= baseStart) throw new Error("基準線取樣失敗：Trigger 點後方資料不足。");

      const baselineMean = calcMean(kinHTElevData.slice(baseStart, baseEnd));
      const baselineSD = calcSD(kinHTElevData.slice(baseStart, baseEnd), baselineMean);
      const startThreshold = baselineMean + sdMultiplier * baselineSD;

      // 1. 尋找開始舉手 (Start)
      let idxKinStart = -1;
      for (let i = baseEnd; i < kinHTElevData.length; i++) {
        if (kinHTElevData[i] > startThreshold) {
          idxKinStart = i;
          break;
        }
      }
      if (idxKinStart === -1) throw new Error(`找不到「開始舉手」點！(HTElevation 未超過基準閥值 ${startThreshold.toFixed(2)})`);
      const tStart = (idxKinStart - kinTrigIdx) / kinSR;

      // 2. 尋找最小肩水平面 (MinPlane)
      let minPlaneVal1 = Infinity;
      let idxKinMinPlane = -1;
      for (let i = idxKinStart; i < kinHTPlaneData.length; i++) {
        if (kinHTPlaneData[i] < minPlaneVal1) {
          minPlaneVal1 = kinHTPlaneData[i];
          idxKinMinPlane = i;
        }
      }
      if (idxKinMinPlane === -1) throw new Error("找不到「最小肩水平面角度」點！");
      const tMinPlane = (idxKinMinPlane - kinTrigIdx) / kinSR;

      // 3. 尋找擊球點 (Impact)
      const handCourseStartSearchIdx = emgTrigIdx + Math.round(tMinPlane * emgSR);
      if (handCourseStartSearchIdx < 0 || handCourseStartSearchIdx >= emgHandCourseData.length) {
        throw new Error(`尋找擊球點失敗：時間超出 Hand Course 數據長度。`);
      }

      let minHandVal = Infinity;
      let idxHandImpact = -1;
      const searchLimitImpact = emgHandCourseData.length; 
      
      for (let i = handCourseStartSearchIdx; i < searchLimitImpact; i++) {
        if (emgHandCourseData[i] < minHandVal) {
          minHandVal = emgHandCourseData[i];
          idxHandImpact = i;
        }
      }
      if (idxHandImpact === -1) throw new Error("找不到「擊球點」！請確認 Hand Course 訊號格式。");
      const tImpact = (idxHandImpact - emgTrigIdx) / emgSR;

      // 4. 尋找最大肩水平面 (MaxPlane)
      const kinStartSearchMaxPlane = kinTrigIdx + Math.round(tImpact * kinSR); 
      if (kinStartSearchMaxPlane >= kinHTPlaneData.length) {
        throw new Error(`尋找最大肩水平面角度失敗：擊球點時間 (${tImpact.toFixed(2)}s) 超出 KIN 數據長度。`);
      }

      let maxPlaneVal2 = -Infinity;
      let idxKinMaxPlane = -1;
      const searchLimitMaxPlane = kinHTPlaneData.length; 
      
      for (let i = kinStartSearchMaxPlane; i < searchLimitMaxPlane; i++) {
        if (kinHTPlaneData[i] > maxPlaneVal2) {
          maxPlaneVal2 = kinHTPlaneData[i];
          idxKinMaxPlane = i;
        }
      }
      if (idxKinMaxPlane === -1) throw new Error("找不到「最大肩水平面角度」點！");
      const tMaxPlane = (idxKinMaxPlane - kinTrigIdx) / kinSR;

      const events = {
        start: Number(tStart.toFixed(3)),
        minPlane: Number(tMinPlane.toFixed(3)),
        impact: Number(tImpact.toFixed(3)),
        maxPlane: Number(tMaxPlane.toFixed(3)),
        startThreshold: startThreshold
      };

      const emgDataArrays = MUSCLE_LIST.map((muscle, idx) => {
        let colIdx = emgHeaders.findIndex(h => h.toLowerCase().replace(/\s+/g, '').includes(muscle.toLowerCase()));
        if (colIdx === -1) colIdx = Math.min(idx, emgFileResult.length - 1);

        const emgData = emgFileResult[colIdx];
        // 確保順序：20~450Hz Bandpass (Zero-Lag) -> Rectify -> Zero-Lag LPF (RMS)
        const filtered = bandpassFilter(emgData, bpHigh, bpLow, emgSR);
        const rectified = new Float64Array(filtered.length);
        for(let i = 0; i < filtered.length; i++) rectified[i] = Math.abs(filtered[i]);
        const windowSize = Math.max(1, Math.floor(emgSR * (rmsWindowMs / 1000)));
        const rmsEnvelope = calculateRMS(rectified, windowSize);

        return { muscle, rmsEnvelope };
      });

      const { emgMetricsList, kinMetricsData } = generateServeMetrics(events, emgDataArrays, kinFileResult, kinHeaders, kinTrigIdx, kinSR, emgSR, emgTrigIdx);

      const chartData = [];
      const drawKinStartIdx = 0; 
      const drawKinEndIdx = kinHTPlaneData.length - 1; 

      for (let i = drawKinStartIdx; i <= drawKinEndIdx; i++) {
        const t = (i - kinTrigIdx) / kinSR; 

        // 透過相對時間 t 進行雙對齊
        const matchingEmgIdx = emgTrigIdx + Math.round(t * emgSR);
        
        const dataPoint = {
          time: Math.round(t * 1000) / 1000,
          htElev: Math.round(kinHTElevData[i] * 100) / 100,
          htPlane: Math.round(kinHTPlaneData[i] * 100) / 100,
          handCourse: (matchingEmgIdx >= 0 && matchingEmgIdx < emgHandCourseData.length) ? Math.round(emgHandCourseData[matchingEmgIdx] * 100) / 100 : null
        };

        MUSCLE_LIST.forEach((m, mIdx) => {
          const rms = emgDataArrays[mIdx].rmsEnvelope;
          dataPoint[`emg_${m}`] = (matchingEmgIdx >= 0 && matchingEmgIdx < rms.length) ? Math.round(rms[matchingEmgIdx] * 10000) / 10000 : null;
        });

        chartData.push(dataPoint);
      }

      setAnalysisResult({
        chartData,
        events,
        emgMetricsList,
        kinMetricsData,
        emgDataArrays,
        kinTrigIdx,
        emgTrigIdx
      });

    } catch (err) {
      setErrorMessage(`運算錯誤: ${err.message}`);
    }
  };

  const handleSaveData = () => {
    if (!analysisResult) return;

    const firstMuscle = MUSCLE_LIST[0];
    const currentDataCount = (taskTennisServeData[firstMuscle] || []).length;
    if (currentDataCount >= 3) {
      setToastMessage(`❌ 已達 3 次儲存上限！請先至資料庫刪除舊資料。`);
      setTimeout(() => setToastMessage(null), 3000);
      return;
    }

    setTaskTennisServeData(prev => {
      const newData = { ...prev };
      analysisResult.emgMetricsList.forEach(item => {
        if (!newData[item.muscle]) newData[item.muscle] = [];
        newData[item.muscle] = [...newData[item.muscle], {
          'Cocking': item.metrics.cocking.mean,
          'Acceleration': item.metrics.acceleration.mean,
          'Deceleration': item.metrics.deceleration.mean
        }];
      });
      return newData;
    });

    setTaskTennisServeAngleData(prev => {
      const newData = { ...prev };
      analysisResult.kinMetricsData.forEach(kin => {
        if (kin.field === '-' || !kin.field) return;
        if (!newData[kin.field]) newData[kin.field] = [];
        newData[kin.field] = [...newData[kin.field], {
          'Start': kin.start,
          'MinPlane': kin.minPlane,
          'Impact': kin.impact,
          'MaxPlane': kin.maxPlane
        }];
      });
      return newData;
    });

    setToastMessage(`✅ 成功儲存！已將發球 EMG 與 Kinematics 寫入資料庫 (第 ${currentDataCount + 1} 次)`);
    setTimeout(() => setToastMessage(null), 3000);
  };

  const handleChartMouseDown = useCallback((e) => {
    if (e && e.activePayload && analysisResult) {
      const time = e.activePayload[0].payload.time;
      const evs = analysisResult.events;
      const minD = Math.min(
        Math.abs(time - evs.start), 
        Math.abs(time - evs.minPlane), 
        Math.abs(time - evs.impact), 
        Math.abs(time - evs.maxPlane)
      );
      if (minD < 0.1) {
        if (minD === Math.abs(time - evs.start)) setDraggingMarker('start');
        else if (minD === Math.abs(time - evs.minPlane)) setDraggingMarker('minPlane');
        else if (minD === Math.abs(time - evs.impact)) setDraggingMarker('impact');
        else if (minD === Math.abs(time - evs.maxPlane)) setDraggingMarker('maxPlane');
      }
    }
  }, [analysisResult]);

  const handleChartMouseMove = useCallback((e) => {
    if (draggingMarker && e && e.activePayload && analysisResult) {
      const time = e.activePayload[0].payload.time;
      const newEvents = { ...analysisResult.events };
      
      if (draggingMarker === 'start') {
         if (time < newEvents.minPlane) newEvents.start = time;
      } else if (draggingMarker === 'minPlane') {
         if (time > newEvents.start && time < newEvents.impact) newEvents.minPlane = time;
      } else if (draggingMarker === 'impact') {
         if (time > newEvents.minPlane && time < newEvents.maxPlane) newEvents.impact = time;
      } else if (draggingMarker === 'maxPlane') {
         if (time > newEvents.impact) newEvents.maxPlane = time;
      }

      const { emgMetricsList, kinMetricsData } = generateServeMetrics(
        newEvents, 
        analysisResult.emgDataArrays, 
        kinFileResult, 
        kinHeaders, 
        analysisResult.kinTrigIdx, 
        kinSR, 
        emgSR,
        analysisResult.emgTrigIdx
      );

      setAnalysisResult({
        ...analysisResult,
        events: newEvents,
        emgMetricsList,
        kinMetricsData
      });
    }
  }, [draggingMarker, analysisResult, kinFileResult, kinHeaders, kinSR, emgSR]);

  const handleChartMouseUp = useCallback(() => setDraggingMarker(null), []);

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500 relative" onMouseUp={handleChartMouseUp} onMouseLeave={handleChartMouseUp}>
      {toastMessage && (
        <div className="fixed top-8 left-1/2 transform -translate-x-1/2 z-50 bg-slate-800 text-white px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-3 animate-in slide-in-from-top-4 duration-300">
          <span className="font-bold text-sm">{toastMessage}</span>
        </div>
      )}

      <header className="max-w-7xl mx-auto flex flex-col xl:flex-row justify-between items-start xl:items-center bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6 gap-4">
        <div className="flex items-center gap-4 shrink-0">
          <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500 hover:text-slate-800">
            <ArrowLeft size={24} />
          </button>
          <div className="bg-blue-600 p-3 rounded-2xl shadow-lg">
            <Target className="text-white w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900">網球發球分析</h1>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
          <label className={`flex items-center gap-2 px-5 py-2.5 rounded-2xl transition-all shadow-sm cursor-pointer text-sm font-bold shrink-0 ${emgFileResult ? 'bg-indigo-100 text-indigo-700' : 'bg-indigo-600 hover:bg-indigo-700 text-white'}`}>
            <Upload size={18} /> {emgFileResult ? '已載入 EMG 檔' : '載入 EMG 檔'}
            <input type="file" className="hidden" accept=".csv,.txt" onChange={handleEmgUpload} />
          </label>
          <label className={`flex items-center gap-2 px-5 py-2.5 rounded-2xl transition-all shadow-sm cursor-pointer text-sm font-bold shrink-0 ${kinFileResult ? 'bg-emerald-100 text-emerald-700' : 'bg-emerald-600 hover:bg-emerald-700 text-white'}`}>
            <Upload size={18} /> {kinFileResult ? '已載入 KINEMATIC 檔' : '載入 KINEMATIC 檔'}
            <input type="file" className="hidden" accept=".csv,.txt" onChange={handleKinUpload} />
          </label>
        </div>
      </header>

      <main className="max-w-7xl mx-auto space-y-6">
        {errorMessage && (
          <div className="bg-rose-50 border border-rose-200 text-rose-700 px-6 py-4 rounded-2xl font-bold flex items-center gap-3">
            <Info size={20} /> {errorMessage}
          </div>
        )}

        {emgFileResult && kinFileResult && (
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 space-y-4">
            <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <Link size={18} className="text-indigo-500" /> 訊號通道對應與分期設定
            </h3>
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="space-y-3 bg-indigo-50/50 p-4 rounded-2xl border border-indigo-100/50">
                <div className="text-xs font-bold text-indigo-800 mb-2 uppercase tracking-wide">EMG 設定</div>
                <div className="flex flex-col gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">肌肉分析通道 (RMS):</span>
                  <div className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-600 bg-white shadow-sm flex items-center gap-2">
                    <Activity size={14} className="text-indigo-400" />
                    自動解析: {MUSCLE_LIST.join(', ')}
                  </div>
                </div>
                <div className="flex items-center justify-between gap-2 border-t border-indigo-100/50 pt-2 mt-2">
                  <div className="flex flex-col gap-1 flex-1">
                    <span className="text-[11px] font-semibold text-slate-500">Hand Course (擊球低谷):</span>
                    <select value={emgHandCourseColIdx} onChange={e=>setEmgHandCourseColIdx(Number(e.target.value))} className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-700 bg-white">
                      {emgHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              <div className="space-y-3 bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                <div className="text-xs font-bold text-emerald-800 mb-2 uppercase tracking-wide">Kinematic 同步與關鍵點通道</div>
                <div className="flex items-center gap-2">
                  <div className="flex flex-col gap-1 flex-1">
                    <span className="text-[11px] font-semibold text-slate-500">Trigger 通道 (同步起點):</span>
                    <select value={kinTrigColIdx} onChange={e=>setKinTrigColIdx(Number(e.target.value))} className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-700">
                      {kinHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                    </select>
                  </div>
                  <div className="flex flex-col gap-1 w-20">
                    <span className="text-[11px] font-semibold text-slate-500">閥值(V):</span>
                    <input type="number" step="0.5" value={trigThresh} onChange={e=>setTrigThresh(Number(e.target.value))} className="p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center text-rose-600" />
                  </div>
                </div>
                <div className="flex flex-col gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">HTElevation (尋找開始舉手):</span>
                  <select value={kinHTElevColIdx} onChange={e=>setKinHTElevColIdx(Number(e.target.value))} className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-700">
                    {kinHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>
                <div className="flex flex-col gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">HTPlaneOfElev (尋找肩水平面特徵):</span>
                  <select value={kinHTPlaneColIdx} onChange={e=>setKinHTPlaneColIdx(Number(e.target.value))} className="p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-700">
                    {kinHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>
              </div>

              <div className="space-y-3 bg-amber-50/50 p-4 rounded-2xl border border-amber-100/50">
                <div className="text-xs font-bold text-amber-800 mb-2 uppercase tracking-wide">演算法判定設定</div>
                <div className="grid grid-cols-2 gap-x-2 gap-y-1.5">
                  <div>
                    <span className="text-[11px] font-semibold text-slate-500 mb-1 block">基準線取樣 (Trigger後 N筆):</span>
                    <input type="number" value={baselineFrames} onChange={e=>setBaselineFrames(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center bg-white" />
                  </div>
                  <div>
                    <span className="text-[11px] font-semibold text-slate-500 mb-1 block">判定閥值 (+SD):</span>
                    <input type="number" value={sdMultiplier} onChange={e=>setSdMultiplier(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-slate-200 text-sm font-bold text-center bg-white" />
                  </div>
                  <div className="col-span-2">
                    <span className="text-[11px] font-semibold text-slate-500 mb-1 block">突波濾除閾值 (deg/frame):</span>
                    <input type="number" value={kinSpikeThresh} onChange={e=>setKinSpikeThresh(Number(e.target.value))} className="w-full p-1.5 rounded-lg border border-amber-300 text-sm font-black text-amber-700 text-center bg-white" />
                  </div>
                  <div className="col-span-2 border-t border-amber-200/50 pt-2 mt-1 flex justify-between">
                    <span className="flex items-center gap-1 text-[10px] bg-white px-2 py-1 rounded shadow-sm border border-slate-100">EMG Hz <input type="number" value={emgSR} onChange={e=>setEmgSR(Number(e.target.value))} className="w-10 text-center font-bold text-indigo-600 outline-none"/></span>
                    <span className="flex items-center gap-1 text-[10px] bg-white px-2 py-1 rounded shadow-sm border border-slate-100">KIN Hz <input type="number" value={kinSR} onChange={e=>setKinSR(Number(e.target.value))} className="w-10 text-center font-bold text-emerald-600 outline-none"/></span>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t border-slate-100">
              <button onClick={processTennisServeTask} className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-xl font-bold transition-colors shadow-md active:scale-95 flex items-center gap-2">
                <Activity size={20} /> 執行分析
              </button>
            </div>
          </div>
        )}

        {analysisResult && (
          <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <MetricCard title="總動作時長" value={((analysisResult.events.maxPlane - analysisResult.events.start)).toFixed(2)} unit="s" icon={<Info className="text-slate-500" />} />
              <MetricCard title="揮臂期時長" value={((analysisResult.events.minPlane - analysisResult.events.start)).toFixed(2)} unit="s" icon={<Activity className="text-yellow-500" />} />
              <MetricCard title="加速期時長" value={((analysisResult.events.impact - analysisResult.events.minPlane)).toFixed(2)} unit="s" icon={<Waves className="text-rose-500" />} />
              <MetricCard title="減速期時長" value={((analysisResult.events.maxPlane - analysisResult.events.impact)).toFixed(2)} unit="s" icon={<Layers className="text-blue-500" />} />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6" style={{ cursor: draggingMarker ? 'col-resize' : 'default' }}>
              <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
                <div className="flex items-center justify-between mb-2 border-l-2 border-indigo-400 pl-2">
                  <h3 className="text-xs font-bold text-slate-500">EMG</h3>
                  <select 
                    value={visibleEmgMuscle} 
                    onChange={e => setVisibleEmgMuscle(e.target.value)}
                    className="text-[10px] font-bold border border-slate-200 rounded px-2 py-1 bg-slate-50 text-indigo-700 outline-none cursor-pointer"
                  >
                    {MUSCLE_LIST.map(m => <option key={m} value={m}>顯示圖表: {m}</option>)}
                  </select>
                </div>
                <div className="h-[250px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart data={analysisResult.chartData} syncId="serveSync" onMouseDown={handleChartMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleChartMouseUp}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="time" type="number" domain={['dataMin', 'dataMax']} tick={{fontSize: 10}} hide />
                      <YAxis tick={{fontSize: 10}} width={40} />
                      <Tooltip contentStyle={{fontSize:'12px', borderRadius:'12px'}} labelFormatter={(l)=>`Time: ${l}s`} />
                      
                      <ReferenceArea x1={analysisResult.events.start} x2={analysisResult.events.minPlane} fill="#fef08a" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.minPlane} x2={analysisResult.events.impact} fill="#fca5a5" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.impact} x2={analysisResult.events.maxPlane} fill="#bae6fd" fillOpacity={0.1} />
                      
                      <ReferenceLine x={0} stroke="#94a3b8" strokeWidth={1.5} label={{value:'Trigger (t=0)', position:'insideBottomLeft', fill:'#94a3b8', fontSize:10}} />
                      <ReferenceLine x={analysisResult.events.start} stroke="#10b981" strokeDasharray="3 3" label={{value:'Start', position:'insideTopLeft', fill:'#10b981', fontSize:10}} style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.minPlane} stroke="#eab308" strokeDasharray="3 3" label={{value:'MinPlane', position:'insideTopLeft', fill:'#eab308', fontSize:10}} style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.impact} stroke="#ef4444" strokeWidth={2} label={{value:'Impact', position:'insideTopLeft', fill:'#ef4444', fontSize:10}} style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.maxPlane} stroke="#8b5cf6" strokeDasharray="3 3" label={{value:'MaxPlane', position:'insideTopRight', fill:'#8b5cf6', fontSize:10}} style={{ cursor: 'col-resize' }} />

                      <Area type="monotone" dataKey={`emg_${visibleEmgMuscle}`} stroke="#4f46e5" fill="#e0e7ff" fillOpacity={0.5} isAnimationActive={false} name={`EMG (${visibleEmgMuscle})`} />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
                <div className="flex justify-center gap-4 mt-2 text-[10px] font-bold text-slate-500">
                  <span className="flex items-center gap-1"><div className="w-2 h-2 bg-indigo-500 rounded-full"></div> EMG {visibleEmgMuscle}</span>
                </div>
              </div>

              <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
                <h3 className="text-xs font-bold text-slate-500 mb-2 border-l-2 border-amber-500 pl-2">Kinematics 角度變化 (HTElev & HTPlane)</h3>
                <div className="h-[250px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={analysisResult.chartData} syncId="serveSync" onMouseDown={handleChartMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleChartMouseUp}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="time" type="number" domain={['dataMin', 'dataMax']} tick={{fontSize: 10}} label={{value:'Time (s)', position:'insideBottom', offset:-5, fontSize:10, fill:'#94a3b8'}} />
                      <YAxis tick={{fontSize: 10}} width={40} />
                      <Tooltip contentStyle={{fontSize:'12px', borderRadius:'12px'}} labelFormatter={(l)=>`Time: ${l}s`} />
                      
                      <ReferenceArea x1={analysisResult.events.start} x2={analysisResult.events.minPlane} fill="#fef08a" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.minPlane} x2={analysisResult.events.impact} fill="#fca5a5" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.impact} x2={analysisResult.events.maxPlane} fill="#bae6fd" fillOpacity={0.1} />

                      <ReferenceLine x={0} stroke="#94a3b8" strokeWidth={1.5} />
                      <ReferenceLine x={analysisResult.events.start} stroke="#10b981" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.minPlane} stroke="#eab308" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.impact} stroke="#ef4444" strokeWidth={2} style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.maxPlane} stroke="#8b5cf6" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />
                      <ReferenceLine y={analysisResult.events.startThreshold} stroke="#10b981" strokeDasharray="1 3" strokeWidth={1} label={{value:'Start Threshold', position:'insideBottomRight', fill:'#10b981', fontSize:9}} />

                      <Line type="monotone" dataKey="htElev" stroke="#10b981" dot={false} strokeWidth={2} isAnimationActive={false} name="HT Elevation" />
                      <Line type="monotone" dataKey="htPlane" stroke="#f59e0b" dot={false} strokeWidth={2} isAnimationActive={false} name="HT Plane Of Elev" />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
                 <div className="flex justify-center gap-4 mt-2 text-[10px] font-bold text-slate-500">
                  <span className="flex items-center gap-1"><div className="w-2 h-2 bg-emerald-500 rounded-full"></div> HTElevation</span>
                  <span className="flex items-center gap-1"><div className="w-2 h-2 bg-amber-500 rounded-full"></div> HTPlaneOfElev</span>
                </div>
              </div>

              <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 md:col-span-2">
                <h3 className="text-xs font-bold text-slate-500 mb-2 border-l-2 border-pink-500 pl-2">Hand Course RT 軌跡</h3>
                <div className="h-[250px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={analysisResult.chartData} syncId="serveSync" onMouseDown={handleChartMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleChartMouseUp}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="time" type="number" domain={['dataMin', 'dataMax']} tick={{fontSize: 10}} label={{value:'Time (s)', position:'insideBottom', offset:-5, fontSize:10, fill:'#94a3b8'}} />
                      <YAxis tick={{fontSize: 10}} width={40} domain={['auto', 'auto']} />
                      <Tooltip contentStyle={{fontSize:'12px', borderRadius:'12px'}} labelFormatter={(l)=>`Time: ${l}s`} />
                      
                      <ReferenceArea x1={analysisResult.events.start} x2={analysisResult.events.minPlane} fill="#fef08a" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.minPlane} x2={analysisResult.events.impact} fill="#fca5a5" fillOpacity={0.1} />
                      <ReferenceArea x1={analysisResult.events.impact} x2={analysisResult.events.maxPlane} fill="#bae6fd" fillOpacity={0.1} />

                      <ReferenceLine x={0} stroke="#94a3b8" strokeWidth={1.5} />
                      <ReferenceLine x={analysisResult.events.start} stroke="#10b981" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.minPlane} stroke="#eab308" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.impact} stroke="#ef4444" strokeWidth={2} style={{ cursor: 'col-resize' }} />
                      <ReferenceLine x={analysisResult.events.maxPlane} stroke="#8b5cf6" strokeDasharray="3 3" style={{ cursor: 'col-resize' }} />

                      <Line type="monotone" dataKey="handCourse" stroke="#ec4899" dot={false} strokeWidth={2.5} isAnimationActive={false} name="Hand Course RT" />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
                <div className="flex justify-center gap-4 mt-2 text-[10px] font-bold text-slate-500">
                  <span className="flex items-center gap-1"><div className="w-2 h-2 bg-pink-500 rounded-full"></div> Hand Course RT</span>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
              <div className="flex items-center justify-between mb-4">
                <h3 className="text-md font-bold text-slate-800 flex items-center gap-2">
                  <Activity size={18} className="text-indigo-500" /> 
                  EMG {MUSCLE_LIST.length}條肌肉全活化分期指標
                </h3>
              </div>
              <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-inner max-h-[400px]">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-100 text-slate-600 font-bold sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="p-4 border-b border-slate-300">目標肌肉</th>
                      <th className="p-4 border-b border-slate-300">發球時期 (Phase)</th>
                      <th className="p-4 border-b border-slate-300">起訖時間 (s)</th>
                      <th className="p-4 border-b border-slate-300">區間平均 (Mean RMS)</th>
                      <th className="p-4 border-b border-slate-300 text-rose-600">區間峰值 (Peak RMS)</th>
                      <th className="p-4 border-b border-slate-300">峰值發生時間 (s)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {analysisResult.emgMetricsList.map((item) => (
                      <React.Fragment key={item.muscle}>
                        <tr className="bg-slate-50/50 hover:bg-slate-100/50 transition-colors">
                          <td rowSpan={3} className="p-4 font-black text-slate-700 align-top border-r border-slate-200 border-b-2 border-b-slate-200">
                            {item.muscle}
                          </td>
                          <td className="p-3 text-yellow-700 font-bold text-xs"><span className="w-1.5 h-1.5 inline-block bg-yellow-400 rounded-full mr-2"></span>揮臂期 (Cocking)</td>
                          <td className="p-3 font-mono text-xs text-slate-500">{analysisResult.events.start} ~ {analysisResult.events.minPlane}</td>
                          <td className="p-3 font-mono font-bold text-slate-700">{item.metrics.cocking.mean}</td>
                          <td className="p-3 font-mono font-bold text-rose-600">{item.metrics.cocking.peak}</td>
                          <td className="p-3 font-mono text-slate-500">{item.metrics.cocking.peakTime}</td>
                        </tr>
                        <tr className="bg-white hover:bg-slate-50 transition-colors">
                          <td className="p-3 text-rose-700 font-bold text-xs"><span className="w-1.5 h-1.5 inline-block bg-rose-400 rounded-full mr-2"></span>加速期 (Acceleration)</td>
                          <td className="p-3 font-mono text-xs text-slate-500">{analysisResult.events.minPlane} ~ {analysisResult.events.impact}</td>
                          <td className="p-3 font-mono font-bold text-indigo-700 bg-indigo-50 px-2 rounded">{item.metrics.acceleration.mean}</td>
                          <td className="p-3 font-mono font-bold text-rose-600">{item.metrics.acceleration.peak}</td>
                          <td className="p-3 font-mono text-slate-500">{item.metrics.acceleration.peakTime}</td>
                        </tr>
                        <tr className="bg-white hover:bg-slate-50 transition-colors border-b-2 border-b-slate-200">
                          <td className="p-3 text-blue-700 font-bold text-xs"><span className="w-1.5 h-1.5 inline-block bg-blue-400 rounded-full mr-2"></span>減速期 (Deceleration)</td>
                          <td className="p-3 font-mono text-xs text-slate-500">{analysisResult.events.impact} ~ {analysisResult.events.maxPlane}</td>
                          <td className="p-3 font-mono font-bold text-slate-700">{item.metrics.deceleration.mean}</td>
                          <td className="p-3 font-mono font-bold text-rose-600">{item.metrics.deceleration.peak}</td>
                          <td className="p-3 font-mono text-slate-500">{item.metrics.deceleration.peakTime}</td>
                        </tr>
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
              <h3 className="text-md font-bold text-slate-800 flex items-center gap-2 mb-4">
                <Layers size={18} className="text-amber-500" /> 
                15 項關鍵運動學參數
              </h3>
              <div className="overflow-x-auto max-h-[500px] border border-slate-200 rounded-xl relative scroll-smooth shadow-inner">
                <table className="w-full text-left text-xs bg-white">
                  <thead className="bg-slate-100 text-slate-600 font-bold sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="p-3 border-b border-slate-300">Kinematic 參數名稱</th>
                      <th className="p-3 border-b border-slate-300 text-emerald-700">開始舉手<br/><span className="text-[10px] font-mono text-emerald-500 font-normal">t = {analysisResult.events.start}s</span></th>
                      <th className="p-3 border-b border-slate-300 text-amber-700">最小肩水平面角度<br/><span className="text-[10px] font-mono text-amber-500 font-normal">t = {analysisResult.events.minPlane}s</span></th>
                      <th className="p-3 border-b border-slate-300 text-rose-700">擊球點<br/><span className="text-[10px] font-mono text-rose-500 font-normal">t = {analysisResult.events.impact}s</span></th>
                      <th className="p-3 border-b border-slate-300 text-purple-700">最大肩水平面角度<br/><span className="text-[10px] font-mono text-purple-500 font-normal">t = {analysisResult.events.maxPlane}s</span></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {analysisResult.kinMetricsData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 font-bold text-slate-700">{row.field}</td>
                        <td className="p-3 font-mono">{row.start}</td>
                        <td className="p-3 font-mono">{row.minPlane}</td>
                        <td className="p-3 font-mono font-bold text-rose-600 bg-rose-50/30">{row.impact}</td>
                        <td className="p-3 font-mono">{row.maxPlane}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-blue-50 border border-blue-200 p-5 rounded-3xl flex flex-wrap items-center justify-between gap-4 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="bg-blue-100 p-2 rounded-xl text-blue-600">
                  <Save size={20} />
                </div>
                <div>
                  <h3 className="font-bold text-blue-800">一鍵寫入資料庫</h3>
                  <p className="text-xs text-blue-600 font-medium">將所有分析結果 (3階段 EMG + 4節點角度) 完整寫入總資料庫</p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <button onClick={handleSaveData} className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-2.5 rounded-xl font-bold transition-colors shadow-sm active:scale-95">
                  一鍵寫入資料庫
                </button>
              </div>
            </div>

          </div>
        )}
      </main>
    </div>
  );
};

// --- MVIC 分析模組 ---
const MvicAnalysis = ({ onBack, mvicData, setMvicData }) => {
  const [analysisResult, setAnalysisResult] = useState(null);
  const [activeDataPoint, setActiveDataPoint] = useState(null);
  const [errorMessage, setErrorMessage] = useState(null);
  const [toastMessage, setToastMessage] = useState(null); 

  const [saveTarget, setSaveTarget] = useState(MUSCLE_LIST[0]);
  const [samplingRate, setSamplingRate] = useState(1500); 
  const [analysisOffsetSec, setAnalysisOffsetSec] = useState(1); 
  const [analysisDurationSec, setAnalysisDurationSec] = useState(3); 

  const [bpHigh, setBpHigh] = useState(20);
  const [bpLow, setBpLow] = useState(450);
  const [rmsWindowMs, setRmsWindowMs] = useState(20);

  const [sdMultiplier, setSdMultiplier] = useState(5);
  const [consecutiveSamples, setConsecutiveSamples] = useState(1000);
  const [appliedBaseline, setAppliedBaseline] = useState(null);

  const [parsedFileResult, setParsedFileResult] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [selectedColumnIndex, setSelectedColumnIndex] = useState(1); 
  const [manualColInput, setManualColInput] = useState("2");
  const [previewRows, setPreviewRows] = useState([]);
  const [chartKey, setChartKey] = useState(0);
  const [onsetSample, setOnsetSample] = useState(0);
  const [isDragging, setIsDragging] = useState(false);

  const [isManualBaselineMode, setIsManualBaselineMode] = useState(false);
  const [manualBaseStart, setManualBaseStart] = useState(null);
  const [manualBaseEnd, setManualBaseEnd] = useState(null);
  const [isSelectingBase, setIsSelectingBase] = useState(false);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        setErrorMessage(null);
        setIsManualBaselineMode(false);
        setManualBaseStart(null);
        setManualBaseEnd(null);
        
        const { finalHeaders, trimmedColumns, validRowCount, interpolatedCount } = parseDataContent(e.target.result);

        setHeaders(finalHeaders);
        setParsedFileResult(trimmedColumns);

        const previewLimit = Math.min(1000, validRowCount);
        const rows = [];
        for (let r = 0; r < previewLimit; r++) {
          const row = [];
          for (let c = 0; c < trimmedColumns.length; c++) {
            row.push(trimmedColumns[c][r]); 
          }
          rows.push(row);
        }
        setPreviewRows(rows);

        const initialColIndex = trimmedColumns.length > 1 ? 1 : 0;
        setSelectedColumnIndex(initialColIndex);
        setManualColInput(initialColIndex + 1);
        setAppliedBaseline(null);
        processEMG(trimmedColumns[initialColIndex]);

        if (interpolatedCount > 0) {
          setToastMessage(`⚠️ 偵測到 ${interpolatedCount} 筆遺失數據，已自動線性插值補齊！`);
          setTimeout(() => setToastMessage(null), 4000);
        }

      } catch (err) {
        console.error(err);
        setErrorMessage(`檔案解析失敗: ${err.message}`);
        setAnalysisResult(null);
        setPreviewRows([]);
      }
    };
    reader.readAsText(file);
  };

  const handleAnalyzeClick = () => {
    const val = parseInt(manualColInput);
    if (isNaN(val) || val < 1 || val > (headers.length || 999)) {
      setErrorMessage(`請輸入有效的欄位數字`);
      return;
    }

    setAnalysisResult(null); 
    setErrorMessage(null);
    setChartKey(prev => prev + 1); 
    setIsManualBaselineMode(false);
    setManualBaseStart(null);
    setManualBaseEnd(null);
    setAppliedBaseline(null);

    setTimeout(() => {
      const targetIndex = val - 1;
      setSelectedColumnIndex(targetIndex);
      if (parsedFileResult && parsedFileResult[targetIndex]) {
        processEMG(parsedFileResult[targetIndex]);
      } else {
        setErrorMessage("找不到對應欄位的數據！");
      }
    }, 50); 
  };

  const processEMG = (data, customBaseline = null) => {
    if (!data || data.length < 10) {
      setErrorMessage("該欄位數據點過少 (< 10)，無法進行 DSP 運算！");
      setAnalysisResult(null);
      return;
    }
    setErrorMessage(null);

    const filtered = bandpassFilter(data, bpHigh, bpLow, samplingRate);
    const rectified = new Float64Array(filtered.length);
    for(let i = 0; i < filtered.length; i++) rectified[i] = Math.abs(filtered[i]);
    
    const windowSize = Math.max(1, Math.floor(samplingRate * (rmsWindowMs / 1000)));
    const rmsEnvelope = calculateRMS(rectified, windowSize);

    let baseStart = 0;
    let baseEnd = Math.min(2000, rmsEnvelope.length);

    if (customBaseline) {
      baseStart = Math.min(customBaseline.start, customBaseline.end);
      baseEnd = Math.max(customBaseline.start, customBaseline.end);
      baseStart = Math.max(0, baseStart);
      baseEnd = Math.min(rmsEnvelope.length, baseEnd);
      if (baseStart === baseEnd) baseEnd = Math.min(baseStart + 10, rmsEnvelope.length);
    }

    const baselinePoints = baseEnd - baseStart;
    let sum = 0;
    for(let i = baseStart; i < baseEnd; i++) sum += rmsEnvelope[i];
    const baselineMean = sum / (baselinePoints || 1);
    
    let sumSqDiff = 0;
    for(let i = baseStart; i < baseEnd; i++) sumSqDiff += Math.pow(rmsEnvelope[i] - baselineMean, 2);
    const baselineSD = Math.sqrt(sumSqDiff / (baselinePoints || 1));
    const threshold = baselineMean + sdMultiplier * baselineSD;

    let initialOnset = -1;
    let overThresholdCount = 0;
    for (let i = baseEnd; i < rmsEnvelope.length; i++) {
      if (rmsEnvelope[i] > threshold) {
        overThresholdCount++;
        if (overThresholdCount >= consecutiveSamples) { 
          initialOnset = i - consecutiveSamples + 1;
          break;
        }
      } else {
        overThresholdCount = 0; 
      }
    }
    if (initialOnset === -1) initialOnset = baseEnd;

    setOnsetSample(initialOnset);

    const chartData = [];
    const step = Math.max(1, Math.floor(data.length / 2000)); 
    
    for (let i = 0; i < data.length; i += step) {
      chartData.push({
        sample: i,
        original: Math.round(data[i] * 10000) / 10000,
        processed: Math.round(rectified[i] * 10000) / 10000,
        rms: Math.round(rmsEnvelope[i] * 10000) / 10000,
      });
    }

    setAnalysisResult({
      chartData,
      fullRms: rmsEnvelope,
      baselineMean: baselineMean,
      threshold: Math.round(threshold * 10000) / 10000
    });
  };

  const displayMetrics = useMemo(() => {
    if (!analysisResult || !analysisResult.fullRms) return null;
    
    const startAnalysis = onsetSample + Math.floor(samplingRate * analysisOffsetSec);
    const endAnalysis = startAnalysis + Math.floor(samplingRate * analysisDurationSec);
    
    const safeStart = Math.max(0, startAnalysis);
    const safeEnd = Math.max(safeStart, Math.min(endAnalysis, analysisResult.fullRms.length - 1));
    
    if (safeEnd <= safeStart) return null;
    
    const stableWindow = analysisResult.fullRms.slice(safeStart, safeEnd);
    if (stableWindow.length === 0) return null;

    const meanRMS = stableWindow.reduce((a, b) => a + b, 0) / stableWindow.length;
    const peakRMS = Math.max(...stableWindow);
    const sdRMS = Math.sqrt(stableWindow.reduce((s, v) => s + Math.pow(v - meanRMS, 2), 0) / stableWindow.length);
    const cv = meanRMS > 0 ? (sdRMS / meanRMS) * 100 : 0;
    const snr = 20 * Math.log10(meanRMS / (analysisResult.baselineMean || 0.001));

    return {
      meanRMS: meanRMS.toFixed(4),
      peakRMS: peakRMS.toFixed(4),
      cv: cv.toFixed(2),
      snr: snr.toFixed(2),
      startAnalysis: safeStart,
      endAnalysis: safeEnd
    };
  }, [analysisResult, onsetSample, samplingRate, analysisOffsetSec, analysisDurationSec]);

  const handleSaveMvicData = () => {
    if (!displayMetrics) return;
    
    if (mvicData[saveTarget].length >= 3) {
      setToastMessage(`❌ 【${saveTarget}】已達 3 次測試儲存上限！請先至資料庫刪除舊資料。`);
      setTimeout(() => setToastMessage(null), 3000);
      return;
    }

    const valueToSave = parseFloat(displayMetrics.meanRMS);
    setMvicData(prev => ({
      ...prev,
      [saveTarget]: [...prev[saveTarget], valueToSave]
    }));
    
    setToastMessage(`✅ 成功儲存！目標肌肉：${saveTarget}，數值：${valueToSave} mV`);
    setTimeout(() => setToastMessage(null), 3000);
  };

  const handleMouseDown = useCallback((e) => {
    if (e && e.activePayload) {
      const clickX = e.activePayload[0].payload.sample;
      if (isManualBaselineMode) {
        setManualBaseStart(clickX);
        setManualBaseEnd(clickX);
        setIsSelectingBase(true);
      } else {
        const tolerance = analysisResult.chartData.length > 0 ? analysisResult.chartData[analysisResult.chartData.length - 1].sample * 0.05 : 100;
        if (Math.abs(clickX - onsetSample) < tolerance) {
          setIsDragging(true);
        }
      }
    }
  }, [onsetSample, analysisResult, isManualBaselineMode]);

  const handleChartMouseMove = useCallback((state) => {
    if (state && state.activePayload) {
      setActiveDataPoint(state.activePayload[0].payload);
      const currentX = state.activePayload[0].payload.sample;
      
      if (isManualBaselineMode && isSelectingBase) {
        setManualBaseEnd(currentX);
      } else if (isDragging) {
        setOnsetSample(Math.max(0, currentX));
      }
    }
  }, [isDragging, isManualBaselineMode, isSelectingBase]);

  const handleMouseUp = useCallback(() => {
    setIsDragging(false);
    setIsSelectingBase(false);
  }, []);

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500 relative" onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp}>
      
      {toastMessage && (
        <div className="fixed top-8 left-1/2 transform -translate-x-1/2 z-50 bg-slate-800 text-white px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-3 animate-in slide-in-from-top-4 duration-300">
          <span className="font-bold text-sm">{toastMessage}</span>
        </div>
      )}

      <header className="max-w-7xl mx-auto flex flex-col xl:flex-row justify-between items-start xl:items-center bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6 gap-4">
        <div className="flex items-center gap-4 shrink-0">
          <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500 hover:text-slate-800">
            <ArrowLeft size={24} />
          </button>
          <div className="bg-indigo-600 p-3 rounded-2xl shadow-lg">
            <Activity className="text-white w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900">MVIC 基準分析</h1>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
          <div className="flex items-center bg-indigo-50 border border-indigo-100 px-3 py-1.5 rounded-xl shadow-sm">
            <span className="text-xs font-bold text-indigo-600 mr-2 shrink-0">分析第幾欄:</span>
            <input 
              type="number" 
              value={manualColInput} 
              onChange={(e) => setManualColInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAnalyzeClick();
              }}
              className="w-12 bg-transparent text-sm font-black text-indigo-800 focus:outline-none text-center border-b border-indigo-300 border-dashed"
              placeholder="輸入"
            />
            <span className="text-[10px] text-indigo-400 ml-2 font-mono whitespace-nowrap truncate max-w-[100px]">({headers[selectedColumnIndex] || '未載入檔案'})</span>
            <button
              onClick={handleAnalyzeClick}
              className="ml-3 shrink-0 bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1 rounded-lg text-xs font-bold transition-all shadow-sm active:scale-95"
            >
              分析
            </button>
          </div>

          <div className="flex items-center bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl hidden md:flex">
            <span className="text-xs font-semibold text-slate-500 mr-2 shrink-0">Bandpass:</span>
            <input 
              type="number" 
              value={bpHigh} 
              onChange={(e) => setBpHigh(Number(e.target.value))}
              className="w-10 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none text-center"
              title="高通頻率 (Hz)"
            />
            <span className="text-xs text-slate-400 mx-1">-</span>
            <input 
              type="number" 
              value={bpLow} 
              onChange={(e) => setBpLow(Number(e.target.value))}
              className="w-10 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none text-center"
              title="低通頻率 (Hz)"
            />
            <span className="text-xs text-slate-400 ml-1">Hz</span>
        </div>

        <div className="flex items-center bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl hidden md:flex">
          <span className="text-xs font-semibold text-slate-500 mr-2 shrink-0">Lowpass:</span>
          <input 
            type="number" 
            value={rmsWindowMs} 
            onChange={(e) => setRmsWindowMs(Number(e.target.value))}
            className="w-10 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none text-center"
            title="RMS 移動窗口大小"
          />
          <span className="text-xs text-slate-400 ml-1">ms</span>
        </div>

        <div className="flex items-center bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl hidden md:flex">
          <span className="text-xs font-semibold text-slate-500 mr-2 shrink-0">採樣率:</span>
            <input 
              type="number" 
              value={samplingRate} 
              onChange={(e) => setSamplingRate(parseInt(e.target.value))}
              className="w-14 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none"
            />
          </div>
          
          <div className="flex items-center bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl hidden md:flex">
            <span className="text-xs font-semibold text-slate-500 mr-2 shrink-0">延遲(s):</span>
            <input 
              type="number" 
              step="0.5"
              value={analysisOffsetSec} 
              onChange={(e) => setAnalysisOffsetSec(parseFloat(e.target.value) || 0)}
              className="w-10 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none"
            />
          </div>

          <div className="flex items-center bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl hidden md:flex">
            <span className="text-xs font-semibold text-slate-500 mr-2 shrink-0">時長(s):</span>
            <input 
              type="number" 
              step="0.5"
              value={analysisDurationSec} 
              onChange={(e) => setAnalysisDurationSec(parseFloat(e.target.value) || 0)}
              className="w-10 bg-transparent text-sm font-bold text-indigo-600 focus:outline-none"
            />
          </div>

          <label className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-2xl transition-all shadow-md cursor-pointer text-sm font-bold shrink-0 xl:ml-auto">
            <Upload size={18} /> 載入數據
            <input type="file" className="hidden" accept=".csv,.txt" onChange={handleFileUpload} />
          </label>
        </div>
      </header>

      <main className="max-w-7xl mx-auto space-y-6">
        
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <MetricCard title={`中間 ${analysisDurationSec}s 平均 (Mean RMS)`} value={displayMetrics?.meanRMS || '--'} unit="mV" icon={<BarChart className="text-blue-500" />} />
          <MetricCard title="分析區間峰值 (Peak RMS)" value={displayMetrics?.peakRMS || '--'} unit="mV" icon={<Activity className="text-rose-500" />} />
          <MetricCard title="變異係數 (CV)" value={displayMetrics ? `${displayMetrics.cv}%` : '--'} unit="" icon={<ShieldCheck className={(displayMetrics && parseFloat(displayMetrics.cv) > 15) ? "text-amber-500" : "text-emerald-500"} />} />
          <MetricCard title="訊雜比 (SNR)" value={displayMetrics ? `${displayMetrics.snr}` : '--'} unit="dB" icon={<Info className="text-indigo-500" />} />
        </div>

        {analysisResult && displayMetrics && (
          <div className="bg-emerald-50 border border-emerald-200 p-5 rounded-3xl flex flex-wrap items-center justify-between gap-4 shadow-sm">
            <div className="flex items-center gap-3">
              <div className="bg-emerald-100 p-2 rounded-xl text-emerald-600">
                <Save size={20} />
              </div>
              <div>
                <h3 className="font-bold text-emerald-800">儲存分析結果</h3>
                <p className="text-xs text-emerald-600 font-medium">將當前的 Mean RMS 數值寫入歷史資料庫</p>
              </div>
            </div>
            
            <div className="flex flex-wrap items-center gap-3">
              <span className="text-sm font-bold text-emerald-700">目標肌肉:</span>
              <select
                value={saveTarget}
                onChange={e => setSaveTarget(e.target.value)}
                className="px-4 py-2 rounded-xl border border-emerald-300 bg-white font-bold text-emerald-800 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
              >
                {MUSCLE_LIST.map(m => (
                  <option key={m} value={m}>{m} (已存 {mvicData[m].length}/3 次)</option>
                ))}
              </select>
              <button
                onClick={handleSaveMvicData}
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-2 rounded-xl font-bold transition-colors shadow-sm active:scale-95 flex items-center gap-2"
              >
                寫入資料庫
              </button>
            </div>
          </div>
        )}

        {analysisResult && displayMetrics && (
          <div className="flex flex-col sm:flex-row items-center justify-between bg-indigo-50/80 px-5 py-3 rounded-2xl border border-indigo-100 shadow-sm mt-2">
            <span className="text-sm font-bold text-indigo-800 flex items-center gap-2">
              <Crosshair size={16} /> 
              基準點定位與分析區間
            </span>
            <div className="flex items-center gap-4 text-xs font-bold text-indigo-600 mt-2 sm:mt-0">
              <span className="flex items-center gap-1">
                啟動閥值起點: <span className="text-sm bg-white px-2 py-0.5 rounded shadow-sm">{onsetSample}</span>
              </span>
              <span className="text-indigo-300">|</span>
              <span className="flex items-center gap-1 text-slate-600">
                動態分析區間 (+{analysisOffsetSec}s ~ +{analysisOffsetSec + analysisDurationSec}s): 
                <span className="font-mono bg-slate-100 px-2 py-0.5 rounded shadow-sm border border-slate-200">
                  {displayMetrics.startAnalysis} ~ {displayMetrics.endAnalysis}
                </span>
              </span>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between mb-2 h-8 overflow-hidden">
              <h3 className="text-md font-bold text-slate-700 flex items-center gap-2 flex-1 min-w-0 whitespace-nowrap overflow-hidden text-ellipsis">
                <Waves size={18} className="text-slate-400 shrink-0" />
                <span className="truncate">(1) 原始未處理信號 {headers.length > 0 ? `- ${headers[selectedColumnIndex]}` : ''}</span>
              </h3>
              <div className="text-[11px] font-mono text-slate-500 w-[150px] text-right shrink-0 ml-2">
                {activeDataPoint ? `X: ${activeDataPoint.sample} | Y: ${activeDataPoint.original} mV` : '\u00A0'}
              </div>
            </div>
            
            <div className="h-[300px] w-full" style={{ cursor: isManualBaselineMode ? 'crosshair' : (isDragging ? 'col-resize' : 'default') }}>
              {analysisResult ? (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={analysisResult.chartData} onMouseDown={handleMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleMouseUp} syncId={`emgSync-${chartKey}`}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="sample" type="number" domain={['dataMin', 'dataMax']} hide />
                    <YAxis axisLine={false} tick={{fontSize: 10}} />
                    <Tooltip content={<SimpleTooltip dataKey="original" label="原始" color="#94a3b8" />} />
                    
                    {displayMetrics && (
                      <ReferenceArea x1={displayMetrics.startAnalysis} x2={displayMetrics.endAnalysis} fill="#4f46e5" fillOpacity={0.05} />
                    )}
                    {isManualBaselineMode && manualBaseStart !== null && manualBaseEnd !== null && (
                      <ReferenceArea x1={manualBaseStart} x2={manualBaseEnd} fill="#f59e0b" fillOpacity={0.3} />
                    )}
                    <ReferenceLine x={onsetSample} stroke="#3b82f6" strokeWidth={2.5} style={{ cursor: 'ew-resize' }} label={{ value: `啟動起點`, position: 'insideTopLeft', fill: '#3b82f6', fontSize: 11, fontWeight: 'bold' }} />

                    <Line type="monotone" dataKey="original" stroke="#94a3b8" strokeWidth={1} dot={false} isAnimationActive={true} animationDuration={500} />
                  </LineChart>
                </ResponsiveContainer>
              ) : <Placeholder error={errorMessage} />}
            </div>
          </div>

          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between mb-2 h-8 overflow-hidden">
              <h3 className="text-md font-bold text-slate-700 flex items-center gap-2 flex-1 min-w-0 whitespace-nowrap overflow-hidden text-ellipsis">
                <Layers size={18} className="text-indigo-600 shrink-0" />
                <span className="truncate">(2) 濾波整流與 RMS 包絡線</span>
              </h3>
              <div className="text-[11px] font-mono text-indigo-600 font-bold w-[150px] text-right shrink-0 ml-2">
                {activeDataPoint ? `X: ${activeDataPoint.sample} | RMS: ${activeDataPoint.rms} mV` : '\u00A0'}
              </div>
            </div>

            <div className="h-[300px] w-full" style={{ cursor: isManualBaselineMode ? 'crosshair' : (isDragging ? 'col-resize' : 'default') }}>
              {analysisResult ? (
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={analysisResult.chartData} onMouseDown={handleMouseDown} onMouseMove={handleChartMouseMove} onMouseUp={handleMouseUp} syncId={`emgSync-${chartKey}`}>
                    <defs>
                      <linearGradient id="colorProc" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="#4f46e5" stopOpacity={0.1}/>
                        <stop offset="95%" stopColor="#4f46e5" stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="sample" type="number" domain={['dataMin', 'dataMax']} tick={{fontSize: 10}} />
                    <YAxis axisLine={false} tick={{fontSize: 10}} />
                    <Tooltip content={<SimpleTooltip dataKey="rms" label="RMS" color="#4f46e5" />} />
                    
                    {displayMetrics && (
                      <ReferenceArea x1={displayMetrics.startAnalysis} x2={displayMetrics.endAnalysis} fill="#4f46e5" fillOpacity={0.08} />
                    )}
                    {isManualBaselineMode && manualBaseStart !== null && manualBaseEnd !== null && (
                      <ReferenceArea x1={manualBaseStart} x2={manualBaseEnd} fill="#f59e0b" fillOpacity={0.3} />
                    )}
                    <ReferenceLine x={onsetSample} stroke="#3b82f6" strokeWidth={2.5} style={{ cursor: 'ew-resize' }} label={{ value: `啟動起點`, position: 'insideTopLeft', fill: '#3b82f6', fontSize: 11, fontWeight: 'bold' }} />
                    <ReferenceLine y={analysisResult.threshold} stroke="#ef4444" strokeDasharray="3 3" />

                    <Area type="monotone" dataKey="processed" stroke="#e2e8f0" fill="url(#colorProc)" dot={false} strokeWidth={1} isAnimationActive={true} animationDuration={500} />
                    <Line type="monotone" dataKey="rms" stroke="#4f46e5" strokeWidth={2} dot={false} isAnimationActive={true} animationDuration={500} />
                  </AreaChart>
                </ResponsiveContainer>
              ) : <Placeholder error={errorMessage} />}
            </div>
          </div>
        </div>

        {analysisResult && (
          <div className="bg-amber-50 border border-amber-200 p-5 rounded-3xl flex flex-col xl:flex-row items-start xl:items-center justify-between gap-4 shadow-sm transition-all duration-300">
            <div className="flex items-center gap-3">
              <div className="bg-amber-100 p-2 rounded-xl text-amber-600">
                <Settings2 size={20} />
              </div>
              <div>
                <h3 className="font-bold text-amber-800">進階定位設定 (手動補救)</h3>
                <p className="text-xs text-amber-600 font-medium">自訂啟動判定條件，或手動框選圖表上的靜態範圍作為新基準</p>
              </div>
            </div>
            
            <div className="flex flex-wrap items-center gap-3">
              <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-amber-200 shadow-sm">
                 <span className="text-xs font-bold text-amber-700">閥值: Mean +</span>
                 <input type="number" step="0.5" value={sdMultiplier} onChange={e => setSdMultiplier(Number(e.target.value))} className="w-14 text-center text-sm font-black text-amber-600 focus:outline-none" />
                 <span className="text-xs font-bold text-amber-700">× SD</span>
              </div>
              <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-amber-200 shadow-sm">
                 <span className="text-xs font-bold text-amber-700">需連續:</span>
                 <input type="number" value={consecutiveSamples} onChange={e => setConsecutiveSamples(Number(e.target.value))} className="w-16 text-center text-sm font-black text-amber-600 focus:outline-none" />
                 <span className="text-xs font-bold text-amber-700">筆</span>
              </div>
              
              <button 
                onClick={() => {
                  processEMG(parsedFileResult[selectedColumnIndex], appliedBaseline);
                  setToastMessage("✅ 參數已套用！圖表已重新尋找起點。");
                  setTimeout(() => setToastMessage(null), 3000);
                }}
                className="bg-amber-100 hover:bg-amber-200 text-amber-700 px-4 py-2 rounded-xl font-bold transition-colors shadow-sm active:scale-95 text-xs"
              >
                套用參數
              </button>

              {!isManualBaselineMode ? (
                <button
                  onClick={() => {
                    setIsManualBaselineMode(true);
                    setManualBaseStart(null);
                    setManualBaseEnd(null);
                  }}
                  className="bg-amber-500 hover:bg-amber-600 text-white px-5 py-2 rounded-xl font-bold transition-colors shadow-sm active:scale-95 text-sm xl:ml-2"
                >
                  手動框選新基準
                </button>
              ) : (
                <div className="flex items-center gap-3 bg-white p-1.5 rounded-2xl border border-amber-400 shadow-md animate-in zoom-in-95 duration-200 xl:ml-2">
                  <span className="text-xs font-bold text-amber-700 px-3">
                    {manualBaseStart !== null && manualBaseEnd !== null
                      ? `已選區間: ${Math.min(manualBaseStart, manualBaseEnd)} ~ ${Math.max(manualBaseStart, manualBaseEnd)}`
                      : '請在圖表拖曳...'}
                  </span>
                  <button
                    onClick={() => {
                      setIsManualBaselineMode(false);
                      setManualBaseStart(null);
                      setManualBaseEnd(null);
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-600 px-4 py-1.5 rounded-xl font-bold transition-colors text-xs"
                  >
                    取消
                  </button>
                  <button
                    onClick={() => {
                      if (manualBaseStart !== null && manualBaseEnd !== null) {
                        const newBaseline = { start: manualBaseStart, end: manualBaseEnd };
                        setAppliedBaseline(newBaseline);
                        processEMG(parsedFileResult[selectedColumnIndex], newBaseline);
                        setIsManualBaselineMode(false);
                        setToastMessage("✅ 基準重設成功！圖表已重新計算閥值並尋找起點。");
                        setTimeout(() => setToastMessage(null), 3000);
                      } else {
                        setToastMessage("⚠️ 請先在圖表上拖曳選取基準範圍！");
                        setTimeout(() => setToastMessage(null), 3000);
                      }
                    }}
                    className="bg-amber-600 hover:bg-amber-700 text-white px-4 py-1.5 rounded-xl font-bold transition-colors shadow-sm active:scale-95 text-xs"
                  >
                    套用選區並分析
                  </button>
                </div>
              )}
            </div>
          </div>
        )}

        {previewRows.length > 0 && (
          <div className="bg-white rounded-3xl shadow-sm border border-slate-100 p-6 overflow-hidden mt-6">
            <div className="flex items-center justify-between mb-4">
              <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                <Layers size={18} className="text-slate-400" />
                原始資料預覽（為維持效能，僅顯示前 1000 列）
              </h4>
              <span className="text-[11px] font-bold text-slate-400 bg-slate-50 px-2 py-1 rounded-lg">
                共 {headers.length} 欄，顯示前 {previewRows.length} 筆
              </span>
            </div>
            
            <div className="overflow-auto max-h-[400px] border border-slate-200 rounded-xl relative scroll-smooth bg-slate-50/30 shadow-inner">
              <table className="min-w-full text-xs text-slate-700 bg-white">
                <thead>
                  <tr>
                    <th className="sticky top-0 z-20 px-3 py-3 text-center font-bold text-slate-500 bg-slate-200 border-b border-slate-300 shadow-sm whitespace-nowrap">
                      Index
                    </th>
                    {headers.map((h, idx) => (
                      <th key={idx} className={`sticky top-0 z-20 px-3 py-3 text-left font-bold whitespace-nowrap border-b border-slate-300 shadow-sm ${idx === selectedColumnIndex ? 'text-indigo-700 bg-indigo-100' : 'bg-slate-100'}`}>
                        {h} {idx === selectedColumnIndex && '✓'}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {previewRows.map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-3 py-1.5 font-mono text-slate-400 text-center bg-slate-50/80 border-r border-slate-100">
                        {rIdx + 1}
                      </td>
                      {row.map((cell, cIdx) => (
                        <td key={cIdx} className={`px-3 py-1.5 font-mono whitespace-nowrap ${cIdx === selectedColumnIndex ? 'bg-indigo-50/20 text-indigo-700 font-bold' : ''}`}>
                          {cell}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="bg-slate-900 rounded-3xl p-8 text-white grid grid-cols-1 md:grid-cols-2 gap-8 mt-6">
          <div>
            <h4 className="text-indigo-400 font-bold text-xs uppercase tracking-widest mb-4">分析演算法說明</h4>
            <div className="space-y-3 text-sm text-slate-400">
              <p>• <b>處理流程</b>：原始信號 → 自訂帶通濾波 (Bandpass) → 全波整流 (絕對值) → 自訂 RMS 移動窗口平滑 (Lowpass)。</p>
              <p>• <b>自動定位</b>：取前 2000 個 Sample (或自訂區間) 作為靜態基準，以 $Mean + {sdMultiplier} \times SD$ 為啟動閥值，且需<b>連續 {consecutiveSamples} 筆</b>超過閥值才判定為啟動。圖表上的藍色粗線代表此起點。</p>
            </div>
          </div>
          <div>
            <h4 className="text-indigo-400 font-bold text-xs uppercase tracking-widest mb-4">動態參數對應 (LabVIEW 相容)</h4>
            <p className="text-sm text-slate-400 mb-4">已完整移植 LabVIEW 設定參數。系統會根據藍色起點線，加上上方設定的「延遲分析秒數」作為真正分析的起始點，並擷取「分析時長」作為穩定期計算基準。您可以拖曳起點線來校正整段區域。</p>
            <div className="flex gap-4">
              <div className="flex items-center gap-2 text-[10px] text-slate-400">
                <div className="w-3 h-1 bg-slate-400"></div> 原始數值
              </div>
              <div className="flex items-center gap-2 text-[10px] text-indigo-400 font-bold">
                <div className="w-3 h-1 bg-indigo-500"></div> RMS 包絡線
              </div>
              <div className="flex items-center gap-2 text-[10px] text-indigo-400 font-bold">
                <div className="w-3 h-3 bg-indigo-500 opacity-20"></div> 動態分析區間
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
};

// --- 擴充模組 Placeholder ---
const ModulePlaceholder = ({ title, icon, description, onBack }) => (
  <div className="min-h-screen bg-[#f1f5f9] p-6 font-sans text-slate-800 animate-in fade-in duration-500">
    <header className="max-w-7xl mx-auto flex items-center gap-4 bg-white p-6 rounded-3xl shadow-sm border border-slate-100 mb-6">
      <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-500"><ArrowLeft size={24} /></button>
      <div className="bg-blue-500 p-3 rounded-2xl shadow-lg text-white">{icon}</div>
      <div><h1 className="text-xl font-bold text-slate-900">{title}</h1></div>
    </header>
    <main className="max-w-7xl mx-auto bg-white h-[60vh] rounded-3xl shadow-sm border border-slate-100 flex flex-col items-center justify-center text-slate-400">
      <div className="bg-slate-50 p-8 rounded-full mb-4">{icon}</div><h2 className="text-xl font-bold text-slate-600 mb-2">模組建置中</h2><p className="text-sm">此為預留擴充區塊，未來可匯入專屬的 {title} 演算法。</p>
    </main>
  </div>
);

// --- 主應用程式 (Router 與狀態共享層) ---
const App = () => {
  const [currentView, setCurrentView] = useState('home');
  const [isExporting, setIsExporting] = useState(false);
  
  const initialDataObj = MUSCLE_LIST.reduce((acc, muscle) => ({ ...acc, [muscle]: [] }), {});
  
  const [mvicData, setMvicData] = useState(initialDataObj);
  // Lifting 獨立狀態
  const [taskLiftEmgData, setTaskLiftEmgData] = useState(initialDataObj);
  const [taskLiftAngleData, setTaskLiftAngleData] = useState({});
  
  const [taskTennisServeData, setTaskTennisServeData] = useState(initialDataObj);
  const [taskTennisServeAngleData, setTaskTennisServeAngleData] = useState({});

  const handleExportExcel = async () => {
    try {
      setIsExporting(true);
      const XLSX = await loadXLSX();
      const wb = XLSX.utils.book_new();

      // MVIC 匯出
      const mvicRows = [];
      MUSCLE_LIST.forEach(muscle => {
        const trials = mvicData[muscle];
        const mean = calcMean(trials);
        const sd = calcSD(trials, mean);
        mvicRows.push({
          Muscle: muscle,
          Trial_1: trials[0] !== undefined ? trials[0] : '',
          Trial_2: trials[1] !== undefined ? trials[1] : '',
          Trial_3: trials[2] !== undefined ? trials[2] : '',
          Mean: trials.length > 0 ? mean : '',
          SD: trials.length > 1 ? sd : ''
        });
      });
      const wsMvic = XLSX.utils.json_to_sheet(mvicRows);
      XLSX.utils.book_append_sheet(wb, wsMvic, "MVIC");

      // 動態展平多重區段的通用寫入函式
      const createMultiPhaseSheet = (dataObj, sheetName, defaultKeys, keyName, phases) => {
        const savedKeys = Object.keys(dataObj).filter(k => dataObj[k].length > 0);
        const keysToExport = savedKeys.length > 0 ? savedKeys : defaultKeys;
        
        const rows = keysToExport.map(key => {
          const trials = dataObj[key] || [];
          const row = { [keyName]: key };
          
          [0, 1, 2].forEach(tIdx => {
            const trial = trials[tIdx] || {};
            phases.forEach(phase => {
              row[`T${tIdx+1}_${phase}`] = trial[phase] !== undefined ? trial[phase] : '';
            });
          });

          // 各階段三重複的平均
          phases.forEach(phase => {
            const vals = trials.map(t => t[phase]).filter(v => v !== undefined && v !== '');
            row[`Mean_${phase}`] = vals.length > 0 ? calcMean(vals).toFixed(4) : '';
          });
          return row;
        });
        const ws = XLSX.utils.json_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      };
      
      // Lifting 匯出
      const liftPhases = ['Up_30-60', 'Up_60-90', 'Up_90-120', 'Down_120-90', 'Down_90-60', 'Down_60-30'];
      const liftAnglePhases = ['Up_30', 'Up_60', 'Up_90', 'Down_90', 'Down_60', 'Down_30'];
      createMultiPhaseSheet(taskLiftEmgData, "Lifting_EMG", MUSCLE_LIST, "Muscle", liftPhases);
      const liftAngleKeys = Object.keys(taskLiftAngleData).length > 0 ? Object.keys(taskLiftAngleData) : ['AngleChannel'];
      createMultiPhaseSheet(taskLiftAngleData, "Lifting_Angles", liftAngleKeys, "Channel", liftAnglePhases);

      // Tennis Serve 匯出
      const serveEmgPhases = ['Cocking', 'Acceleration', 'Deceleration'];
      const serveAnglePhases = ['Start', 'MinPlane', 'Impact', 'MaxPlane'];
      createMultiPhaseSheet(taskTennisServeData, "TennisServe_EMG", MUSCLE_LIST, "Muscle", serveEmgPhases);
      const serveAngleKeys = Object.keys(taskTennisServeAngleData).length > 0 ? Object.keys(taskTennisServeAngleData) : ['AngleChannel'];
      createMultiPhaseSheet(taskTennisServeAngleData, "TennisServe_Angles", serveAngleKeys, "Channel", serveAnglePhases);

      // Summary 匯出 (擷取特定指標做大彙整)
      const summaryRows = MUSCLE_LIST.map(muscle => {
        const liftTrials = taskLiftEmgData[muscle] || [];
        const allLiftVals = [];
        liftTrials.forEach(t => {
           liftPhases.forEach(p => { 
             if (t[p] !== undefined && t[p] !== '') allLiftVals.push(t[p]); 
           });
        });

        // 擷取網球加速期的平均
        const serveTrials = taskTennisServeData[muscle] || [];
        const allServeAccVals = [];
        serveTrials.forEach(t => {
           if (t['Acceleration'] !== undefined && t['Acceleration'] !== '') allServeAccVals.push(t['Acceleration']);
        });

        return {
          Muscle: muscle,
          MVIC_Mean: calcMean(mvicData[muscle]) || '',
          Lift_Overall_Mean_RMS: allLiftVals.length > 0 ? calcMean(allLiftVals).toFixed(4) : '',
          TennisServe_Acc_Mean_RMS: allServeAccVals.length > 0 ? calcMean(allServeAccVals).toFixed(4) : ''
        };
      });
      const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
      XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

      XLSX.writeFile(wb, "EMG_Research_Data_SPSS.xlsx");
    } catch (err) {
      alert("匯出失敗: " + err.message);
    } finally {
      setIsExporting(false);
    }
  };

  const renderHome = () => (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 p-6 md:p-10 animate-in fade-in duration-500">
      <header className="max-w-6xl mx-auto mb-12 text-center md:text-left">
        <h1 className="text-4xl font-black text-slate-900 tracking-tight flex items-center justify-center md:justify-start gap-3">
          <Activity className="text-indigo-600" size={36} /> EMG 科研整合平台
        </h1>
        <p className="text-slate-500 mt-2 font-medium">Musculoskeletal & Biomechanics Laboratory Center</p>
      </header>
      <main className="max-w-6xl mx-auto space-y-12">
        <section>
          <div className="flex items-center gap-2 mb-6 border-b border-slate-200 pb-2"><Activity size={20} className="text-indigo-600" /><h2 className="text-xl font-bold text-slate-800">信號分析模組</h2></div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
            <div onClick={() => setCurrentView('mvic')} className="bg-white p-6 rounded-3xl border border-slate-200 hover:border-indigo-400 hover:shadow-lg transition-all cursor-pointer group">
              <div className="bg-indigo-100 w-12 h-12 rounded-2xl flex items-center justify-center mb-4 group-hover:bg-indigo-600 transition-colors"><BarChart className="text-indigo-600 group-hover:text-white transition-colors" /></div>
              <h3 className="text-lg font-bold text-slate-900 mb-1">MVIC 分析</h3>
            </div>
            <div onClick={() => setCurrentView('task_lift')} className="bg-white p-6 rounded-3xl border border-slate-200 hover:border-blue-400 hover:shadow-lg transition-all cursor-pointer group">
              <div className="bg-blue-50 w-12 h-12 rounded-2xl flex items-center justify-center mb-4 group-hover:bg-blue-500 transition-colors"><ArrowUpRight className="text-blue-500 group-hover:text-white transition-colors" /></div>
              <h3 className="text-lg font-bold text-slate-900 mb-1">舉手動作分析</h3>
            </div>
            <div onClick={() => setCurrentView('task_tennis_serve')} className="bg-white p-6 rounded-3xl border border-slate-200 hover:border-blue-400 hover:shadow-lg transition-all cursor-pointer group">
              <div className="bg-blue-50 w-12 h-12 rounded-2xl flex items-center justify-center mb-4 group-hover:bg-blue-500 transition-colors"><Target className="text-blue-500 group-hover:text-white transition-colors" /></div>
              <h3 className="text-lg font-bold text-slate-900 mb-1">網球發球分析</h3>
            </div>
          </div>
        </section>
        <section>
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-6 border-b border-slate-200 pb-2 gap-4">
            <div className="flex items-center gap-2"><Database size={20} className="text-emerald-600" /><h2 className="text-xl font-bold text-slate-800">數據結果資料庫</h2></div>
            <button onClick={handleExportExcel} disabled={isExporting} className={`flex items-center gap-2 px-5 py-2.5 rounded-xl transition-all shadow-sm text-sm font-bold text-white ${isExporting ? 'bg-slate-400 cursor-not-allowed' : 'bg-emerald-600 hover:bg-emerald-700 active:scale-95'}`}>
              {isExporting ? <Activity className="animate-spin" size={18} /> : <FileSpreadsheet size={18} />} {isExporting ? '生成 Excel 中...' : '匯出 Excel (SPSS 格式相容)'}
            </button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div onClick={() => setCurrentView('result_mvic')} className="bg-slate-900 p-6 rounded-3xl border border-slate-800 hover:border-emerald-500 hover:shadow-xl transition-all cursor-pointer group relative overflow-hidden">
              <div className="absolute top-0 right-0 p-4 opacity-10"><Database size={100} /></div>
              <div className="flex justify-between items-center mb-4 relative z-10">
                <div className="bg-slate-800 w-12 h-12 rounded-2xl flex items-center justify-center group-hover:bg-emerald-500 transition-colors"><FolderOpen className="text-emerald-400 group-hover:text-white transition-colors" /></div>
                <span className="text-xs font-mono text-emerald-400 bg-emerald-900/50 px-2 py-1 rounded-md border border-emerald-800/50">MVIC Storage</span>
              </div>
              <h3 className="text-lg font-bold text-white mb-1 relative z-10">MVIC 歷史數據庫</h3><p className="text-xs text-slate-400 relative z-10">包含各肌肉三重複測試之 Mean RMS, 綜合平均與標準差計算。</p>
            </div>
            <div onClick={() => setCurrentView('result_task')} className="bg-slate-900 p-6 rounded-3xl border border-slate-800 hover:border-blue-500 hover:shadow-xl transition-all cursor-pointer group relative overflow-hidden">
              <div className="absolute top-0 right-0 p-4 opacity-10"><Activity size={100} /></div>
              <div className="flex justify-between items-center mb-4 relative z-10">
                <div className="bg-slate-800 w-12 h-12 rounded-2xl flex items-center justify-center group-hover:bg-blue-500 transition-colors"><Database className="text-blue-400 group-hover:text-white transition-colors" /></div>
                <span className="text-xs font-mono text-slate-500 bg-slate-800 px-2 py-1 rounded-md">Tasks Data</span>
              </div>
              <h3 className="text-lg font-bold text-white mb-1 relative z-10">任務數據總表</h3><p className="text-xs text-slate-400 relative z-10">包含舉手、網球發球等動態任務之分析成果與多區段彙整。</p>
            </div>
          </div>
        </section>
      </main>
    </div>
  );

  switch (currentView) {
    case 'home': return renderHome();
    case 'mvic': return <MvicAnalysis onBack={() => setCurrentView('home')} mvicData={mvicData} setMvicData={setMvicData} />;
    case 'result_mvic': return <MvicDatabase mvicData={mvicData} setMvicData={setMvicData} onBack={() => setCurrentView('home')} />;
    case 'task_lift': return <LiftingAnalysis onBack={() => setCurrentView('home')} taskLiftEmgData={taskLiftEmgData} setTaskLiftEmgData={setTaskLiftEmgData} taskLiftAngleData={taskLiftAngleData} setTaskLiftAngleData={setTaskLiftAngleData} />;
    case 'task_tennis_serve': return <TennisServeAnalysis onBack={() => setCurrentView('home')} taskTennisServeData={taskTennisServeData} setTaskTennisServeData={setTaskTennisServeData} taskTennisServeAngleData={taskTennisServeAngleData} setTaskTennisServeAngleData={setTaskTennisServeAngleData} />;
    case 'result_task': 
      return <TaskDatabase 
               onBack={() => setCurrentView('home')} 
               taskLiftEmgData={taskLiftEmgData} setTaskLiftEmgData={setTaskLiftEmgData} 
               taskLiftAngleData={taskLiftAngleData} setTaskLiftAngleData={setTaskLiftAngleData} 
               taskTennisServeData={taskTennisServeData} setTaskTennisServeData={setTaskTennisServeData}
               taskTennisServeAngleData={taskTennisServeAngleData} setTaskTennisServeAngleData={setTaskTennisServeAngleData}
             />;
    default: return renderHome();
  }
};

const MetricCard = ({ title, value, unit, icon }) => (
  <div className="bg-white p-4 rounded-3xl border border-slate-100 shadow-sm flex flex-col justify-center">
    <div className="p-1.5 bg-slate-50 w-fit rounded-lg mb-2">{icon}</div>
    <div className="flex items-baseline gap-1">
      <h3 className="text-xl font-black text-slate-900 font-mono leading-none">{value}</h3>
      <span className="text-[10px] font-bold text-slate-400 uppercase">{unit}</span>
    </div>
    <p className="text-[9px] font-bold text-slate-500 mt-1 uppercase tracking-wider line-clamp-1">{title}</p>
  </div>
);

const SimpleTooltip = ({ active, payload, dataKey, label, color }) => {
  if (active && payload && payload.length) {
    return (
      <div className="bg-white/90 backdrop-blur-md p-2 border border-slate-100 shadow-xl rounded-xl text-[10px]">
        <p className="font-bold text-slate-400">Time: {payload[0].payload.time || payload[0].payload.sample}s</p>
        <p style={{ color }} className="font-bold font-mono">{label}: {payload[0].payload[dataKey]} {dataKey.includes('angle')?'°':'mV'}</p>
      </div>
    );
  }
  return null;
};

const Placeholder = ({ error }) => (
  <div className="w-full h-full flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-3xl text-slate-400 italic text-sm p-4 text-center bg-slate-50/50">
    {error ? <span className="text-rose-500 font-bold">{error}</span> : "等待數據載入..."}
  </div>
);

export default App;
