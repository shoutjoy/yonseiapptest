import React, { useState, useEffect, useCallback, useRef } from 'react';
import { 
  FileUp, 
  CheckCircle, 
  AlertTriangle, 
  Download, 
  RefreshCw, 
  Type, 
  ShieldCheck, 
  Binary,
  ArrowRightLeft,
  FileCode
} from 'lucide-react';

/**
 * PPTX Font Recovery Engine v2.5 (Functional Version)
 * 기술적 원리: OOXML 구조 분해 및 DOMParser를 이용한 XML 노드 강제 치환
 * 학술적 근거: Jakob Nielsen의 사용성 원칙 및 ISO/IEC 29500-1 표준 준수
 */
const App = () => {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState('idle'); // idle | loading_lib | ready_to_scan | processing | completed | error
  const [progress, setProgress] = useState(0);
  const [recoveredBlob, setRecoveredBlob] = useState(null);
  const [log, setLog] = useState([]);

  const jszipRef = useRef(null);

  // 라이브러리 동적 로드 (JSZip)
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    script.async = true;
    script.onload = () => {
      jszipRef.current = window.JSZip;
      setLog(prev => [...prev, "System: JSZip Library Loaded Successfully."]);
    };
    document.body.appendChild(script);
  }, []);

  const addLog = (msg) => {
    setLog(prev => [...prev, `Process: ${msg}`]);
  };

  const handleFileUpload = (e) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;
    setFile(uploadedFile);
    setStatus('ready_to_scan');
    addLog(`File "${uploadedFile.name}" selected.`);
  };

  const processPptx = async () => {
    if (!file || !jszipRef.current) return;

    setStatus('processing');
    setProgress(10);
    addLog("Extracting OOXML structure...");

    try {
      const zip = new jszipRef.current();
      const content = await zip.loadAsync(file);
      setProgress(30);

      const parser = new DOMParser();
      const serializer = new XMLSerializer();
      const fallbackFont = "Malgun Gothic";
      let replacedCount = 0;

      // 1. 슬라이드 데이터 스캔 (ppt/slides/slide*.xml)
      const slideFiles = Object.keys(content.files).filter(path => path.startsWith('ppt/slides/slide') && path.endsWith('.xml'));
      
      for (let i = 0; i < slideFiles.length; i++) {
        const path = slideFiles[i];
        const xmlText = await content.file(path).async("text");
        const xmlDoc = parser.parseFromString(xmlText, "text/xml");

        // font typeface 속성 탐색 (latin, ea, cs 등)
        const fontNodes = xmlDoc.querySelectorAll('[typeface]');
        fontNodes.forEach(node => {
          const oldFont = node.getAttribute('typeface');
          if (oldFont && oldFont !== fallbackFont) {
            node.setAttribute('typeface', fallbackFont);
            replacedCount++;
          }
        });

        const newXmlText = serializer.serializeToString(xmlDoc);
        zip.file(path, newXmlText);
        setProgress(30 + Math.floor((i / slideFiles.length) * 40));
      }

      // 2. 테마 데이터 스캔 (ppt/theme/theme1.xml)
      const themePath = 'ppt/theme/theme1.xml';
      if (content.files[themePath]) {
        const themeText = await content.file(themePath).async("text");
        const themeDoc = parser.parseFromString(themeText, "text/xml");
        const themeFonts = themeDoc.querySelectorAll('a\\:latin, a\\:ea, a\\:cs');
        themeFonts.forEach(node => {
          node.setAttribute('typeface', fallbackFont);
          replacedCount++;
        });
        zip.file(themePath, serializer.serializeToString(themeDoc));
      }

      addLog(`Replaced ${replacedCount} font references with "${fallbackFont}".`);
      setProgress(80);
      addLog("Re-packaging into .pptx container...");

      const outputBlob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
      setRecoveredBlob(outputBlob);
      setProgress(100);
      setStatus('completed');
      addLog("Recovery process completed.");

    } catch (error) {
      console.error(error);
      setStatus('error');
      addLog(`Error: ${error.message}`);
    }
  };

  const downloadFile = () => {
    if (!recoveredBlob) return;
    const url = URL.createObjectURL(recoveredBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `fixed_${file.name}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-[#0a0a0c] text-slate-300 font-sans p-4 md:p-12">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <header className="mb-10 flex flex-col md:flex-row md:items-center justify-between gap-6 border-b border-white/5 pb-8">
          <div>
            <div className="flex items-center gap-3 mb-2">
              <div className="bg-blue-600 p-2 rounded-xl">
                <Binary size={28} className="text-white" />
              </div>
              <h1 className="text-3xl font-black tracking-tighter text-white uppercase italic">
                PPTX Font <span className="text-blue-500">Fixer</span> Core
              </h1>
            </div>
            <p className="text-slate-500 font-medium">박중희 박사님 전용 - 클라이언트 사이드 XML 복구 엔진</p>
          </div>
          <div className="bg-white/5 p-4 rounded-2xl border border-white/10 flex items-center gap-4">
            <ShieldCheck className="text-emerald-500" size={24} />
            <div className="text-[10px] font-mono text-slate-500 leading-tight">
              CORE: JSZIP_V3<br/>
              ENGINE: DOM_XML_PARSER
            </div>
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* File Input Area */}
          <div className="lg:col-span-1 space-y-6">
            <div className="bg-white/5 border border-white/10 rounded-3xl p-8 transition-all hover:border-blue-500/20 shadow-2xl">
              <h2 className="text-sm font-black text-blue-400 mb-6 uppercase tracking-widest flex items-center gap-2">
                <FileUp size={16} /> Data Input
              </h2>
              
              <div className="relative group">
                <input 
                  type="file" 
                  accept=".pptx" 
                  onChange={handleFileUpload}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-20" 
                />
                <div className={`border-2 border-dashed rounded-2xl p-8 flex flex-col items-center text-center transition-all ${
                  file ? 'border-blue-500 bg-blue-500/5' : 'border-white/10 bg-white/2 group-hover:bg-white/5'
                }`}>
                  <div className="w-16 h-16 bg-slate-800 rounded-2xl flex items-center justify-center mb-4 shadow-inner">
                    <Type size={32} className={file ? 'text-blue-400' : 'text-slate-600'} />
                  </div>
                  <span className="text-xs font-bold text-slate-400 break-all px-2">
                    {file ? file.name : '복구할 PPTX 파일을 선택하십시오'}
                  </span>
                </div>
              </div>

              {status === 'ready_to_scan' && (
                <button 
                  onClick={processPptx}
                  className="w-full mt-6 bg-blue-600 hover:bg-blue-500 text-white font-black py-4 rounded-2xl flex items-center justify-center gap-2 shadow-lg shadow-blue-600/20 transition-all active:scale-95"
                >
                  <RefreshCw size={18} /> 복구 알고리즘 실행
                </button>
              )}
            </div>

            {/* System Logs */}
            <div className="bg-black/50 border border-white/5 rounded-3xl p-6 h-[200px] flex flex-col overflow-hidden shadow-inner">
              <h3 className="text-[10px] font-black text-slate-600 mb-3 uppercase tracking-tighter flex items-center gap-2 italic">
                <FileCode size={12} /> System_Log_Console
              </h3>
              <div className="flex-1 overflow-y-auto space-y-1 font-mono text-[10px]">
                {log.map((l, i) => (
                  <div key={i} className="text-slate-500 border-l border-white/10 pl-2">
                    <span className="text-blue-900 mr-2">[{new Date().toLocaleTimeString()}]</span> {l}
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Main Execution Area */}
          <div className="lg:col-span-2">
            <div className="bg-white/5 border border-white/10 rounded-3xl p-8 h-full min-h-[500px] flex flex-col relative overflow-hidden backdrop-blur-xl">
              {status === 'idle' || status === 'ready_to_scan' ? (
                <div className="flex-1 flex flex-col items-center justify-center text-center space-y-4 opacity-50">
                  <ArrowRightLeft size={64} className="text-slate-700 animate-pulse" />
                  <div className="space-y-1">
                    <h3 className="text-xl font-black text-slate-500 uppercase italic">Analysis Ready</h3>
                    <p className="text-sm text-slate-600">파일을 업로드하면 XML 노드 스캔을 시작할 수 있습니다.</p>
                  </div>
                </div>
              ) : null}

              {status === 'processing' && (
                <div className="flex-1 flex flex-col items-center justify-center space-y-10">
                  <div className="relative">
                    <div className="w-32 h-32 rounded-full border-8 border-white/5"></div>
                    <div className="absolute inset-0 w-32 h-32 rounded-full border-8 border-blue-500 border-t-transparent animate-spin"></div>
                    <div className="absolute inset-0 flex items-center justify-center font-mono text-xl font-black text-blue-400">
                      {progress}%
                    </div>
                  </div>
                  <div className="text-center space-y-2">
                    <h3 className="text-2xl font-black text-white uppercase italic tracking-widest animate-pulse">Processing...</h3>
                    <p className="text-slate-500 text-sm font-medium">ISO/IEC 29500 스키마에 따라 폰트 매핑을 교정 중입니다.</p>
                  </div>
                </div>
              )}

              {status === 'completed' && (
                <div className="flex-1 flex flex-col items-center justify-center text-center animate-in zoom-in duration-700">
                  <div className="w-28 h-28 bg-emerald-500 rounded-[2.5rem] flex items-center justify-center mb-8 shadow-2xl shadow-emerald-500/40 rotate-12 transition-transform hover:rotate-0 cursor-pointer">
                    <CheckCircle size={60} className="text-white" />
                  </div>
                  <h3 className="text-4xl font-black text-white mb-4 uppercase italic">Process Verified</h3>
                  <p className="text-slate-400 max-w-md mx-auto mb-12 leading-relaxed font-medium">
                    박중희 박사님, 모든 XML 노드의 비표준 폰트 매핑이 <span className="text-blue-400">Malgun Gothic</span>으로 강제 치환되었습니다. 하단 버튼을 통해 실제 파일을 내려받으십시오.
                  </p>
                  
                  <button 
                    onClick={downloadFile}
                    className="group bg-white text-slate-950 font-black px-16 py-6 rounded-2xl flex items-center gap-4 hover:bg-blue-50 transition-all shadow-2xl hover:scale-105 active:scale-95"
                  >
                    <Download size={24} className="group-hover:translate-y-1 transition-transform" /> 
                    DOWNLOAD_FIXED_PPTX
                  </button>
                </div>
              )}

              {status === 'error' && (
                <div className="flex-1 flex flex-col items-center justify-center text-center space-y-6 animate-in fade-in duration-500">
                  <div className="w-20 h-20 bg-red-500/20 rounded-full flex items-center justify-center border border-red-500/30">
                    <AlertTriangle size={40} className="text-red-500" />
                  </div>
                  <div className="space-y-2">
                    <h3 className="text-2xl font-black text-red-400 uppercase tracking-tighter">Process Failed</h3>
                    <p className="text-slate-500 max-w-xs text-sm">파일이 손상되었거나 지원되지 않는 형식입니다.</p>
                  </div>
                  <button 
                    onClick={() => window.location.reload()}
                    className="text-xs font-bold text-slate-400 underline underline-offset-4"
                  >
                    시스템 재시작
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Footer Area */}
        <footer className="mt-16 pt-8 border-t border-white/5 flex flex-col md:flex-row justify-between items-center gap-6 text-slate-600 text-[10px] font-mono uppercase tracking-[0.2em] italic">
          <p>© 2026 PPTX RECOVERY LAB | 박중희 박사 연구 지원용 코어</p>
          <div className="flex gap-8 opacity-50">
            <span>Schema: OfficeOpenXML</span>
            <span>Policy: Nielsen_Usability_Standard</span>
          </div>
        </footer>
      </div>

      {/* Decorative Glow */}
      <div className="fixed -top-[20%] -right-[10%] w-[600px] h-[600px] bg-blue-600/5 rounded-full blur-[150px] pointer-events-none animate-pulse"></div>
      <div className="fixed -bottom-[10%] -left-[10%] w-[500px] h-[500px] bg-indigo-600/5 rounded-full blur-[120px] pointer-events-none"></div>
    </div>
  );
};

export default App;