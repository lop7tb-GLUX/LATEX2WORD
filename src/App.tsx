/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  FileText, 
  Download, 
  Copy, 
  RefreshCw, 
  BookOpen, 
  HelpCircle, 
  Info,
  Type,
  FileDown,
  Github,
  Maximize2,
  CheckCircle2
} from 'lucide-react';
import TextareaAutosize from 'react-textarea-autosize';
import katex from 'katex';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Math, 
  MathRun,
  AlignmentType,
  HeadingLevel
} from 'docx';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';

// --- Types ---

interface ContentSegment {
  type: 'text' | 'math-inline' | 'math-display' | 'heading';
  content: string;
  level?: number;
}

// --- Utils ---

/**
 * Parses LaTeX string into segments of text, inline math, display math, and headings.
 */
const parseLatex = (input: string): ContentSegment[] => {
  const lines = input.split('\n');
  const segments: ContentSegment[] = [];

  lines.forEach(line => {
    if (!line.trim()) {
      segments.push({ type: 'text', content: '' });
      return;
    }

    // Basic Heading detection
    if (line.startsWith('\\section{')) {
      const title = line.match(/\\section\{([^}]+)\}/)?.[1] || '';
      segments.push({ type: 'heading', content: title, level: 1 });
      return;
    }
    if (line.startsWith('\\subsection{')) {
      const title = line.match(/\\subsection\{([^}]+)\}/)?.[1] || '';
      segments.push({ type: 'heading', content: title, level: 2 });
      return;
    }

    // Math detection within a line
    // This is a simplified regex-based parser
    let remaining = line;
    const mathRegex = /(\$\$[\s\S]*?\$\$|\$[\s\S]*?\$)/g;
    let match;
    let lastIndex = 0;

    const matches: { index: number; content: string; type: 'inline' | 'display' }[] = [];
    while ((match = mathRegex.exec(line)) !== null) {
      const isDisplay = match[0].startsWith('$$');
      matches.push({
        index: match.index,
        content: isDisplay ? match[0].slice(2, -2) : match[0].slice(1, -1),
        type: isDisplay ? 'display' : 'inline'
      });
    }

    if (matches.length === 0) {
      segments.push({ type: 'text', content: line });
    } else {
      let cursor = 0;
      matches.forEach(m => {
        if (m.index > cursor) {
          segments.push({ type: 'text', content: line.substring(cursor, m.index) });
        }
        segments.push({ 
          type: m.type === 'display' ? 'math-display' : 'math-inline', 
          content: m.content 
        });
        cursor = m.index + (m.type === 'display' ? m.content.length + 4 : m.content.length + 2);
      });
      if (cursor < line.length) {
        segments.push({ type: 'text', content: line.substring(cursor) });
      }
    }
  });

  return segments;
};

const TEMPLATES = [
  {
    name: "Công thức bậc 2",
    code: "\\section{Giải phương trình bậc hai}\nPhương trình có dạng $ax^2 + bx + c = 0$.\nCông thức nghiệm:\n$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$\nTrong đó $\\Delta = b^2 - 4ac$."
  },
  {
    name: "Tích phân & Đạo hàm",
    code: "\\section{Giải tích nâng cao}\nĐạo hàm của hàm số:\n$$\\frac{d}{dx} e^x = e^x$$\nTích phân xác định:\n$$\\int_{a}^{b} f(x) dx = F(b) - F(a)$$"
  },
  {
    name: "Ma trận & Hệ phương trình",
    code: "\\section{Đại số tuyến tính}\nMa trận đơn vị:\n$$I_3 = \\begin{pmatrix} 1 & 0 & 0 \\\\ 0 & 1 & 0 \\\\ 0 & 0 & 1 \\end{pmatrix}$$\nĐịnh thức của ma trận A ký hiệu là $|A|$."
  }
];

export default function App() {
  const [latex, setLatex] = useState<string>(TEMPLATES[0].code);
  const [isCopying, setIsCopying] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);

  // Parse segments for preview and export
  const segments = useMemo(() => parseLatex(latex), [latex]);

  const handleExportWord = async () => {
    setIsDownloading(true);
    try {
      const doc = new Document({
        sections: [{
          properties: {},
          children: segments.map(seg => {
            if (seg.type === 'heading') {
              return new Paragraph({
                text: seg.content,
                heading: seg.level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
                spacing: { before: 240, after: 120 }
              });
            }
            if (seg.type === 'math-display') {
              return new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new Math({
                    children: [new MathRun(seg.content)]
                  })
                ],
                spacing: { before: 120, after: 120 }
              });
            }
            if (seg.type === 'math-inline') {
              // docx math-inline needs careful placement within a text paragraph
              // For simplicity in this demo, we'll wrap it in a special marker
              return new Paragraph({
                children: [
                  new TextRun({ text: " " }),
                  new Math({ children: [new MathRun(seg.content)] }),
                  new TextRun({ text: " " })
                ]
              });
            }
            return new Paragraph({
              children: [new TextRun(seg.content)],
              spacing: { after: 100 }
            });
          })
        }]
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, "chuyen-doi-latex.docx");
    } catch (error) {
      console.error("Export failed", error);
    } finally {
      setIsDownloading(false);
    }
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(latex);
    setIsCopying(true);
    setTimeout(() => setIsCopying(false), 2000);
  };

  return (
    <div className="flex flex-col h-screen max-h-screen overflow-hidden text-slate-800">
      {/* --- Header --- */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shadow-sm z-10 shrink-0">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-2 rounded-lg text-white shadow-indigo-100 shadow-lg">
            <FileText size={24} strokeWidth={2.5} />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-slate-900 leading-tight">
              Latex2Word
            </h1>
            <p className="text-[10px] uppercase tracking-widest font-bold text-indigo-500 opacity-80">
              Công cụ chuyển đổi thông minh
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button 
            onClick={copyToClipboard}
            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-full hover:bg-slate-100 transition-all active:scale-95"
          >
            {isCopying ? <CheckCircle2 size={16} className="text-green-500" /> : <Copy size={16} />}
            {isCopying ? "Đã sao chép" : "Sao chép mã"}
          </button>
          
          <button 
            onClick={handleExportWord}
            disabled={isDownloading}
            className="flex items-center gap-2 px-6 py-2 text-sm font-bold text-white bg-indigo-600 rounded-full hover:bg-indigo-700 transition-all shadow-md shadow-indigo-100 disabled:opacity-50 active:scale-95 group"
          >
            {isDownloading ? <RefreshCw size={18} className="animate-spin" /> : <FileDown size={18} className="group-hover:translate-y-0.5 transition-transform" />}
            Tải file Word
          </button>
        </div>
      </header>

      {/* --- Main Content --- */}
      <main className="flex-1 flex overflow-hidden">
        {/* Sidebar Templates */}
        <div className="w-64 bg-slate-50 border-r border-slate-200 flex flex-col shrink-0">
          <div className="p-5 flex-1 overflow-y-auto">
            <div className="flex items-center gap-2 mb-6 text-slate-400">
              <BookOpen size={16} />
              <span className="text-xs font-bold uppercase tracking-wider">Thư viện mẫu</span>
            </div>
            
            <div className="space-y-3">
              {TEMPLATES.map((tmpl, idx) => (
                <button
                  key={idx}
                  onClick={() => setLatex(tmpl.code)}
                  className="w-full text-left p-3 rounded-xl bg-white border border-slate-200 hover:border-indigo-300 hover:shadow-sm transition-all group"
                >
                  <p className="text-sm font-bold text-slate-700 mb-1 group-hover:text-indigo-600 truncate">{tmpl.name}</p>
                  <p className="text-[10px] text-slate-400 uppercase font-mono truncate">Code: {tmpl.code.substring(0, 15)}...</p>
                </button>
              ))}
            </div>

            <div className="mt-8 pt-8 border-t border-slate-200">
              <div className="flex items-center gap-2 mb-4 text-slate-400">
                <Info size={16} />
                <span className="text-xs font-bold uppercase tracking-wider">Hướng dẫn</span>
              </div>
              <ul className="text-xs text-slate-500 space-y-3 leading-relaxed">
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  Sử dụng $...$ cho toán học nội dòng (inline).
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  Sử dụng $$...$$ cho toán học trung tâm (display).
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  Dùng {"\\section{...}"} để tạo đề mục lớn.
                </li>
              </ul>
            </div>
          </div>

          <div className="p-4 bg-slate-100/50 text-[10px] text-slate-400 text-center font-mono uppercase tracking-widest leading-loose">
            Phát triển bởi Loop7TB <br/> dành cho cộng đồng
          </div>
        </div>

        {/* Editor Area */}
        <div className="flex-1 flex flex-col bg-white overflow-hidden">
          <div className="px-4 py-2 bg-slate-50/50 border-b border-slate-100 flex items-center justify-between shrink-0">
             <div className="flex items-center gap-2">
               <Type size={14} className="text-slate-400" />
               <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Trình soạn thảo LaTeX</span>
             </div>
             <div className="text-[10px] font-mono text-slate-300">UTF-8</div>
          </div>
          <div className="flex-1 overflow-y-auto p-8 relative">
            {/* Simple Line Numbers */}
            <div className="absolute left-0 top-8 w-8 pr-2 text-right text-slate-200 font-mono text-sm leading-relaxed pointer-events-none select-none">
              {latex.split('\n').map((_, i) => <div key={i}>{i + 1}</div>)}
            </div>
            <textarea
              value={latex}
              onChange={(e) => setLatex(e.target.value)}
              className="w-full min-h-full resize-none focus:outline-none font-mono text-sm leading-relaxed text-slate-600 ml-4 pb-20"
              spellCheck={false}
              placeholder="Nhập mã LaTeX của bạn tại đây..."
            />
          </div>
        </div>

        {/* Preview Area */}
        <div className="flex-1 flex flex-col bg-slate-100 overflow-hidden border-l border-slate-200">
          <div className="px-4 py-2 bg-white/50 border-b border-slate-100 flex items-center gap-2 shrink-0">
             <Maximize2 size={14} className="text-slate-400" />
             <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Xem trước văn bản</span>
          </div>
          <div className="flex-1 overflow-y-auto p-10 math-preview">
             <div className="max-w-2xl mx-auto bg-white min-h-[120%] p-16 shadow-xl shadow-slate-200/50 rounded-sm border border-slate-200">
               <AnimatePresence mode="wait">
                 <motion.div 
                   key={latex}
                   initial={{ opacity: 0 }}
                   animate={{ opacity: 1 }}
                   transition={{ duration: 0.2 }}
                 >
                   {segments.map((seg, idx) => {
                     if (seg.type === 'heading') {
                       return (
                         <h2 key={idx} className={`font-bold text-indigo-900 tracking-tight mb-4 ${seg.level === 1 ? 'text-2xl mt-8' : 'text-xl mt-6 underline underline-offset-8 decoration-slate-200'}`}>
                           {seg.content}
                         </h2>
                       );
                     }
                     if (seg.type === 'math-display') {
                       try {
                         const html = katex.renderToString(seg.content, { displayMode: true, throwOnError: false });
                         return <div key={idx} dangerouslySetInnerHTML={{ __html: html }} />;
                       } catch (e) {
                         return <div key={idx} className="text-red-500 text-xs italic p-2 border border-red-100 rounded">Lỗi LaTeX: {seg.content}</div>;
                       }
                     }
                     if (seg.type === 'math-inline') {
                       try {
                         const html = katex.renderToString(seg.content, { displayMode: false, throwOnError: false });
                         return <span key={idx} dangerouslySetInnerHTML={{ __html: html }} />;
                       } catch (e) {
                         return <span key={idx} className="text-red-500 text-xs italic">[Lỗi: {seg.content}]</span>;
                       }
                     }
                     // Content rendering with line breaks
                     return (
                       <p key={idx} className="inline text-slate-700 leading-relaxed">
                         {seg.content === '' ? <br /> : seg.content}
                       </p>
                     );
                   })}
                 </motion.div>
               </AnimatePresence>
             </div>
          </div>
        </div>
      </main>

      {/* --- Footer Status --- */}
      <footer className="bg-white border-t border-slate-200 px-6 py-2 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse"></span>
            Sẵn sàng
          </div>
          <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            Ký tự: <span className="text-slate-600">{latex.length}</span>
          </div>
          <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-wider text-slate-400">
            Dòng: <span className="text-slate-600">{latex.split('\n').length}</span>
          </div>
        </div>
        
        <div className="flex items-center gap-4 text-slate-300">
           <a href="#" className="hover:text-indigo-500 transition-colors"><Github size={14} /></a>
           <a href="#" className="hover:text-indigo-500 transition-colors"><HelpCircle size={14} /></a>
        </div>
      </footer>
    </div>
  );
}
