import { useState, useRef } from 'react';
import { Copy, Eraser, Check, Upload, Download } from 'lucide-react';
import * as mammoth from 'mammoth';
import JSZip from 'jszip';

interface FileData {
  originalBlob: Blob;
  text: string;
  fileName: string;
}

function App() {
  const [inputText, setInputText] = useState('');
  const [outputText, setOutputText] = useState('');
  const [copied, setCopied] = useState(false);
  const [fileName, setFileName] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileData, setFileData] = useState<FileData | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const cleanText = () => {
    const cleaned = inputText.replace(/\s+/g, ' ').trim();
    setOutputText(cleaned);
  };

  const clearAll = () => {
    setInputText('');
    setOutputText('');
    setFileName('');
    setFileData(null);
    setCopied(false);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(outputText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.name.endsWith('.docx')) {
      alert('Please upload a DOCX file');
      return;
    }

    setIsProcessing(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      const text = result.value;
      setInputText(text);
      setFileName(file.name);
      setFileData({
        originalBlob: file,
        text: text,
        fileName: file.name,
      });
    } catch (error) {
      alert('Error reading document. Please ensure it is a valid DOCX file.');
      console.error(error);
    } finally {
      setIsProcessing(false);
    }
  };

  const cleanDocxPreservingFormat = async (
    originalBlob: Blob,
    cleanedText: string
  ): Promise<Blob> => {
    const arrayBuffer = await originalBlob.arrayBuffer();
    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);

    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) throw new Error('Invalid DOCX file');

    const originalText = inputText;
    let cleanedXml = documentXml;

    const textNodes = documentXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    let cleanedTextIndex = 0;
    let originalTextIndex = 0;

    for (const node of textNodes) {
      const match = node.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
      if (!match) continue;

      const nodeText = match[1];
      if (!nodeText.trim()) continue;

      const originalTextRemaining = originalText.slice(originalTextIndex);
      if (originalTextRemaining.startsWith(nodeText)) {
        const cleanedTextRemaining = cleanedText.slice(cleanedTextIndex);
        let replacementText = '';

        for (let i = 0; i < nodeText.length; i++) {
          if (cleanedTextIndex < cleanedText.length) {
            replacementText += cleanedText[cleanedTextIndex];
            if (nodeText[i] !== ' ' || cleanedText[cleanedTextIndex] !== ' ') {
              cleanedTextIndex++;
            }
          }
        }

        const oldNode = node;
        const newNode = oldNode.replace(
          nodeText,
          replacementText.slice(0, nodeText.length)
        );
        cleanedXml = cleanedXml.replace(oldNode, newNode);
        originalTextIndex += nodeText.length;
      }
    }

    zip.file('word/document.xml', cleanedXml);
    return await zip.generateAsync({ type: 'blob' });
  };

  const downloadDocx = async () => {
    if (!outputText) return;

    try {
      let blob: Blob;

      if (fileData) {
        blob = await cleanDocxPreservingFormat(fileData.originalBlob, outputText);
      } else {
        const { Document, Packer, Paragraph } = await import('docx');
        const doc = new Document({
          sections: [
            {
              properties: {},
              children: [
                new Paragraph({
                  text: outputText,
                }),
              ],
            },
          ],
        });
        blob = await Packer.toBlob(doc);
      }

      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName ? `cleaned_${fileName}` : 'cleaned_document.docx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      alert('Error creating document');
      console.error(error);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      <div className="max-w-5xl mx-auto px-6 py-12">
        <div className="text-center mb-12">
          <h1 className="text-4xl font-bold text-slate-800 mb-3">Text Space Cleaner</h1>
          <p className="text-slate-600 text-lg">Remove extra spaces between words instantly</p>
        </div>

        <div className="mb-6 flex justify-center">
          <label className="flex items-center gap-2 px-6 py-3 bg-blue-50 border-2 border-dashed border-blue-300 rounded-lg cursor-pointer hover:bg-blue-100 transition-colors">
            <Upload className="w-5 h-5 text-blue-600" />
            <span className="font-medium text-blue-600">Upload DOCX</span>
            <input
              ref={fileInputRef}
              type="file"
              accept=".docx"
              onChange={handleFileUpload}
              disabled={isProcessing}
              className="hidden"
            />
          </label>
          {fileName && <span className="ml-4 text-slate-600">{fileName}</span>}
          {isProcessing && <span className="ml-4 text-slate-500 animate-pulse">Processing...</span>}
        </div>

        <div className="grid md:grid-cols-2 gap-6">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-slate-700">Input Text</h2>
              <span className="text-sm text-slate-500">{inputText.length} characters</span>
            </div>
            <textarea
              value={inputText}
              onChange={(e) => setInputText(e.target.value)}
              placeholder="Paste your text here with extra    spaces..."
              className="w-full h-64 p-4 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none text-slate-700"
            />
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-slate-700">Cleaned Text</h2>
              <span className="text-sm text-slate-500">{outputText.length} characters</span>
            </div>
            <textarea
              value={outputText}
              readOnly
              placeholder="Cleaned text will appear here..."
              className="w-full h-64 p-4 border border-slate-200 rounded-lg bg-slate-50 resize-none text-slate-700"
            />
          </div>
        </div>

        <div className="flex justify-center gap-4 mt-8 flex-wrap">
          <button
            onClick={cleanText}
            disabled={!inputText.trim()}
            className="px-8 py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 active:scale-95 transition-all disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:bg-blue-600 shadow-sm"
          >
            Clean Text
          </button>

          <button
            onClick={copyToClipboard}
            disabled={!outputText}
            className="px-6 py-3 bg-slate-600 text-white font-medium rounded-lg hover:bg-slate-700 active:scale-95 transition-all disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:bg-slate-600 shadow-sm flex items-center gap-2"
          >
            {copied ? (
              <>
                <Check className="w-4 h-4" />
                Copied!
              </>
            ) : (
              <>
                <Copy className="w-4 h-4" />
                Copy
              </>
            )}
          </button>

          <button
            onClick={downloadDocx}
            disabled={!outputText}
            className="px-6 py-3 bg-green-600 text-white font-medium rounded-lg hover:bg-green-700 active:scale-95 transition-all disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:bg-green-600 shadow-sm flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            Download DOCX
          </button>

          <button
            onClick={clearAll}
            className="px-6 py-3 bg-slate-200 text-slate-700 font-medium rounded-lg hover:bg-slate-300 active:scale-95 transition-all shadow-sm flex items-center gap-2"
          >
            <Eraser className="w-4 h-4" />
            Clear
          </button>
        </div>
      </div>
    </div>
  );
}

export default App;
