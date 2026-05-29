import React, { useState, useRef } from 'react';
import { IndicatorRow, AttachedFile } from '../types';
import { uploadEvidence } from '../services/storageService';

interface UpdateModalProps {
  isOpen: boolean;
  onClose: () => void;
  row: IndicatorRow;
  sectionId: string;
  quarter: 'q1' | 'q2' | 'q3' | 'q4';
  onSave: (sectionId: string, rowId: string, quarter: 'q1' | 'q2' | 'q3' | 'q4', text: string, files: AttachedFile[]) => void;
}

export const UpdateModal: React.FC<UpdateModalProps> = ({ isOpen, onClose, row, sectionId, quarter, onSave }) => {
  const currentData = row.accomplishment[quarter];
  const [text, setText] = useState(currentData.text);
  const [files, setFiles] = useState<AttachedFile[]>(currentData.evidenceFiles);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setIsUploading(true);
      try {
        const file = e.target.files[0];
        const url = await uploadEvidence(file);
        const newFile: AttachedFile = {
          id: Date.now().toString(),
          name: file.name,
          url: url,
          dateUploaded: new Date().toISOString()
        };
        setFiles(prev => [...prev, newFile]);
      } catch (err) {
        alert("Failed to upload");
      } finally {
        setIsUploading(false);
      }
    }
  };

  const handleSave = () => {
    onSave(sectionId, row.id, quarter, text, files);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-2xl flex flex-col max-h-[90vh]">
        <div className="p-6 border-b">
          <h2 className="text-xl font-bold text-slate-800">Update Accomplishment - {quarter.toUpperCase()}</h2>
          <p className="text-sm text-slate-500 mt-1 truncate">{row.performanceIndicator}</p>
        </div>
        
        <div className="p-6 overflow-y-auto flex-1">
          <label className="block text-sm font-medium text-slate-700 mb-2">
            Accomplishment Narrative / Status
          </label>
          <textarea
            className="w-full border border-slate-300 rounded-md p-3 min-h-[150px] focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder="Enter the accomplishments for this period..."
          ></textarea>

          <div className="mt-6">
            <div className="flex justify-between items-center mb-2">
              <label className="block text-sm font-medium text-slate-700">
                Evidence / Attachments
              </label>
              <button 
                onClick={() => fileInputRef.current?.click()}
                disabled={isUploading}
                className="text-sm bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1 rounded transition-colors flex items-center gap-2"
              >
                {isUploading ? 'Uploading...' : '+ Upload File'}
              </button>
              <input 
                type="file" 
                hidden 
                ref={fileInputRef} 
                onChange={handleFileUpload}
              />
            </div>

            {files.length === 0 ? (
              <div className="text-sm text-slate-400 italic border border-dashed border-slate-200 rounded p-4 text-center">
                No files attached yet.
              </div>
            ) : (
              <ul className="space-y-2">
                {files.map(file => (
                  <li key={file.id} className="flex items-center justify-between bg-blue-50 p-2 rounded border border-blue-100">
                    <div className="flex items-center gap-2 overflow-hidden">
                      <svg className="w-4 h-4 text-blue-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13"></path></svg>
                      <a href={file.url} target="_blank" rel="noreferrer" className="text-sm text-blue-700 hover:underline truncate">
                        {file.name}
                      </a>
                    </div>
                    <button 
                      onClick={() => setFiles(files.filter(f => f.id !== file.id))}
                      className="text-red-500 hover:text-red-700"
                    >
                      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"></path></svg>
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>

        <div className="p-4 border-t bg-gray-50 flex justify-end gap-3 rounded-b-lg">
          <button 
            onClick={onClose}
            className="px-4 py-2 text-slate-700 hover:bg-slate-200 rounded transition-colors"
          >
            Cancel
          </button>
          <button 
            onClick={handleSave}
            className="px-4 py-2 bg-blue-600 text-white hover:bg-blue-700 rounded transition-colors font-medium shadow-sm"
          >
            Save Changes
          </button>
        </div>
      </div>
    </div>
  );
};