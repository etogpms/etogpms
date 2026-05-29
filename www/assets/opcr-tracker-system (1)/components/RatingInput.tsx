import React, { useState } from 'react';
import { Rating } from '../types';

interface RatingInputProps {
  currentRating: Rating;
  onSave: (rating: Rating) => void;
  onClose: () => void;
}

export const RatingInput: React.FC<RatingInputProps> = ({ currentRating, onSave, onClose }) => {
  const [values, setValues] = useState(currentRating);

  const handleChange = (key: keyof Rating, val: string) => {
    const num = val === '' ? null : parseFloat(val);
    setValues(prev => ({
      ...prev,
      [key]: num
    }));
  };

  const calculateAve = () => {
    const { q, e, t } = values;
    const nums = [q, e, t].filter(n => n !== null) as number[];
    if (nums.length === 0) return null;
    return parseFloat((nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(2));
  };

  const handleSave = () => {
    const ave = calculateAve();
    onSave({ ...values, ave });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-transparent" onClick={onClose}>
      <div className="bg-white p-4 rounded shadow-xl border border-slate-200 w-64" onClick={e => e.stopPropagation()}>
        <h3 className="font-bold text-xs mb-3 text-slate-700 uppercase tracking-wide">Edit Ratings</h3>
        <div className="grid grid-cols-3 gap-2 mb-4">
          <div>
            <label className="block text-xs font-semibold text-center mb-1">Q</label>
            <input 
              type="number" step="0.1" max="5" min="1"
              value={values.q ?? ''} 
              onChange={e => handleChange('q', e.target.value)}
              className="w-full text-center border p-1 rounded text-sm"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-center mb-1">E</label>
            <input 
              type="number" step="0.1" max="5" min="1"
              value={values.e ?? ''} 
              onChange={e => handleChange('e', e.target.value)}
              className="w-full text-center border p-1 rounded text-sm"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-center mb-1">T</label>
            <input 
              type="number" step="0.1" max="5" min="1"
              value={values.t ?? ''} 
              onChange={e => handleChange('t', e.target.value)}
              className="w-full text-center border p-1 rounded text-sm"
            />
          </div>
        </div>
        <div className="flex justify-between">
          <button onClick={onClose} className="text-xs text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={handleSave} className="text-xs bg-blue-600 text-white px-3 py-1 rounded hover:bg-blue-700">Apply</button>
        </div>
      </div>
    </div>
  );
};