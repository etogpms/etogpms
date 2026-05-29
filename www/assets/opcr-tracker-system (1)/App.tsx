import React, { useEffect, useState } from 'react';
import { OpcrData, AttachedFile, Rating, IndicatorRow } from './types';
import { loadOpcrData, saveOpcrData } from './services/storageService';
import { UpdateModal } from './components/UpdateModal';
import { RatingInput } from './components/RatingInput';

const App: React.FC = () => {
  const [data, setData] = useState<OpcrData | null>(null);
  const [loading, setLoading] = useState(true);
  
  // Modal States
  const [editingRow, setEditingRow] = useState<{sectionId: string, row: IndicatorRow, quarter: 'q1'|'q2'|'q3'|'q4'} | null>(null);
  const [ratingRow, setRatingRow] = useState<{sectionId: string, row: IndicatorRow} | null>(null);

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    const loadedData = await loadOpcrData();
    setData(loadedData);
    setLoading(false);
  };

  const handleUpdateAccomplishment = async (sectionId: string, rowId: string, quarter: 'q1'|'q2'|'q3'|'q4', text: string, files: AttachedFile[]) => {
    if (!data) return;
    
    const newData = { ...data };
    const section = newData.sections.find(s => s.id === sectionId);
    if (section) {
      const row = section.rows.find(r => r.id === rowId);
      if (row) {
        row.accomplishment[quarter] = { text, evidenceFiles: files };
        setData(newData);
        await saveOpcrData(newData);
      }
    }
  };

  const handleUpdateRating = async (rating: Rating) => {
    if (!data || !ratingRow) return;
    
    const newData = { ...data };
    const section = newData.sections.find(s => s.id === ratingRow.sectionId);
    if (section) {
      const row = section.rows.find(r => r.id === ratingRow.row.id);
      if (row) {
        row.rating = rating;
        setData(newData);
        await saveOpcrData(newData);
      }
    }
    setRatingRow(null); // Close popover
  };

  const handleDepartmentChange = async (val: string) => {
      if (!data) return;
      const newData = {...data, department: val};
      setData(newData);
      await saveOpcrData(newData);
  };

  if (loading || !data) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-8 print:p-0">
      {/* Action Bar - Hidden in Print */}
      <div className="max-w-[1400px] mx-auto mb-6 flex justify-between items-center no-print">
        <div>
          <h1 className="text-2xl font-bold text-slate-800">OPCR Digital Tracker</h1>
          <p className="text-slate-500 text-sm">Field Operations Management Department</p>
        </div>
        <div className="flex gap-3">
          <button 
            onClick={() => window.print()} 
            className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-white rounded hover:bg-slate-700 transition-colors shadow-sm"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
            Print Form
          </button>
        </div>
      </div>

      {/* Main Document Paper */}
      <div className="max-w-[1400px] mx-auto bg-white shadow-lg print:shadow-none p-6 print:p-0 print-fit">
        
        {/* Header Section */}
        <div className="border border-black mb-0 grid grid-cols-12 text-center">
          <div className="col-span-3 border-r border-black p-2 flex flex-col items-center justify-center">
            <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Metropolitan_Waterworks_and_Sewerage_System_%28MWSS%29.svg/1200px-Metropolitan_Waterworks_and_Sewerage_System_%28MWSS%29.svg.png" alt="MWSS Logo" className="h-16 w-auto mb-2" />
            <p className="text-[9px] font-bold leading-tight">Republic of the Philippines</p>
            <p className="text-[10px] font-bold leading-tight">PANGASIWAAN NG TUBIG AT ALKANTARAJYA SA METRO MANILA</p>
            <p className="text-[10px] font-bold">Metropolitan Waterworks and Sewerage System</p>
            <p className="text-[9px]">Katipunan Road, Balara, Quezon City</p>
          </div>
          <div className="col-span-6 border-r border-black p-2 flex items-center justify-center flex-col">
            <h2 className="font-bold text-lg uppercase">Office Performance Commitment and Review (OPCR)</h2>
            <div className="w-full">
                <input 
                    type="text" 
                    value={data.department} 
                    onChange={(e) => handleDepartmentChange(e.target.value)}
                    className="text-sm font-bold text-center w-full border-b border-gray-300 focus:border-blue-500 outline-none print:border-none print:p-0"
                    placeholder="Enter Department Name"
                />
            </div>
          </div>
          <div className="col-span-3 p-2 flex items-center justify-center flex-col">
            <p className="font-bold text-sm">RATING PERIOD:</p>
            <p className="text-sm">JANUARY - DECEMBER {data.year}</p>
          </div>
        </div>

        {/* Part I Signatories */}
        <div className="border-x border-b border-black bg-blue-100 p-1">
          <p className="text-xs font-bold">I. Signature approvals during the Performance Planning and Commitment Phase</p>
        </div>
        <div className="border-x border-b border-black grid grid-cols-3 text-center text-xs">
          <div className="border-r border-black p-4 relative">
            <p className="mb-8 italic text-[10px] text-left px-2">
              I, <b>{data.signatories.committing}</b>, Manager, {data.department}, commit to deliver and agree to be rated on the attainment of the following targets...
            </p>
            <div className="mt-8">
              <p className="font-bold underline uppercase">{data.signatories.committing}</p>
              <p>Manager, {data.department}</p>
            </div>
          </div>
          <div className="border-r border-black p-4">
            <p className="mb-8 italic text-[10px]">I hereby certify that the commitments stated herein was discussed with the ratee.</p>
            <div className="mt-8">
              <p className="font-bold underline uppercase">{data.signatories.approving}</p>
              <p>Acting Deputy Administrator, Engineering and Technical Operations Group</p>
            </div>
          </div>
          <div className="grid grid-cols-2">
            <div className="border-r border-black p-4 flex flex-col justify-end">
                <p className="text-[10px] mb-8">Discussed and agreed with the PMT.</p>
                <p className="font-bold underline uppercase">{data.signatories.pmt}</p>
                <p>PMT Chairperson</p>
            </div>
            <div className="p-4 flex flex-col justify-end">
                <p className="text-[10px] mb-8">Approved by:</p>
                <p className="font-bold underline uppercase">{data.signatories.head}</p>
                <p>Administrator</p>
            </div>
          </div>
        </div>

        {/* Table Header */}
        <table className="w-full border-collapse border border-black text-xs">
          <thead>
            <tr className="bg-[#00838F] text-white text-center font-bold">
              <th className="border border-black p-2 w-[25%] align-middle" rowSpan={2}>MFOs AND PERFORMANCE INDICATORS</th>
              <th className="border border-black p-2 w-[10%] align-middle" rowSpan={2}>DEPARTMENT<br/>FY {data.year} BUDGET</th>
              <th className="border border-black p-2 w-[20%] align-middle" rowSpan={2}>SUCCESS INDICATOR CY {data.year} TARGETS</th>
              <th className="border border-black p-2 w-[10%] align-middle" rowSpan={2}>RESPONSIBLE DIVISION/<br/>INDIVIDUALS</th>
              <th className="border border-black p-1 h-6" colSpan={4}>RATING</th>
              <th className="border border-black p-2 w-[25%] align-middle" rowSpan={2}>ACCOMPLISHMENT/UPDATE<br/><span className="text-[9px] font-normal">Annual / Semestral / Quarterly</span></th>
            </tr>
            <tr className="bg-[#00838F] text-white text-center font-bold h-8">
              <th className="border border-black w-8">Q<sup>1</sup></th>
              <th className="border border-black w-8">E<sup>2</sup></th>
              <th className="border border-black w-8">T<sup>3</sup></th>
              <th className="border border-black w-8">Ave</th>
            </tr>
          </thead>
          <tbody>
            {data.sections.map((section) => (
              <React.Fragment key={section.id}>
                {/* Section Title */}
                <tr className="bg-gray-200">
                  <td colSpan={9} className="border border-black p-1 font-bold pl-2">
                    {section.title}
                  </td>
                </tr>

                {/* Rows */}
                {section.rows.map((row) => (
                  <tr key={row.id} className="align-top hover:bg-gray-50">
                    <td className="border border-black p-2 whitespace-pre-wrap">{row.performanceIndicator}</td>
                    <td className="border border-black p-2 text-center">{row.budget}</td>
                    <td className="border border-black p-2 whitespace-pre-wrap">{row.successIndicator}</td>
                    <td className="border border-black p-2 text-center whitespace-pre-wrap">{row.responsibleDivision}</td>
                    
                    {/* Ratings */}
                    <td className="border border-black p-0 relative group">
                      <div className="absolute inset-0 flex items-center justify-center cursor-pointer hover:bg-yellow-50" onClick={() => setRatingRow({sectionId: section.id, row})}>
                        {row.rating.q}
                      </div>
                    </td>
                    <td className="border border-black p-0 relative group">
                       <div className="absolute inset-0 flex items-center justify-center cursor-pointer hover:bg-yellow-50" onClick={() => setRatingRow({sectionId: section.id, row})}>
                        {row.rating.e}
                      </div>
                    </td>
                    <td className="border border-black p-0 relative group">
                       <div className="absolute inset-0 flex items-center justify-center cursor-pointer hover:bg-yellow-50" onClick={() => setRatingRow({sectionId: section.id, row})}>
                        {row.rating.t}
                      </div>
                    </td>
                    <td className="border border-black p-0 relative bg-gray-50">
                       <div className="absolute inset-0 flex items-center justify-center font-bold">
                        {row.rating.ave}
                      </div>
                    </td>

                    {/* Accomplishment - Divided into 4 Mini Rows visually or just a list */}
                    <td className="border border-black p-0">
                      <div className="grid grid-rows-4 divide-y divide-black h-full">
                        {['q1', 'q2', 'q3', 'q4'].map((q) => {
                          const quarterKey = q as 'q1'|'q2'|'q3'|'q4';
                          const acc = row.accomplishment[quarterKey];
                          const hasContent = acc.text || acc.evidenceFiles.length > 0;
                          
                          return (
                            <div 
                              key={q} 
                              className="p-1 min-h-[40px] text-[10px] hover:bg-blue-50 cursor-pointer transition-colors relative group"
                              onClick={() => setEditingRow({sectionId: section.id, row, quarter: quarterKey})}
                            >
                              <span className="font-bold text-gray-400 absolute top-0.5 right-1 select-none pointer-events-none">{q.toUpperCase()}</span>
                              {hasContent ? (
                                <div>
                                  <p className="line-clamp-2">{acc.text}</p>
                                  {acc.evidenceFiles.length > 0 && (
                                    <div className="flex gap-1 mt-1">
                                       <span className="bg-blue-100 text-blue-800 px-1 rounded text-[9px] flex items-center">
                                         📎 {acc.evidenceFiles.length}
                                       </span>
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <span className="text-gray-300 italic p-1 block group-hover:text-blue-400">Add update...</span>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
          </tbody>
        </table>

        {/* Footer / Legend */}
        <div className="mt-4 text-[10px]">
            <p>Legend: 1 - Quality, 2 - Efficiency, 3 - Timeliness</p>
        </div>
        
        {/* Part II Signatories */}
         <div className="border-x border-t border-black bg-blue-100 p-1 mt-4 border-b">
          <p className="text-xs font-bold">II. Signature approvals after performance monitoring, assessment, feedback by To be signed after the Performance Review and Feedback by the PMT</p>
        </div>
        <div className="border-x border-b border-black grid grid-cols-4 text-center text-xs">
            <div className="border-r border-black p-8">
                <p className="text-left mb-8">Prepared by:</p>
                <div className="mt-8">
                     <p className="font-bold underline uppercase">{data.signatories.committing}</p>
                     <p>Manager, {data.department}</p>
                </div>
            </div>
             <div className="border-r border-black p-8">
                <p className="text-left mb-8">Reviewed by:</p>
                <div className="mt-8">
                     <p className="font-bold underline uppercase">{data.signatories.approving}</p>
                     <p>Acting Deputy Administrator</p>
                </div>
            </div>
             <div className="border-r border-black p-8">
                <p className="text-left mb-8">Recommended by:</p>
                <div className="mt-8">
                     <p className="font-bold underline uppercase">{data.signatories.pmt}</p>
                     <p>PMT Chairperson</p>
                </div>
            </div>
             <div className="p-8">
                <p className="text-left mb-8">The final rating was discussed with the ratee. Approved and Rated by:</p>
                <div className="mt-8">
                     <p className="font-bold underline uppercase">{data.signatories.head}</p>
                     <p>Administrator</p>
                </div>
            </div>
        </div>

      </div>

      {/* Modals */}
      {editingRow && (
        <UpdateModal
          isOpen={!!editingRow}
          onClose={() => setEditingRow(null)}
          row={editingRow.row}
          sectionId={editingRow.sectionId}
          quarter={editingRow.quarter}
          onSave={handleUpdateAccomplishment}
        />
      )}

      {ratingRow && (
        <RatingInput
            currentRating={ratingRow.row.rating}
            onClose={() => setRatingRow(null)}
            onSave={handleUpdateRating}
        />
      )}

    </div>
  );
};

export default App;