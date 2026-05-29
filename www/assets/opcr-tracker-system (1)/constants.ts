import { CategoryType, OpcrData } from './types';

export const INITIAL_DATA: OpcrData = {
  id: '1',
  department: 'Field Operations Management Department',
  year: '2025',
  lastUpdated: new Date().toISOString(),
  signatories: {
    committing: "RYAN JAMES V. AYSON",
    approving: "PATRICK JAMES B. DIZON",
    pmt: "JOSE ALFREDO B. ESCOTO, JR.",
    head: "LEONOR C. CLEOFAS, CESO IV"
  },
  sections: [
    {
      id: 'sec_a',
      title: CategoryType.MFO,
      rows: [
        {
          id: '1.1',
          performanceIndicator: '1.1 Assistance in the Monitoring of commercial operations of Stage 1 & 2\n1.1.1 Availability of supply (capacity) based on average Minimum Committed Volume (MCV)',
          budget: '',
          successIndicator: 'Meet or exceed the required present MCV (192.93 MLD)',
          responsibleDivision: 'AME',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        },
        {
          id: '1.1.2',
          performanceIndicator: '1.1.2 Annual Bulk Water Charge (BWC for 2025 due to the 2024 CPI Adjustment)',
          budget: '',
          successIndicator: 'MWSS Board Approval within the third quarter of 2025',
          responsibleDivision: 'AME',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        },
        {
          id: 'mfo2_1.1.1',
          performanceIndicator: 'MF02: Operations and Maintenance... \n1.1.1 West Zone downstream projects (MWSI)',
          budget: '',
          successIndicator: 'Submission of Quarterly Monitoring Report 10 days from the end of the quarter',
          responsibleDivision: 'RJVA/East and West Divisions',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        }
      ]
    },
    {
      id: 'sec_b',
      title: CategoryType.STO,
      rows: [
        {
          id: 'sto_1.1',
          performanceIndicator: '1.1 MWSS Multi-Level Parking Project-Phase 2 (Design & Build)',
          budget: '',
          successIndicator: '100% Progress Accomplishment',
          responsibleDivision: 'RJVA/AME/MCMP',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        },
        {
          id: 'sto_2.1',
          performanceIndicator: '2.1 Submission of Monthly Client Satisfaction Survey Report',
          budget: '',
          successIndicator: 'Submission to the CART on the 15th day of the following month\n90% of CSM Survey SQDs rated as "Agree"',
          responsibleDivision: 'ACH/FOMD Personnel',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        }
      ]
    },
    {
      id: 'sec_c',
      title: CategoryType.GAS,
      rows: [
        {
          id: 'gas_1.1',
          performanceIndicator: '1.1 Procurement of Construction Takeoff and Estimating Software (Plans wift)',
          budget: '',
          successIndicator: 'TOR submitted to the BAC by February 2025.',
          responsibleDivision: 'FOMD Assigned Personnel',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        },
        {
          id: 'gas_2.1',
          performanceIndicator: '2.1 Compilation of the required documents (ROR, ROR Matrix, Manual of Operations)',
          budget: '',
          successIndicator: '100% Signed Controlled Copy of Documents prior to Audit',
          responsibleDivision: 'TIRDD',
          rating: { q: null, e: null, t: null, ave: null },
          accomplishment: {
            q1: { text: '', evidenceFiles: [] },
            q2: { text: '', evidenceFiles: [] },
            q3: { text: '', evidenceFiles: [] },
            q4: { text: '', evidenceFiles: [] },
          }
        }
      ]
    }
  ]
};