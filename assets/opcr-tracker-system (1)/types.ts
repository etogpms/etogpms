export enum CategoryType {
  MFO = 'A. Major Final Outputs (MFOs)/Operations',
  STO = 'B. Support to Operations (STO)',
  GAS = 'C. General Administrative Services (GAS)'
}

export interface Rating {
  q: number | null; // Quality
  e: number | null; // Efficiency
  t: number | null; // Timeliness
  ave: number | null; // Average
}

export interface Accomplishment {
  text: string;
  evidenceFiles: AttachedFile[];
}

export interface AttachedFile {
  id: string;
  name: string;
  url: string;
  dateUploaded: string;
}

export interface IndicatorRow {
  id: string;
  performanceIndicator: string; // Col 1
  budget: string; // Col 2
  successIndicator: string; // Col 3 (Targets)
  responsibleDivision: string; // Col 4
  rating: Rating; // Col 5 (Q, E, T, Ave)
  accomplishment: {
    q1: Accomplishment;
    q2: Accomplishment;
    q3: Accomplishment;
    q4: Accomplishment;
  };
}

export interface CategorySection {
  id: string;
  title: CategoryType;
  rows: IndicatorRow[];
}

export interface OpcrData {
  id: string;
  department: string;
  year: string;
  lastUpdated: string;
  signatories: {
    committing: string;
    approving: string;
    pmt: string;
    head: string;
  };
  sections: CategorySection[];
}