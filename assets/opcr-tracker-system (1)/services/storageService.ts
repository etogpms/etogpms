import { INITIAL_DATA } from '../constants';
import { OpcrData } from '../types';

const STORAGE_KEY = 'mwss_opcr_data_single_v1';

// Helper to simulate delay
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export const loadOpcrData = async (): Promise<OpcrData> => {
  await delay(600);
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    return JSON.parse(stored);
  }
  return INITIAL_DATA;
};

export const saveOpcrData = async (data: OpcrData): Promise<void> => {
  await delay(400);
  // Update timestamp
  data.lastUpdated = new Date().toISOString();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
};

export const uploadEvidence = async (file: File): Promise<string> => {
  // Simulate file upload
  await delay(1000);
  return URL.createObjectURL(file);
};