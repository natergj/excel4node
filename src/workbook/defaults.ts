import { JSZipGeneratorOptions } from 'jszip';
import { Font, WorkbookView } from '../types/general';

const jsZipDefaults: JSZipGeneratorOptions = {
  compression: 'DEFLATE',
};

export function getJsZipProperties(overrides: Partial<JSZipGeneratorOptions> = {}): JSZipGeneratorOptions {
  return {
    ...jsZipDefaults,
    ...overrides,
  };
}

const defaultFont: Partial<Font> = {
  color: 'FF000000',
  name: 'Calibri',
  size: 12,
  family: 'roman',
};

export function getDefaultFont(overrides: Partial<Font> = {}) {
  return {
    ...defaultFont,
    ...overrides,
  };
}
export const dateFormat = 'm/d/yy';

const workbookView: WorkbookView = {
  // Defaults defined in §18.2.30
  activeTab: 0,
  autoFilterDateGrouping: true,
  firstSheet: 0,
  minimized: false,
  showHorizontalScroll: true,
  showSheetTabs: true,
  showVerticalScroll: true,
  tabRatio: 600,
  visibility: 'visible',
  windowHeight: 17440, // default of Excel 2016 for Mac
  windowWidth: 28040, // default of Excel 2016 for Mac
  xWindow: 5180, // default of Excel 2016 for Mac
  yWindow: 3060, // default of Excel 2016 for Mac
};

export function createWorkbookView(overrides: Partial<WorkbookView> = {}): WorkbookView {
  return {
    ...workbookView,
    ...overrides,
  };
}
