export enum ST_Objects {
  ALL = 'all',
  NONE = 'none',
  PLACEHOLDERS = 'placeholders',
}

export enum ST_UpdateLinks {
  ALWAYS = 'always',
  NEVER = 'never',
  USERSET = 'userSet',
}

export enum FontCharset {
  ANSI_CHARSET = 0,
  DEFAULT_CHARSET = 1,
  SYMBOL_CHARSET = 2,
  MAC_CHARSET = 77,
  SHIFTJIS_CHARSET = 128,
  HANGEUL_CHARSET = 129,
  HANGUL_CHARSET = 129,
  JOHAB_CHARSET = 130,
  GB2312_CHARSET = 134,
  CHINESEBIG5_CHARSET = 136,
  GREEK_CHARSET = 161,
  TURKISH_CHARSET = 162,
  VIETNAMESE_CHARSET = 163,
  HEBREW_CHARSET = 177,
  ARABIC_CHARSET = 178,
  BALTIC_CHARSET = 186,
  RUSSIAN_CHARSET = 204,
  THAI_CHARSET = 222,
  EASTEUROPE_CHARSET = 238,
  OEM_CHARSET = 255,
}

export type DataBarColor = {
  auto: boolean;
  indexed: number;
  rgb: string;
  theme: number;
  tint: number;
};

export type FontFamily = 'not_applicable' | 'roman' | 'swiss' | 'modern' | 'script' | 'decorative';

export type ST_Visibility = 'hidden' | 'veryHidden' | 'visible';

// §18.18.33
export type ST_FontScheme = 'none' | 'major' | 'minor';

export enum FontVertAlign {
  BASELINE = 'baseline',
  SUBSCRIPT = 'subscript',
  SUPERSCRIPT = 'superscript',
}

// Derived from §18.8.22
export type Font = {
  bold: boolean;
  charset: FontCharset;
  color: string | DataBarColor;
  condense: boolean;
  extend: boolean;
  family: FontFamily;
  italic: boolean;
  name: string;
  outline: boolean;
  scheme: ST_FontScheme;
  shadow: boolean;
  strike: boolean;
  size: number;
  underline: boolean;
  vertAlign: FontVertAlign;
};

// §18.2.30
export type WorkbookView = {
  activeTab?: number;
  autoFilterDateGrouping?: boolean;
  firstSheet?: number;
  minimized?: boolean;
  showHorizontalScroll?: boolean;
  showSheetTabs?: boolean;
  showVerticalScroll?: boolean;
  tabRatio?: number;
  visibility?: ST_Visibility;
  windowHeight?: number;
  windowWidth?: number;
  xWindow?: number;
  yWindow?: number;
  showComments?: boolean;
};

export type WorkbookProperties = {
  allowRefreshQuery: boolean;
  autoCompressPictures: boolean;
  backupFile: boolean;
  checkCompatibility: boolean;
  codeName: string;
  date1904: boolean; // See §18.17.4.1
  dateCompatibility: boolean; // See §18.17.4.1
  // defaultThemeVersion: string; // dependant on opening application; not supported in excel4node library.
  filterPrivacy: boolean;
  hidePivotFieldList: boolean;
  promptedSolutions: boolean;
  publishItems: boolean;
  // refreshAllConnections: boolean; // external data sources not supported by excel4node library
  // saveExternalLinkValues: boolean; // external data sources not supported by excel4node library
  showBorderUnselectedTables: boolean;
  showInkAnnotation: boolean;
  showObjects: ST_Objects;
  showPivotChartFilter: boolean;
  updateLinks: ST_UpdateLinks;
};
