export interface IWorksheetOptsHeaderFooter {
  evenFooter: string; //  Even footer text
  evenHeader: string; //  Even header text
  firstFooter: string; //  First footer text
  firstHeader: string; //  First header text
  oddFooter: string; //  Odd footer text
  oddHeader: string; //  Odd header text
  alignWithMargins: boolean; //  Should header/footer align with margins
  differentFirst: boolean; //  Should header/footer show a different header/footer on first page
  differentOddEven: boolean; //  Should header/footer show a different header/footer on odd and even pages
  scaleWithDoc: boolean; //  Should header/footer scale when doc zoom is changed
}
