import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { Worksheet } from '../../worksheet';

export default async function addContentTypes(builder: IWorkbookBuilder) {
  builder.wb.sheets.forEach((ws: Worksheet) => {
    ws.addSheetToXlsxPackage(builder);
  });
}
