import { Worksheet } from '../../worksheet';
import { IWorkbookBuilder } from '.';

export default async function addContentTypes(builder: IWorkbookBuilder) {
  builder.wb.sheets.forEach((ws: Worksheet) => {
    ws.addSheetToXlsxPackage(builder);
  });
}
