import * as path from 'path';
import * as fs from 'fs';
import * as TJS from 'typescript-json-schema';

const settings: TJS.PartialArgs = {
  required: true,
};

const compilerOptions: TJS.CompilerOptions = {
  strictNullChecks: true,
  required: true,
};

const entryFile = path.resolve(__dirname, '../src/index.ts');
const program = TJS.getProgramFromFiles([entryFile], compilerOptions);
const generator = TJS.buildGenerator(program, settings);

if (generator) {
  const symbols = generator.getSymbols();
  symbols.forEach(symbol => {
    if (
      symbol.fullyQualifiedName.includes('excel4node') &&
      symbol.fullyQualifiedName.includes('types') &&
      !symbol.fullyQualifiedName.includes('node_modules')
    ) {
      const outFile = getSchemaPath(symbol.fullyQualifiedName);
      const schema = TJS.generateSchema(program, symbol.name, settings);
      fs.writeFileSync(outFile, JSON.stringify(schema, null, '  '));
    }
  });
}
process.exit();

function getSchemaPath(fullyQualifiedName: string): string {
  const path = fullyQualifiedName.match(/"(.*?)"/);
  if (!path) {
    return '';
  }
  const parts = path[0].split('/');
  parts[parts.indexOf('types')] = 'schemas';
  return `${parts.join('/').replace(/"/g, '')}.schema.json`;
}
