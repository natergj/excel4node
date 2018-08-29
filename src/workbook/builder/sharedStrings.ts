import xmlbuilder from 'xmlbuilder';
import IWorkbookBuilder from '../types/IWorkbookBuilder';
import { getDataStream } from '../../utils/dataStream';
import CTColor from '../../style/CTColor';

export default function addSharedStrings(builder: IWorkbookBuilder) {
  const { wb, xlsx } = builder;

  const dataStream = getDataStream();
  const writer = xmlbuilder.streamWriter(dataStream);

  const xml = xmlbuilder
    .create('sst', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
      allowSurrogateChars: true,
    })
    .att('count', wb.sharedStrings.size)
    .att('uniqueCount', wb.sharedStrings.size)
    .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

  wb.sharedStrings.forEach(s => {
    if (typeof s === 'string') {
      xml
        .ele('si')
        .ele('t')
        .txt(s);
    } else if (s instanceof Array) {
      let thisSI = xml.ele('si');
      let theseRuns = []; // §18.4.4 r (Rich Text Run)
      let currProps = {};
      let curRun;
      let i = 0;
      while (i < s.length) {
        if (typeof s[i] === 'string') {
          if (curRun === undefined) {
            theseRuns.push({
              props: {},
              text: '',
            });
            curRun = theseRuns[theseRuns.length - 1];
          }
          curRun.text = curRun.text + s[i];
        } else if (typeof s[i] === 'object') {
          theseRuns.push({
            props: {},
            text: '',
          });
          curRun = theseRuns[theseRuns.length - 1];
          Object.keys(s[i]).forEach(k => {
            currProps[k] = s[i][k];
          });
          Object.keys(currProps).forEach(k => {
            curRun.props[k] = currProps[k];
          });
          if (s[i].value !== undefined) {
            curRun.text = s[i].value;
          }
        }
        i++;
      }

      theseRuns.forEach(run => {
        if (Object.keys(run).length < 1) {
          thisSI.ele('t', run.text).att('xml:space', 'preserve');
        } else {
          let thisRun = thisSI.ele('r');
          let thisRunProps = thisRun.ele('rPr');
          typeof run.props.name === 'string' ? thisRunProps.ele('rFont').att('val', run.props.name) : null;
          run.props.bold === true ? thisRunProps.ele('b') : null;
          run.props.italics === true ? thisRunProps.ele('i') : null;
          run.props.strike === true ? thisRunProps.ele('strike') : null;
          run.props.outline === true ? thisRunProps.ele('outline') : null;
          run.props.shadow === true ? thisRunProps.ele('shadow') : null;
          run.props.condense === true ? thisRunProps.ele('condense') : null;
          run.props.extend === true ? thisRunProps.ele('extend') : null;
          if (typeof run.props.color === 'string') {
            let thisColor = new CTColor(run.props.color);
            thisColor.addToXMLele(thisRunProps);
          }
          typeof run.props.size === 'number' ? thisRunProps.ele('sz').att('val', run.props.size) : null;
          run.props.underline === true ? thisRunProps.ele('u') : null;
          typeof run.props.vertAlign === 'string' ? thisRunProps.ele('vertAlign').att('val', run.props.vertAlign) : null;
          thisRun.ele('t', run.text).att('xml:space', 'preserve');
        }
      });
    }
  });

  xlsx.folder('xl').file('sharedStrings.xml', dataStream);

  xml.end(writer);
  dataStream.end();
}
