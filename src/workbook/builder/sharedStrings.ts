import xmlbuilder from 'xmlbuilder';
import { getDataStream } from '../../utils/dataStream';
import CTColor from '../../style/CTColor';
import { IWorkbookBuilder } from '.';

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

  wb.sharedStrings.forEach((index, str) => {
    if (typeof str === 'string') {
      xml
        .ele('si')
        .ele('t')
        .txt(str);
    } else if (str instanceof Array) {
      const thisSI = xml.ele('si');
      const theseRuns = []; // §18.4.4 r (Rich Text Run)
      const currProps = {};
      let curRun;
      let i = 0;
      while (i < str.length) {
        if (typeof str[i] === 'string') {
          if (curRun === undefined) {
            theseRuns.push({ props: {}, text: '' });
            curRun = theseRuns[theseRuns.length - 1];
          }
          curRun.text = curRun.text + str[i];
        } else if (typeof str[i] === 'object') {
          theseRuns.push({ props: {}, text: '' });
          curRun = theseRuns[theseRuns.length - 1];
          Object.keys(str[i]).forEach(k => {
            currProps[k] = str[i][k];
          });
          Object.keys(currProps).forEach(k => {
            curRun.props[k] = currProps[k];
          });
          if (str[i].value !== undefined) {
            curRun.text = str[i].value;
          }
        }
        i++;
      }

      theseRuns.forEach(run => {
        if (Object.keys(run).length < 1) {
          thisSI.ele('t', run.text).att('xml:space', 'preserve');
        } else {
          const thisRun = thisSI.ele('r');
          const thisRunProps = thisRun.ele('rPr');
          if (typeof run.props.name === 'string') {
            thisRunProps.ele('rFont').att('val', run.props.name);
          }
          if (run.props.bold === true) {
            thisRunProps.ele('b');
          }
          if (run.props.italics === true) {
            thisRunProps.ele('i');
          }
          if (run.props.strike === true) {
            thisRunProps.ele('strike');
          }
          if (run.props.outline === true) {
            thisRunProps.ele('outline');
          }
          if (run.props.shadow === true) {
            thisRunProps.ele('shadow');
          }
          if (run.props.condense === true) {
            thisRunProps.ele('condense');
          }
          if (run.props.extend === true) {
            thisRunProps.ele('extend');
          }
          if (typeof run.props.color === 'string') {
            const thisColor = new CTColor(run.props.color);
            thisColor.addToXMLele(thisRunProps);
          }
          if (typeof run.props.size === 'number') {
            thisRunProps.ele('sz').att('val', run.props.size);
          }
          if (run.props.underline === true) {
            thisRunProps.ele('u');
          }
          if (typeof run.props.vertAlign === 'string') {
            thisRunProps.ele('vertAlign').att('val', run.props.vertAlign);
          }
          thisRun.ele('t', run.text).att('xml:space', 'preserve');
        }
      });
    }
  });

  xlsx.folder('xl').file('sharedStrings.xml', dataStream);

  xml.end(writer);
  dataStream.end();
}
