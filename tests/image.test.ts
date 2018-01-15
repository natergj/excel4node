import * as test from "tape";
import xl from "../source/index";
import Picture from "../source/lib/drawing/picture";
import * as path from "path";
import * as fs from "fs";

test("Test adding images", t => {
  var wb = new xl.Workbook();
  var ws = wb.addWorksheet("test 1");

  ws.addImage({
    path: path.resolve(__dirname, "../sampleFiles/thumbs-up.jpg"),
    type: "picture",
    position: {
      type: "absoluteAnchor",
      x: "1in",
      y: "2in",
    },
  });

  ws.addImage({
    path: path.resolve(__dirname, "../sampleFiles/logo.png"),
    type: "picture",
    position: {
      type: "oneCellAnchor",
      from: {
        col: 1,
        colOff: "0.5in",
        row: 1,
        rowOff: 0,
      },
    },
  });

  ws.addImage({
    image: fs.readFileSync(path.resolve(__dirname, "../sampleFiles/logo.png")),
    type: "picture",
    fileName: "logo.png",
    position: {
      type: "twoCellAnchor",
      from: {
        col: 1,
        colOff: 0,
        row: 10,
        rowOff: 0,
      },
      to: {
        col: 4,
        colOff: 0,
        row: 13,
        rowOff: 0,
      },
    },
  });

  let pics = ws.drawingCollection.drawings;
  t.ok(
    pics[0] instanceof Picture &&
      pics[1] instanceof Picture &&
      pics[2] instanceof Picture,
    "3 new picture successfully created"
  );

  try {
    ws.addImage({
      path: path.resolve(__dirname, "../sampleFiles/logo.png"),
      type: "picture",
      position: {
        type: "twoCellAnchor",
        from: {
          col: 1,
          colOff: 0,
          row: 10,
          rowOff: 0,
        },
      },
    });
    t.notOk(
      pics[3] instanceof Picture,
      "Adding twoCellAnchor picture without specifying to position should throw error"
    );
  } catch (e) {
    t.ok(
      e instanceof TypeError,
      "Adding twoCellAnchor picture without specifying to position should throw error"
    );
  }

  t.end();
});
