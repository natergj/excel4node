import Drawing from "./drawing";
import Picture from "./picture";

class DrawingCollection {
  public drawings;

  constructor() {
    this.drawings = [];
  }

  get length() {
    return this.drawings.length;
  }

  add(opts) {
    switch (opts.type) {
      case "picture":
        let newPic = new Picture(opts);
        this.drawings.push(newPic);
        return newPic;

      default:
        throw new TypeError("this option is not yet supported");
    }
  }

  get isEmpty() {
    if (this.drawings.length === 0) {
      return true;
    } else {
      return false;
    }
  }
}

export default { DrawingCollection, Drawing, Picture };
