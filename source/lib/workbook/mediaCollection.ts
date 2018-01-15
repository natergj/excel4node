import * as fs from "fs";

export default class MediaCollection {
  private items;

  constructor() {
    this.items = [];
  }

  add(item) {
    if (typeof item === "string") {
      fs.accessSync(item, fs.constants.R_OK);
    }

    this.items.push(item);
    return this.items.length;
  }

  get isEmpty() {
    if (this.items.length === 0) {
      return true;
    } else {
      return false;
    }
  }
}
