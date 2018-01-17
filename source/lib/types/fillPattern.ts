function items() {
  this.opts = [
    // §18.18.55 ST_PatternType (Pattern Type)
    'darkDown',
    'darkGray',
    'darkGrid',
    'darkHorizontal',
    'darkTrellis',
    'darkUp',
    'darkVerical',
    'gray0625',
    'gray125',
    'lightDown',
    'lightGray',
    'lightGrid',
    'lightHorizontal',
    'lightTrellis',
    'lightUp',
    'lightVertical',
    'mediumGray',
    'none',
    'solid',
  ];
  this.opts.forEach((o, i) => {
    this[o] = i + 1;
  });
}

items.prototype.validate = function(val) {
  if (this[val] === undefined) {
    const opts = [];
    for (const name in this) {
      if (this.hasOwnProperty(name)) {
        opts.push(name);
      }
    }
    throw new TypeError(
      'Invalid value for ST_PatternType; Value must be one of ' +
        this.opts.join(', '),
    );
  } else {
    return true;
  }
};

export default new items();
