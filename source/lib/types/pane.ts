// §18.18.52 ST_Pane (Pane Types)

function items() {
  const opts = ['bottomLeft', 'bottomRight', 'topLeft', 'topRight'];
  opts.forEach((o, i) => {
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
      'Invalid value for sheetview.pane.activePane; Value must be one of ' +
        opts.join(', '),
    );
  } else {
    return true;
  }
};

export default new items();
