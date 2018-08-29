// subset of §20.1.10.48 ST_PresetColorVal (Preset Color Value)
export enum EXCEL_COLOR {
  aqua = 'FF33CCCC',
  black = 'FF000000',
  blue = 'FF0000FF',
  bluegray = 'FF666699',
  brightGreen = 'FF00FF00',
  brown = 'FF993300',
  darkBlue = 'FF000080',
  darkGreen = 'FF003300',
  darkRed = 'FF800000',
  darkTeal = 'FF003366',
  darkYellow = 'FF808000',
  gold = 'FFFFCC00',
  gray25 = 'FFC0C0C0',
  gray40 = 'FF969696',
  gray50 = 'FF808080',
  gray80 = 'FF333333',
  green = 'FF008000',
  indigo = 'FF333399',
  lavender = 'FFCC99FF',
  lightBlue = 'FF3366FF',
  lightGreen = 'FFCCFFCC',
  lightOrange = 'FFFF9900',
  lightTurquoise = 'FFCCFFFF',
  lightYellow = 'FFFFFF99',
  lime = 'FF99CC00',
  oliveGreen = 'FF333300',
  orange = 'FFFF6600',
  paleBlue = 'FF99CCFF',
  pink = 'FFFF00FF',
  plum = 'FF993366',
  red = 'FFFF0000',
  rose = 'FFFF99CC',
  seaGreen = 'FF339966',
  skyBlue = 'FF00CCFF',
  tan = 'FFFFCC99',
  teal = 'FF008080',
  turquoise = 'FF00FFFF',
  violet = 'FF800080',
  white = 'FFFFFFFF',
  yellow = 'FFFFFF00',
}

export function getColor(val) {
  // check for RGB, RGBA or Excel Color Names and return RGBA
  if (typeof this[val.toLowerCase()] === 'string') {
    // val was a named color that matches predefined list. return corresponding color
    return this[val.toLowerCase()];
  } else if (val.length === 8 && /^[a-fA-F0-9()]+$/.test(val)) {
    // val is already a properly formatted color string, return upper case version of itself
    return val.toUpperCase();
  } else if (val.length === 6 && /^[a-fA-F0-9()]+$/.test(val)) {
    // val is color code without Alpha, add it and return
    return 'FF' + val.toUpperCase();
  } else if (val.length === 7 && val.substr(0, 1) === '#' && /^[a-fA-F0-9()]+$/.test(val.substr(1))) {
    // val was sent as html style hex code, remove # and add alpha
    return 'FF' + val.substr(1).toUpperCase();
  } else {
    // I don't know what this is, return valid color and console.log error
    return this['white'];
  }
}
