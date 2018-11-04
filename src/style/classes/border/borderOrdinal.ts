import { DataBarColor } from '../color';

enum BorderOrdinalStyle {
  'dashDot' = 'dashDot',
  'dashDotDot' = 'dashDotDot',
  'dashed' = 'dashed',
  'dotted' = 'dotted',
  'double' = 'double',
  'hair' = 'hair',
  'medium' = 'medium',
  'mediumDashDot' = 'mediumDashDot',
  'mediumDashDotDot' = 'mediumDashDotDot',
  'mediumDashed' = 'mediumDashed',
  'none' = 'none',
  'slantDashDot' = 'slantDashDot',
  'thick' = 'thick',
  'thin' = 'thin',
}

export class BorderOrdinal {
  style?: BorderOrdinalStyle;
  color?: DataBarColor;
}
