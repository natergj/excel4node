# Font

```
  type Font = {
    bold: boolean;
    charset: FontCharset;
    color: string | DataBarColor;
    condense: boolean;
    extend: boolean;
    family: FontFamily;
    italic: boolean;
    name: string;
    outline: boolean;
    scheme: ST_FontScheme;
    shadow: boolean;
    strike: boolean;
    size: number;
    underline: boolean;
    vertAlign: FontVertAlign;
  }
```

# FontFamily

```
type FontFamily = 'not_applicable' | 'roman' | 'swiss' | 'modern' | 'script' | 'decorative';
```

# FontCharset

```
enum FontCharset {
  ANSI_CHARSET = 0,
  DEFAULT_CHARSET = 1,
  SYMBOL_CHARSET = 2,
  MAC_CHARSET = 77,
  SHIFTJIS_CHARSET = 128,
  HANGEUL_CHARSET = 129,
  HANGUL_CHARSET = 129,
  JOHAB_CHARSET = 130,
  GB2312_CHARSET = 134,
  CHINESEBIG5_CHARSET = 136,
  GREEK_CHARSET = 161,
  TURKISH_CHARSET = 162,
  VIETNAMESE_CHARSET = 163,
  HEBREW_CHARSET = 177,
  ARABIC_CHARSET = 178,
  BALTIC_CHARSET = 186,
  RUSSIAN_CHARSET = 204,
  THAI_CHARSET = 222,
  EASTEUROPE_CHARSET = 238,
  OEM_CHARSET = 255,
}
```

# DataBarColor

```
type DataBarColor = {
  auto: boolean;
  indexed: number;
  rgb: string;
  theme: number;
  tint: number;
};
```

# FontFamily

```
type FontFamily = 'not_applicable' | 'roman' | 'swiss' | 'modern' | 'script' | 'decorative';
```

# ST_Visibility

```
type ST_Visibility = 'hidden' | 'veryHidden' | 'visible';
```
