function items() { // As defined in §18.3.1.63 pageSetup (Page Setup Settings)
    this.LETTER_PAPER = 1; // Letter paper (8.5 in. by 11 in.)
    this.LETTER_SMALL_PAPER = 2; // Letter small paper (8.5 in. by 11 in.)
    this.TABLOID_PAPER = 3; // Tabloid paper (11 in. by 17 in.)
    this.LEDGER_PAPER = 4; // Ledger paper (17 in. by 11 in.)
    this.LEGAL_PAPER = 5; // Legal paper (8.5 in. by 14 in.)
    this.STATEMENT_PAPER = 6; // Statement paper (5.5 in. by 8.5 in.)
    this.EXECUTIVE_PAPER = 7; // Executive paper (7.25 in. by 10.5 in.)
    this.A3_PAPER = 8; // A3 paper (297 mm by 420 mm)
    this.A4_PAPER = 9; // A4 paper (210 mm by 297 mm)
    this.A4_SMALL_PAPER = 10; // A4 small paper (210 mm by 297 mm)
    this.A5_PAPER = 11; // A5 paper (148 mm by 210 mm)
    this.B4_PAPER = 12; // B4 paper (250 mm by 353 mm)
    this.B5_PAPER = 13; // B5 paper (176 mm by 250 mm)
    this.FOLIO_PAPER = 14; // Folio paper (8.5 in. by 13 in.)
    this.QUARTO_PAPER = 15; // Quarto paper (215 mm by 275 mm)
    this.STANDARD_PAPER_10_BY_14_IN = 16; // Standard paper (10 in. by 14 in.)
    this.STANDARD_PAPER_11_BY_17_IN = 17; // Standard paper (11 in. by 17 in.)
    this.NOTE_PAPER = 18; // Note paper (8.5 in. by 11 in.)
    this.NUMBER_9_ENVELOPE = 19; // #9 envelope (3.875 in. by 8.875 in.)
    this.NUMBER_10_ENVELOPE = 20; // #10 envelope (4.125 in. by 9.5 in.)
    this.NUMBER_11_ENVELOPE = 21; // #11 envelope (4.5 in. by 10.375 in.)
    this.NUMBER_12_ENVELOPE = 22; // #12 envelope (4.75 in. by 11 in.)
    this.NUMBER_14_ENVELOPE = 23; // #14 envelope (5 in. by 11.5 in.)
    this.C_PAPER = 24; // C paper (17 in. by 22 in.)
    this.D_PAPER = 25; // D paper (22 in. by 34 in.)
    this.E_PAPER = 26; // E paper (34 in. by 44 in.)
    this.DL_PAPER = 27; // DL envelope (110 mm by 220 mm)
    this.C5_ENVELOPE = 28; // C5 envelope (162 mm by 229 mm)
    this.C3_ENVELOPE = 29; // C3 envelope (324 mm by 458 mm)
    this.C4_ENVELOPE = 30; // C4 envelope (229 mm by 324 mm)
    this.C6_ENVELOPE = 31; // C6 envelope (114 mm by 162 mm)
    this.C65_ENVELOPE = 32; // C65 envelope (114 mm by 229 mm)
    this.B4_ENVELOPE = 33; // B4 envelope (250 mm by 353 mm)
    this.B5_ENVELOPE = 34; // B5 envelope (176 mm by 250 mm)
    this.B6_ENVELOPE = 35; // B6 envelope (176 mm by 125 mm)
    this.ITALY_ENVELOPE = 36; // Italy envelope (110 mm by 230 mm)
    this.MONARCH_ENVELOPE = 37; // Monarch envelope (3.875 in. by 7.5 in.).
    this.SIX_THREE_QUARTERS_ENVELOPE = 38; // 6 3/4 envelope (3.625 in. by 6.5 in.)
    this.US_STANDARD_FANFOLD = 39; // US standard fanfold (14.875 in. by 11 in.)
    this.GERMAN_STANDARD_FANFOLD = 40; // German standard fanfold (8.5 in. by 12 in.)
    this.GERMAN_LEGAL_FANFOLD = 41; // German legal fanfold (8.5 in. by 13 in.)
    this.ISO_B4 = 42; // ISO B4 (250 mm by 353 mm)
    this.JAPANESE_DOUBLE_POSTCARD = 43; // Japanese double postcard (200 mm by 148 mm)
    this.STANDARD_PAPER_9_BY_11_IN = 44; // Standard paper (9 in. by 11 in.)
    this.STANDARD_PAPER_10_BY_11_IN = 45; // Standard paper (10 in. by 11 in.)
    this.STANDARD_PAPER_15_BY_11_IN = 46; // Standard paper (15 in. by 11 in.)
    this.INVITE_ENVELOPE = 47; // Invite envelope (220 mm by 220 mm)
    this.LETTER_EXTRA_PAPER = 50; // Letter extra paper (9.275 in. by 12 in.)
    this.LEGAL_EXTRA_PAPER = 51; // Legal extra paper (9.275 in. by 15 in.)
    this.TABLOID_EXTRA_PAPER = 52; // Tabloid extra paper (11.69 in. by 18 in.)
    this.A4_EXTRA_PAPER = 53; // A4 extra paper (236 mm by 322 mm)
    this.LETTER_TRANSVERSE_PAPER = 54; // Letter transverse paper (8.275 in. by 11 in.)
    this.A4_TRANSVERSE_PAPER = 55; // A4 transverse paper (210 mm by 297 mm)
    this.LETTER_EXTRA_TRANSVERSE_PAPER = 56; // Letter extra transverse paper (9.275 in. by 12 in.)
    this.SUPER_A_SUPER_A_A4_PAPER = 57; // SuperA/SuperA/A4 paper (227 mm by 356 mm)
    this.SUPER_B_SUPER_B_A3_PAPER = 58; // SuperB/SuperB/A3 paper (305 mm by 487 mm)
    this.LETTER_PLUS_PAPER = 59; // Letter plus paper (8.5 in. by 12.69 in.)
    this.A4_PLUS_PAPER = 60; // A4 plus paper (210 mm by 330 mm)
    this.A5_TRANSVERSE_PAPER = 61; // A5 transverse paper (148 mm by 210 mm)
    this.JIS_B5_TRANSVERSE_PAPER = 62; // JIS B5 transverse paper (182 mm by 257 mm)
    this.A3_EXTRA_PAPER = 63; // A3 extra paper (322 mm by 445 mm)
    this.A5_EXTRA_PAPER = 64; // A5 extra paper (174 mm by 235 mm)
    this.ISO_B5_EXTRA_PAPER = 65; // ISO B5 extra paper (201 mm by 276 mm)
    this.A2_PAPER = 66; // A2 paper (420 mm by 594 mm)
    this.A3_TRANSVERSE_PAPER = 67; // A3 transverse paper (297 mm by 420 mm)
    this.A3_EXTRA_TRANSVERSE_PAPER = 68; // A3 extra transverse paper (322 mm by 445 mm)

    this.opts = [];
    Object.keys(this).forEach((k) => {
        if (typeof this[k] === 'number') {
            this.opts.push(k);
        }
    });
}


items.prototype.validate = function (val) {
    if (this[val.toUpperCase()] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for PAPER_SIZE; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();