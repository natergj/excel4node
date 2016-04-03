class DefinedName { //§18.2.5 definedName (Defined Name)
    constructor(opts) {
        opts.refFormula !== undefined ? this.refFormula = opts.refFormula : null; 
        opts.name !== undefined ? this.name = opts.name : null;
        opts.comment !== undefined ? this.comment = opts.comment : null; 
        opts.customMenu !== undefined ? this.customMenu = opts.customMenu : null; 
        opts.description !== undefined ? this.description = opts.description : null; 
        opts.help !== undefined ? this.help = opts.help : null; 
        opts.statusBar !== undefined ? this.statusBar = opts.statusBar : null; 
        opts.localSheetId !== undefined ? this.localSheetId = opts.localSheetId : null; 
        opts.hidden !== undefined ? this.hidden = opts.hidden : null; 
        opts['function'] !== undefined ? this['function'] = opts['function'] : null; 
        opts.vbProcedure !== undefined ? this.vbProcedure = opts.vbProcedure : null; 
        opts.xlm !== undefined ? this.xlm = opts.xlm : null; 
        opts.functionGroupId !== undefined ? this.functionGroupId = opts.functionGroupId : null; 
        opts.shortcutKey !== undefined ? this.shortcutKey = opts.shortcutKey : null; 
        opts.publishToServer !== undefined ? this.publishToServer = opts.publishToServer : null; 
        opts.workbookParameter !== undefined ? this.workbookParameter = opts.workbookParameter : null; 
    }

    addToXMLele(ele) {
        let dEle = ele.ele('definedName');
        this.comment !== undefined ? dEle.att('comment', this.comment) : null; 
        this.customMenu !== undefined ? dEle.att('customMenu', this.customMenu) : null; 
        this.description !== undefined ? dEle.att('description', this.description) : null; 
        this.help !== undefined ? dEle.att('help', this.help) : null; 
        this.statusBar !== undefined ? dEle.att('statusBar', this.statusBar) : null; 
        this.hidden !== undefined ? dEle.att('hidden', this.hidden) : null; 
        this.localSheetId !== undefined ? dEle.att('localSheetId', this.localSheetId) : null; 
        this.name !== undefined ? dEle.att('name', this.name) : null;
        this['function'] !== undefined ? dEle.att('function', this['function']) : null; 
        this.vbProcedure !== undefined ? dEle.att('vbProcedure', this.vbProcedure) : null; 
        this.xlm !== undefined ? dEle.att('xlm', this.xlm) : null; 
        this.functionGroupId !== undefined ? dEle.att('functionGroupId', this.functionGroupId) : null; 
        this.shortcutKey !== undefined ? dEle.att('shortcutKey', this.shortcutKey) : null; 
        this.publishToServer !== undefined ? dEle.att('publishToServer', this.publishToServer) : null; 
        this.workbookParameter !== undefined ? dEle.att('workbookParameter', this.workbookParameter) : null; 

        this.refFormula !== undefined ? dEle.text(this.refFormula) : null;
    }
}


class DefinedNameCollection { // §18.2.6 definedNames (Defined Names)
    constructor() {
        this.items = [];
    }

    get length() {
        return this.items.length;
    }

    get isEmpty() {
        if (this.items.length === 0) {
            return true;
        } else {
            return false;
        }
    }

    addDefinedName(opts) {
        let item = new DefinedName(opts);
        let newLength = this.items.push(item);
        return this.items[newLength - 1];
    }

    addToXMLele(ele) {
        let dnEle = ele.ele('definedNames');
        this.items.forEach((dn) => {
            dn.addToXMLele(dnEle);
        });
    }
}
module.exports = DefinedNameCollection;