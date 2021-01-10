class ChartsCollection {
    constructor() {
        this.items = [];
    }

    add(item) {
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

module.exports = ChartsCollection;