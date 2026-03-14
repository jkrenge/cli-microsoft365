import fs from 'fs';
export class ScssFile {
    get source() {
        if (!this._source) {
            try {
                this._source = fs.readFileSync(this.path, 'utf-8');
            }
            catch {
                // Do nothing
            }
        }
        return this._source;
    }
    constructor(path) {
        this.path = path;
    }
}
//# sourceMappingURL=ScssFile.js.map