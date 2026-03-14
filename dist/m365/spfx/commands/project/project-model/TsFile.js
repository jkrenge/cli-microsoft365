import fs from 'fs';
import path from 'path';
import ts from 'typescript';
import { tsUtil } from '../../../../../utils/tsUtil.js';
export class TsFile {
    get sourceFile() {
        if (!this._sourceFile) {
            if (!this.source) {
                return undefined;
            }
            try {
                this._sourceFile = tsUtil.createSourceFile(path.basename(this.path), this.source, ts.ScriptTarget.Latest, true);
            }
            catch {
                // Do nothing
            }
        }
        return this._sourceFile;
    }
    get nodes() {
        if (!this._nodes) {
            if (!this.sourceFile) {
                return undefined;
            }
            this._nodes = TsFile.getAsEnumerable(this.sourceFile, this.sourceFile);
        }
        return this._nodes;
    }
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
    static getAsEnumerable(file, node) {
        const nodes = [node];
        node.getChildren(file).forEach(n => {
            nodes.push(...TsFile.getAsEnumerable(file, n));
        });
        return nodes;
    }
}
//# sourceMappingURL=TsFile.js.map