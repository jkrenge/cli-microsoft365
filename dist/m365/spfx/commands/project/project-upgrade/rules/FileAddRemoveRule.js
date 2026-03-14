import fs from 'fs';
import path from 'path';
import { Rule } from '../../Rule.js';
export class FileAddRemoveRule extends Rule {
    constructor(filePath, add, contents) {
        super();
        this.filePath = filePath;
        this.add = add;
        this.contents = contents;
        this.filePath = path.normalize(this.filePath);
    }
    get title() {
        return this.filePath;
    }
    get description() {
        return `${this.add ? 'Add' : 'Remove'} file ${this.filePath}`;
    }
    get resolution() {
        if (this.add) {
            return `add_cmd[BEFOREPATH]"${this.filePath}"[AFTERPATH][BEFORECONTENT]
${this.contents}
[AFTERCONTENT]`;
        }
        else {
            return `remove_cmd "${this.filePath}"`;
        }
    }
    get resolutionType() {
        return 'cmd';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return this.filePath;
    }
    visit(project, notifications) {
        const targetPath = path.join(project.path, this.filePath);
        if ((!this.add && fs.existsSync(targetPath)) ||
            (this.add && !fs.existsSync(targetPath))) {
            this.addFinding(notifications);
            return;
        }
        if (this.add && this.contents) {
            const fileContent = fs.readFileSync(path.join(project.path, this.filePath), 'utf8');
            if (fileContent !== this.contents) {
                this.addFinding(notifications);
            }
        }
    }
}
//# sourceMappingURL=FileAddRemoveRule.js.map