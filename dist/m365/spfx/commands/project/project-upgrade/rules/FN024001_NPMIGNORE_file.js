import { Rule } from '../../Rule.js';
export class FN024001_NPMIGNORE_file extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN024001';
    }
    get title() {
        return `Create .npmignore`;
    }
    get description() {
        return `Create the .npmignore file`;
    }
    get resolution() {
        return `!dist
config

gulpfile.js

release
src
temp

tsconfig.json
tslint.json

*.log

.yo-rc.json
.vscode
`;
    }
    get resolutionType() {
        return 'text';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './.npmignore';
    }
    visit(project, findings) {
        if (!project.npmignore) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN024001_NPMIGNORE_file.js.map