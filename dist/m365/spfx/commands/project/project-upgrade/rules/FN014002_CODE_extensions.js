import { Rule } from '../../Rule.js';
export class FN014002_CODE_extensions extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN014002';
    }
    get title() {
        return '.vscode/extensions.json';
    }
    get description() {
        return `In the .vscode folder, add the extensions.json file`;
    }
    get resolution() {
        return `{
  "recommendations": [
    "msjsdiag.debugger-for-chrome"
  ]
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Recommended';
    }
    get file() {
        return '.vscode/extensions.json';
    }
    visit(project, findings) {
        if (!project.vsCode || !project.vsCode.extensionsJson) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN014002_CODE_extensions.js.map