import { Rule } from '../../Rule.js';
import { stringUtil } from '../../../../../../utils/stringUtil.js';
export class FN014003_CODE_launch extends Rule {
    constructor(contents) {
        super();
        this.contents = contents;
    }
    get id() {
        return 'FN014003';
    }
    get title() {
        return '.vscode/launch.json';
    }
    get description() {
        return `In the .vscode folder, add the launch.json file`;
    }
    get resolution() {
        return this.contents;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Recommended';
    }
    get file() {
        return '.vscode/launch.json';
    }
    visit(project, findings) {
        if (!project.vsCode ||
            !project.vsCode.launchJson ||
            stringUtil.normalizeLineEndings(project.vsCode.launchJson.source) !== stringUtil.normalizeLineEndings(this.contents)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN014003_CODE_launch.js.map