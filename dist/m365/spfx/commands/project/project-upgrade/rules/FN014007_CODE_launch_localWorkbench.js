import path from 'path';
import { JsonRule } from '../../JsonRule.js';
export class FN014007_CODE_launch_localWorkbench extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN014007';
    }
    get title() {
        return 'Local workbench in .vscode/launch.json';
    }
    get description() {
        return `In the .vscode/launch.json file, remove the local workbench launch configuration`;
    }
    get resolution() {
        return ``;
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
            !project.vsCode.launchJson.configurations) {
            return;
        }
        const occurrences = [];
        project.vsCode.launchJson.configurations.forEach((configuration, i) => {
            if (configuration.url &&
                configuration.url.indexOf('/temp/workbench.html') > -1) {
                const node = this.getAstNodeFromFile(project.vsCode.launchJson, `configurations[${i}].url`);
                occurrences.push({
                    file: path.relative(project.path, this.file),
                    resolution: this.resolution,
                    position: this.getPositionFromNode(node)
                });
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN014007_CODE_launch_localWorkbench.js.map