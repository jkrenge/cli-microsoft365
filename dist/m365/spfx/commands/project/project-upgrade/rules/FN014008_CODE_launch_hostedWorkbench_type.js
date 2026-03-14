import path from 'path';
import { JsonRule } from '../../JsonRule.js';
export class FN014008_CODE_launch_hostedWorkbench_type extends JsonRule {
    constructor(type) {
        super();
        this.type = type;
    }
    get id() {
        return 'FN014008';
    }
    get title() {
        return 'Hosted workbench type in .vscode/launch.json';
    }
    get description() {
        return `In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration`;
    }
    get resolution() {
        return `{
  "configurations": [
    {
      "type": "${this.type}"
    }
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
            if (configuration.name === 'Hosted workbench' &&
                configuration.type !== this.type) {
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
//# sourceMappingURL=FN014008_CODE_launch_hostedWorkbench_type.js.map