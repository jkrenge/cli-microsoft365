import path from 'path';
import { JsonRule } from '../../JsonRule.js';
export class FN014009_CODE_launch_hostedWorkbench_url extends JsonRule {
    constructor(url) {
        super();
        this.url = url;
    }
    get id() {
        return 'FN014009';
    }
    get title() {
        return 'Hosted workbench URL in .vscode/launch.json';
    }
    get description() {
        return `In the .vscode/launch.json file, update the url property for the hosted workbench launch configuration`;
    }
    get resolution() {
        return `{
  "configurations": [
    {
      "url": "${this.url}"
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
                configuration.url !== this.url) {
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
//# sourceMappingURL=FN014009_CODE_launch_hostedWorkbench_url.js.map