import path from 'path';
import { JsonRule } from '../../JsonRule.js';
export class FN014006_CODE_launch_sourceMapPathOverrides extends JsonRule {
    constructor(overrideKey, overrideValue) {
        super();
        this.overrideKey = overrideKey;
        this.overrideValue = overrideValue;
    }
    get id() {
        return 'FN014006';
    }
    get title() {
        return 'sourceMapPathOverrides in .vscode/launch.json';
    }
    get description() {
        return `In the .vscode/launch.json file, for each configuration, in the sourceMapPathOverrides property, add "${this.overrideKey}": "${this.overrideValue}"`;
    }
    get resolution() {
        return `{
  "configurations": [
    {
      "sourceMapPathOverrides": {
        "${this.overrideKey}": "${this.overrideValue}"
      }
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
            if (configuration.sourceMapPathOverrides &&
                !configuration.sourceMapPathOverrides[this.overrideKey]) {
                const node = this.getAstNodeFromFile(project.vsCode.launchJson, `configurations[${i}].sourceMapPathOverrides`);
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
//# sourceMappingURL=FN014006_CODE_launch_sourceMapPathOverrides.js.map