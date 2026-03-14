import { spfx } from '../../../../../../utils/spfx.js';
import { ScssRule } from './ScssRule.js';
export class FN022001_SCSS_remove_fabric_react extends ScssRule {
    constructor(importValue) {
        super();
        this.importValue = importValue;
    }
    get id() {
        return 'FN022001';
    }
    get title() {
        return `Scss file import`;
    }
    get description() {
        return `Remove scss file import`;
    }
    get resolution() {
        return `@import '${this.importValue}'`;
    }
    get resolutionType() {
        return 'scss';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return '';
    }
    visit(project, findings) {
        if (spfx.isReactProject(project) === false) {
            return;
        }
        if (!project.scssFiles ||
            project.scssFiles.length === 0) {
            return;
        }
        const occurrences = [];
        project.scssFiles.forEach((file) => {
            if (file.source.indexOf(this.importValue) !== -1) {
                this.addOccurrence(this.resolution, file.path, project.path, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN022001_SCSS_remove_fabric_react.js.map