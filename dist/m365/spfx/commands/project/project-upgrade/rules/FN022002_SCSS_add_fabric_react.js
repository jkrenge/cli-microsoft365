import { spfx } from '../../../../../../utils/spfx.js';
import { ScssRule } from './ScssRule.js';
export class FN022002_SCSS_add_fabric_react extends ScssRule {
    constructor(importValue, addIfContains) {
        super();
        this.importValue = importValue;
        this.addIfContains = addIfContains;
    }
    get id() {
        return 'FN022002';
    }
    get title() {
        return `Scss file import`;
    }
    get description() {
        return `Add scss file import`;
    }
    get resolution() {
        return `@import '${this.importValue}'`;
    }
    get resolutionType() {
        return 'scss';
    }
    get severity() {
        return 'Optional';
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
            const source = file.source;
            if (source.indexOf(this.importValue) === -1) {
                if (!this.addIfContains || source.indexOf(this.addIfContains) > -1) {
                    this.addOccurrence(this.resolution, file.path, project.path, occurrences);
                }
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN022002_SCSS_add_fabric_react.js.map