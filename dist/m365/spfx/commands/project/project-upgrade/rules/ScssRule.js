import path from 'path';
import { Rule } from '../../Rule.js';
export class ScssRule extends Rule {
    get resolutionType() {
        return 'scss';
    }
    get file() {
        return '';
    }
    addOccurrence(resolution, filePath, projectPath, occurrences) {
        occurrences.push({
            file: path.relative(projectPath, filePath),
            resolution: resolution
        });
    }
}
//# sourceMappingURL=ScssRule.js.map