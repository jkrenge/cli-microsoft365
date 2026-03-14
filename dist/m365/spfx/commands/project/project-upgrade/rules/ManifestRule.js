import path from 'path';
import { JsonRule } from '../../JsonRule.js';
export class ManifestRule extends JsonRule {
    get resolutionType() {
        return 'json';
    }
    get file() {
        return '';
    }
    addOccurrence(resolution, filePath, projectPath, node, occurrences) {
        occurrences.push({
            file: path.relative(projectPath, filePath),
            resolution: resolution,
            position: this.getPositionFromNode(node)
        });
    }
}
//# sourceMappingURL=ManifestRule.js.map