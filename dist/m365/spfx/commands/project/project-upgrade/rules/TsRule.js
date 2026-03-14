import path from 'path';
import { Rule } from '../../Rule.js';
export class TsRule extends Rule {
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'ts';
    }
    get file() {
        return '';
    }
    addOccurrence(resolution, filePath, projectPath, node, occurrences) {
        const lineChar = node.getSourceFile().getLineAndCharacterOfPosition(node.getStart());
        occurrences.push({
            file: path.relative(projectPath, filePath),
            position: {
                line: lineChar.line + 1,
                character: lineChar.character + 1
            },
            resolution: resolution
        });
    }
    static getParentOfType(node, typeComparer) {
        const parent = node.parent;
        if (!parent) {
            return undefined;
        }
        if (typeComparer(parent)) {
            return parent;
        }
        return TsRule.getParentOfType(parent, typeComparer);
    }
}
//# sourceMappingURL=TsRule.js.map