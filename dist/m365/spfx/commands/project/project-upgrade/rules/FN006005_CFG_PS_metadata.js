import { JsonRule } from '../../JsonRule.js';
export class FN006005_CFG_PS_metadata extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN006005';
    }
    get title() {
        return 'package-solution.json metadata';
    }
    get description() {
        return `In package-solution.json add metadata section`;
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/package-solution.json';
    }
    visit(project, findings) {
        if (!project.packageSolutionJson ||
            !project.packageSolutionJson.solution) {
            return;
        }
        if (!project.packageSolutionJson.solution.metadata) {
            const solutionDescription = `${project.packageJson?.name} description`;
            const resolution = {
                solution: {
                    metadata: {
                        shortDescription: {
                            default: solutionDescription
                        },
                        longDescription: {
                            default: solutionDescription
                        },
                        screenshotPaths: [],
                        videoUrl: '',
                        categories: []
                    }
                }
            };
            const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution');
            this.addFindingWithOccurrences([{
                    file: this.file,
                    resolution: JSON.stringify(resolution, null, 2),
                    position: this.getPositionFromNode(node)
                }], findings);
        }
    }
}
//# sourceMappingURL=FN006005_CFG_PS_metadata.js.map