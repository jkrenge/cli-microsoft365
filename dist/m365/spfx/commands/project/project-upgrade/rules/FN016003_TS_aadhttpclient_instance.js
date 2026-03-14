import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN016003_TS_aadhttpclient_instance extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN016003';
    }
    get title() {
        return 'AadHttpClient instance';
    }
    get description() {
        return `Refactor the code to get AadHttpClient instance`;
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'ts';
    }
    get severity() {
        return 'Required';
    }
    visit(project, findings) {
        if (!project.tsFiles) {
            return;
        }
        const occurrences = [];
        project.tsFiles.forEach(file => {
            const nodes = file.nodes;
            if (!nodes) {
                return;
            }
            const newAddHttpClient = nodes
                .filter(n => ts.isNewExpression(n))
                .map(n => n)
                .filter(n => n.expression.getText() === 'AadHttpClient');
            newAddHttpClient.forEach(n => {
                let resource = undefined;
                if (n.arguments && n.arguments.length === 2) {
                    resource = n.arguments[1];
                }
                const varDec = TsRule.getParentOfType(n, ts.isVariableDeclaration);
                if (varDec) {
                    const resourceString = resource ? resource.getText() : '/* your resource */';
                    const resolution = `this.context.aadHttpClientFactory
  .getClient(${resourceString})
  .then((client: AadHttpClient): void => {
    // use AadHttpClient here
  });`;
                    this.addOccurrence(resolution, file.path, project.path, varDec, occurrences);
                }
            });
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN016003_TS_aadhttpclient_instance.js.map