import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN016002_TS_msgraphclient_instance extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN016002';
    }
    get title() {
        return 'MSGraphClient instance';
    }
    get description() {
        return `Refactor the code to get MSGraphClient instance`;
    }
    get resolution() {
        return `this.context.msGraphClientFactory
  .getClient()
  .then((client: MSGraphClient): void => {
    // use MSGraphClient here
  });`;
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
            const propertyAccessExpressions = nodes
                .filter(n => ts.isPropertyAccessExpression(n))
                .map(n => n)
                .filter(p => p.name.text === 'consume' && p.expression.getText().endsWith('.serviceScope'));
            propertyAccessExpressions.forEach(p => {
                const callExpression = TsRule.getParentOfType(p, ts.isCallExpression);
                if (!callExpression || callExpression.arguments.length < 1) {
                    return;
                }
                if (!ts.isPropertyAccessExpression(callExpression.arguments[0])) {
                    return;
                }
                const prop = callExpression.arguments[0];
                if (prop.expression.getText() !== 'MSGraphClient' || prop.name.text !== 'serviceKey') {
                    return;
                }
                const varDec = TsRule.getParentOfType(p, ts.isVariableDeclaration);
                if (varDec) {
                    this.addOccurrence(this.resolution, file.path, project.path, varDec, occurrences);
                }
            });
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN016002_TS_msgraphclient_instance.js.map