import ts from 'typescript';
import { TsRule } from "./TsRule.js";
export class FN016001_TS_msgraphclient_packageName extends TsRule {
    constructor(packageName) {
        super();
        this.packageName = packageName;
    }
    get id() {
        return 'FN016001';
    }
    get title() {
        return 'MSGraphClient package name';
    }
    get description() {
        return `Change the name of the package to import MSGraphClient to '${this.packageName}'`;
    }
    get resolution() {
        return ``;
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
            const msGraphImports = nodes
                .filter(n => ts.isImportSpecifier(n))
                .map(n => n)
                .filter(i => i.name.text === 'MSGraphClient');
            msGraphImports.forEach(msGraphImport => {
                const msGraphImportDeclaration = TsRule.getParentOfType(msGraphImport, ts.isImportDeclaration);
                if (!msGraphImportDeclaration) {
                    return;
                }
                const moduleSpecifier = msGraphImportDeclaration.moduleSpecifier.getText();
                if (moduleSpecifier !== `"${this.packageName}"` &&
                    moduleSpecifier !== `'${this.packageName}'`) {
                    const resolution = msGraphImportDeclaration.getText(msGraphImportDeclaration.getSourceFile()).replace(moduleSpecifier, `"${this.packageName}"`);
                    this.addOccurrence(resolution, file.path, project.path, msGraphImportDeclaration, occurrences);
                }
            });
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN016001_TS_msgraphclient_packageName.js.map