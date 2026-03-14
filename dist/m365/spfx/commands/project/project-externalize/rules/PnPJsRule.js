import { BasicDependencyRule } from "./BasicDependencyRule.js";
export class PnPJsRule extends BasicDependencyRule {
    constructor() {
        super(...arguments);
        this.pnpModules = [
            {
                key: "@pnp/odata",
                globalName: "pnp.odata",
                globalDependencies: [
                    "@pnp/common",
                    "@pnp/logging",
                    "tslib"
                ],
                shadowRequire: "require(\"@pnp/odata\");"
            },
            {
                key: "@pnp/common",
                globalName: "pnp.common",
                shadowRequire: "require(\"@pnp/common\");"
            },
            {
                key: "@pnp/logging",
                globalName: "pnp.logging",
                globalDependencies: [
                    "tslib"
                ],
                shadowRequire: "require(\"@pnp/logging\");"
            },
            {
                key: "@pnp/sp",
                globalName: "pnp.sp",
                globalDependencies: [
                    "@pnp/logging",
                    "@pnp/common",
                    "@pnp/odata",
                    "tslib"
                ]
            },
            {
                key: '@pnp/pnpjs',
                globalName: 'pnp'
            }
        ];
    }
    async visit(project) {
        const findings = this.pnpModules
            .map(x => this.getModuleAndParents(project, x.key))
            .reduce((x, y) => [...x, ...y]);
        const files = this.getEntryFilesList(project);
        const rawFileEdits = this.pnpModules.filter(x => findings.find(y => y.key === x.key) !== undefined)
            .filter(x => x.shadowRequire !== undefined)
            .map(x => files.map(y => ({
            action: "add",
            path: y,
            targetValue: x.shadowRequire
        })));
        const fileEdits = rawFileEdits.length > 0 ? rawFileEdits.reduce((x, y) => [...x, ...y]) : [];
        if (findings.filter(x => x.key && x.key !== '@pnp/pnpjs').length > 0) { // we're adding tslib only if we found other packages that are not the bundle which already contains tslib
            findings.push({
                key: 'tslib',
                globalName: 'tslib',
                path: `https://unpkg.com/tslib@^1.10.0/tslib.js`
            });
            fileEdits.push(...files.map(x => ({
                action: "add",
                path: x,
                targetValue: 'require("tslib");'
            })));
        }
        return { entries: findings, suggestions: fileEdits };
    }
    getEntryFilesList(project) {
        return project && project.manifests ? project.manifests.map(x => x.path.replace('.manifest.json', '.ts')) : [];
    }
    getModuleAndParents(project, moduleName) {
        const result = [];
        const moduleConfiguration = this.pnpModules.find(x => x.key === moduleName);
        if (project.packageJson && moduleConfiguration) {
            const version = project.packageJson.dependencies?.[moduleName];
            if (version) {
                result.push({
                    ...moduleConfiguration,
                    path: `https://unpkg.com/${moduleConfiguration.key}@${version}/dist/${moduleName.replace('@pnp/', '')}.es5.umd${moduleName === '@pnp/common' || moduleName === ' @pnp/pnpjs' ? '.bundle' : ''}.min.js`
                });
                if (moduleConfiguration.globalDependencies) {
                    moduleConfiguration.globalDependencies.forEach(dependency => {
                        result.push(...this.getModuleAndParents(project, `@${dependency.replace('/', '.')}`));
                    });
                }
            }
        }
        return result;
    }
}
//# sourceMappingURL=PnPJsRule.js.map