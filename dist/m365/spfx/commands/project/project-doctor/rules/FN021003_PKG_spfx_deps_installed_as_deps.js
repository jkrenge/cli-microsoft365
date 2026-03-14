import { JsonRule } from '../../JsonRule.js';
import * as spfxDeps from '../spfx-deps.js';
export class FN021003_PKG_spfx_deps_installed_as_deps extends JsonRule {
    get id() {
        return 'FN021003';
    }
    get title() {
        return '';
    }
    get description() {
        return '';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './package.json';
    }
    get resolutionType() {
        return 'cmd';
    }
    visit(project, findings) {
        if (!project.version ||
            !project.packageJson ||
            !project.packageJson.devDependencies) {
            return;
        }
        const projectDevDeps = Object.keys(project.packageJson.devDependencies);
        projectDevDeps.forEach(dep => {
            if (!spfxDeps.deps.includes(dep)) {
                return;
            }
            const node = this.getAstNodeFromFile(project.packageJson, `devDependencies.${dep}`);
            this.addFindingWithCustomInfo(`${dep} installed as devDependency`, `Package ${dep} is installed as a devDependency. Install it as a dependency instead`, [{
                    file: this.file,
                    resolution: `install ${dep}@${project.version}`,
                    position: this.getPositionFromNode(node)
                }], findings);
        });
    }
}
//# sourceMappingURL=FN021003_PKG_spfx_deps_installed_as_deps.js.map