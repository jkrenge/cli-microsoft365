var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _DependencyRule_instances, _DependencyRule_needsUpdate, _DependencyRule_getMaxVersion;
import semver from 'semver';
import { JsonRule } from '../../JsonRule.js';
export class DependencyRule extends JsonRule {
    constructor(packageName, packageVersion, isDevDep = false, isOptional = false, add = true) {
        super();
        _DependencyRule_instances.add(this);
        this.packageName = packageName;
        this.packageVersion = packageVersion;
        this.isDevDep = isDevDep;
        this.isOptional = isOptional;
        this.add = add;
    }
    get title() {
        return this.packageName;
    }
    get description() {
        return `${(this.add ? 'Upgrade' : 'Remove')} SharePoint Framework ${(this.isDevDep ? 'dev ' : '')}dependency package ${this.packageName}`;
    }
    get resolution() {
        return this.add ?
            `${(this.isDevDep ? 'installDev' : 'install')} ${this.packageName}@${this.packageVersion}` :
            `${(this.isDevDep ? 'uninstallDev' : 'uninstall')} ${this.packageName}`;
    }
    get resolutionType() {
        return 'cmd';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './package.json';
    }
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    customCondition(project) {
        return false;
    }
    visit(project, findings) {
        if (!project.packageJson) {
            return;
        }
        const projectDependencies = this.isDevDep ? project.packageJson.devDependencies : project.packageJson.dependencies;
        const versionEntry = projectDependencies ? projectDependencies[this.packageName] : '';
        if (this.add) {
            let jsonProperty = this.isDevDep ? 'devDependencies' : 'dependencies';
            if (versionEntry) {
                jsonProperty += `.${this.packageName}`;
                if (__classPrivateFieldGet(this, _DependencyRule_instances, "m", _DependencyRule_needsUpdate).call(this, this.packageVersion, versionEntry)) {
                    const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
                    this.addFindingWithPosition(findings, node);
                }
            }
            else {
                if (!this.isOptional || this.customCondition(project)) {
                    const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
                    this.addFindingWithCustomInfo(this.packageName, this.description.replace('Upgrade', 'Install'), [{
                            file: this.file,
                            resolution: this.resolution,
                            position: this.getPositionFromNode(node)
                        }], findings);
                }
            }
        }
        else {
            const jsonProperty = `${(this.isDevDep ? 'devDependencies' : 'dependencies')}.${this.packageName}`;
            if (versionEntry) {
                const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
_DependencyRule_instances = new WeakSet(), _DependencyRule_needsUpdate = function _DependencyRule_needsUpdate(ruleVersion, currentVersion) {
    try {
        // Get minimum versions for both
        const ruleMin = semver.minVersion(ruleVersion);
        const currentMin = semver.minVersion(currentVersion);
        // Check if ranges overlap
        const rangesOverlap = semver.intersects(ruleVersion, currentVersion);
        if (rangesOverlap) {
            // Even if they overlap, update if rule requires a higher minimum version
            if (ruleMin && currentMin && semver.gt(ruleMin, currentMin)) {
                return true;
            }
            return false;
        }
        // Ranges don't overlap - check if rule range is greater
        // Get the maximum version that satisfies the current range
        const currentMax = __classPrivateFieldGet(this, _DependencyRule_instances, "m", _DependencyRule_getMaxVersion).call(this, currentVersion);
        // If rule's minimum is greater than current's maximum, update is needed
        return !!(ruleMin && currentMax && semver.gt(ruleMin, currentMax));
    }
    catch {
        return false;
    }
}, _DependencyRule_getMaxVersion = function _DependencyRule_getMaxVersion(range) {
    const rangeObj = new semver.Range(range);
    // If it's a specific version (no range operators), return it
    if (semver.valid(range)) {
        return semver.parse(range);
    }
    // For ranges, get the highest version from the set
    // Check the range set to find upper bounds
    let maxVer = null;
    for (const comparatorSet of rangeObj.set) {
        for (const comparator of comparatorSet) {
            if (comparator.operator === '<' || comparator.operator === '<=') {
                const ver = comparator.semver;
                if (!maxVer || semver.gt(ver, maxVer)) {
                    maxVer = ver;
                }
            }
        }
    }
    // If no upper bound found, use minVersion as fallback
    return maxVer || semver.minVersion(range);
};
//# sourceMappingURL=DependencyRule.js.map