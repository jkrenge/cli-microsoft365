import fs from 'fs';
import request from '../../../../../../request.js';
import { BasicDependencyRule } from "./BasicDependencyRule.js";
export class DynamicRule extends BasicDependencyRule {
    constructor() {
        super(...arguments);
        this.restrictedModules = ['react', 'react-dom', '@pnp/sp-clientsvc', '@pnp/sp-taxonomy'];
        this.restrictedNamespaces = ['@types/', '@microsoft/'];
        this.fileVariationSuffixes = ['.min', '.bundle', '-min', '.bundle.min'];
    }
    async visit(project) {
        if (!project.packageJson || !project.packageJson.dependencies) {
            return { entries: [], suggestions: [] };
        }
        const validPackageNames = Object.getOwnPropertyNames(project.packageJson.dependencies)
            .filter(x => this.restrictedNamespaces.map(y => x.indexOf(y) === -1).reduce((y, z) => y && z))
            .filter(x => this.restrictedModules.indexOf(x) === -1);
        const res = await Promise.all(validPackageNames.map((x) => this.getExternalEntryForPackage(x, project)));
        return {
            entries: res.filter(x => x !== undefined),
            suggestions: []
        };
    }
    async getExternalEntryForPackage(packageName, project) {
        const version = project.packageJson.dependencies[packageName];
        const filesPaths = this.getFilePath(packageName).map(x => this.cleanFilePath(x));
        if (!version || filesPaths.length === 0) {
            return undefined;
        }
        const filesPathsVariations = filesPaths
            .map(x => this.fileVariationSuffixes.map(y => x.indexOf(y) === -1 ? x.replace('.js', `${y}.js`) : x))
            .reduce((x, y) => [...x, ...y]);
        const pathsAndVariations = [...filesPaths, ...filesPathsVariations];
        const externalizeEntryCandidates = await Promise.all(pathsAndVariations.map(x => this.getExternalEntryForFilePath(x, packageName, version)));
        const dExternalizeEntryCandidates = externalizeEntryCandidates.filter(x => x !== undefined);
        const minifiedModule = dExternalizeEntryCandidates.find(x => !x.globalName && this.pathContainsMinifySuffix(x.path));
        const minifiedNonModule = dExternalizeEntryCandidates.find(x => x.globalName && this.pathContainsMinifySuffix(x.path));
        const nonMinifiedModule = dExternalizeEntryCandidates.find(x => x.globalName && !this.pathContainsMinifySuffix(x.path));
        const nonMinifiedNonModule = dExternalizeEntryCandidates.find(x => !x.globalName && !this.pathContainsMinifySuffix(x.path));
        const externalizeEntriesPrioritized = [minifiedModule, minifiedNonModule, nonMinifiedModule, nonMinifiedNonModule].filter(x => x !== undefined);
        return externalizeEntriesPrioritized.length > 0 ? externalizeEntriesPrioritized[0] : undefined;
    }
    pathContainsMinifySuffix(path) {
        return this.fileVariationSuffixes
            .map(y => path.indexOf(y))
            .filter(y => y > -1).length > 0;
    }
    async getExternalEntryForFilePath(filePath, packageName, version) {
        const url = this.getFileUrl(packageName, version, filePath);
        const testResult = await this.testUrl(url);
        if (!testResult) {
            return undefined;
        }
        const moduleInfo = await this.getModuleType(url);
        if (moduleInfo.scriptType === 'CommonJs') {
            return undefined; //browsers don't support those module types without an additional library
        }
        else if (moduleInfo.scriptType === 'ES2015' || moduleInfo.scriptType === 'AMD') {
            return {
                key: packageName,
                path: url
            };
        }
        else { //TODO for non-module and UMD we should technically add dependencies as well
            return {
                key: packageName,
                path: url,
                globalName: moduleInfo.exports && moduleInfo.exports.length > 0 ? moduleInfo.exports[0] : packageName // examples where this is not good https://unpkg.com/@pnp/polyfill-ie11@^1.0.2/dist/index.js https://unpkg.com/moment-timezone@^0.5.27/builds/moment-timezone-with-data.js
            };
        }
    }
    getModuleType(url) {
        const requestOptions = {
            url: 'https://scriptcheck-weu-fn.azurewebsites.net/api/v2/script-check',
            headers: { 'content-type': 'application/json', accept: 'application/json', 'x-anonymous': 'true' },
            data: { url: url },
            responseType: 'json'
        };
        return request
            .post(requestOptions)
            .catch(() => {
            return { scriptType: 'non-module', exports: [] };
        });
    }
    getFileUrl(packageName, version, filePath) {
        return `https://unpkg.com/${packageName}@${version}/${filePath}`;
    }
    async testUrl(url) {
        try {
            await request.head({ url: url, headers: { 'x-anonymous': 'true' } });
            return true;
        }
        catch {
            return false;
        }
    }
    getFilePath(packageName) {
        const packageJsonFilePath = `node_modules/${packageName}/package.json`;
        let filesPaths = [];
        try {
            const packageJson = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));
            if (packageJson.module) {
                filesPaths.push(packageJson.module);
            }
            if (packageJson.main) {
                filesPaths.push(packageJson.main);
            }
            if (packageJson.es2015) {
                filesPaths.push(packageJson.es2015);
            }
            if (packageJson.jspm) {
                if (packageJson.jspm.main) {
                    filesPaths.push(packageJson.jspm.main);
                }
                if (packageJson.jspm.files) {
                    filesPaths.push(...packageJson.jspm.files);
                }
            }
            if (packageJson.spm) {
                if (packageJson.spm.main) {
                    filesPaths.push(packageJson.spm.main);
                }
            }
            filesPaths = filesPaths
                .filter((x, idx) => filesPaths.indexOf(x) === idx) //filtering duplicates
                .map(x => x.endsWith('.js') ? x : `${x}.js`); //some package managers omit the file extension
        }
        catch {
            // file doesn't exist, giving up
        }
        return filesPaths;
    }
    cleanFilePath(filePath) {
        return filePath.replace('./', '');
    }
}
//# sourceMappingURL=DynamicRule.js.map