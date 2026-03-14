import { DependencyRule } from "./DependencyRule.js";
export class FN002003_DEVDEP_microsoft_sp_webpart_workbench extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@microsoft/sp-webpart-workbench', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002003';
    }
}
//# sourceMappingURL=FN002003_DEVDEP_microsoft_sp_webpart_workbench.js.map