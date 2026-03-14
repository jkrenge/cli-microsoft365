import { FileAddRemoveRule } from "./FileAddRemoveRule.js";
export class FN015004_FILE_config_tslint_json extends FileAddRemoveRule {
    constructor(add) {
        super('./config/tslint.json', add);
    }
    get id() {
        return 'FN015004';
    }
    get supersedes() {
        return ['FN008001', 'FN008002', 'FN008003'];
    }
}
//# sourceMappingURL=FN015004_FILE_config_tslint_json.js.map