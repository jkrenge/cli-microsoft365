import { FileAddRemoveRule } from "./FileAddRemoveRule.js";
export class FN015011_FILE_tsconfig_json extends FileAddRemoveRule {
    constructor(add, contents) {
        super('./tsconfig.json', add, contents);
    }
    get id() {
        return 'FN015011';
    }
}
//# sourceMappingURL=FN015011_FILE_tsconfig_json.js.map