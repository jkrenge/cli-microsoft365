var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaConnectorListCommand_instances, _PaConnectorListCommand_initOptions;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import flowCommands from '../../../flow/commands.js';
import commands from '../../commands.js';
class PaConnectorListCommand extends PowerAppsCommand {
    get name() {
        return commands.CONNECTOR_LIST;
    }
    get description() {
        return 'Lists custom connectors in the given environment';
    }
    alias() {
        return [flowCommands.CONNECTOR_LIST];
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    constructor() {
        super();
        _PaConnectorListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaConnectorListCommand_instances, "m", _PaConnectorListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const url = `${this.resource}/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27${formatting.encodeQueryParameter(args.options.environmentName)}%27%20and%20IsCustomApi%20eq%20%27True%27`;
        try {
            const connectors = await odata.getAllItems(url);
            if (connectors.length > 0) {
                connectors.forEach(c => {
                    c.displayName = c.properties.displayName;
                });
            }
            await logger.log(connectors);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PaConnectorListCommand_instances = new WeakSet(), _PaConnectorListCommand_initOptions = function _PaConnectorListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    });
};
export default new PaConnectorListCommand();
//# sourceMappingURL=connector-list.js.map