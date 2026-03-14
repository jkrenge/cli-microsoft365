var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppPublishCommand_instances, _TeamsAppPublishCommand_initOptions, _TeamsAppPublishCommand_initValidators;
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsAppPublishCommand extends GraphCommand {
    get name() {
        return commands.APP_PUBLISH;
    }
    get description() {
        return 'Publishes Teams app to the organization\'s app catalog';
    }
    constructor() {
        super();
        _TeamsAppPublishCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppPublishCommand_instances, "m", _TeamsAppPublishCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppPublishCommand_instances, "m", _TeamsAppPublishCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const fullPath = path.resolve(args.options.filePath);
            if (this.verbose) {
                await logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/appCatalogs/teamsApps`,
                headers: {
                    'content-type': 'application/zip',
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: fs.readFileSync(fullPath)
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsAppPublishCommand_instances = new WeakSet(), _TeamsAppPublishCommand_initOptions = function _TeamsAppPublishCommand_initOptions() {
    this.options.unshift({
        option: '-p, --filePath <filePath>'
    });
}, _TeamsAppPublishCommand_initValidators = function _TeamsAppPublishCommand_initValidators() {
    this.validators.push(async (args) => {
        const fullPath = path.resolve(args.options.filePath);
        if (!fs.existsSync(fullPath)) {
            return `File '${fullPath}' not found`;
        }
        if (fs.lstatSync(fullPath).isDirectory()) {
            return `Path '${fullPath}' points to a directory`;
        }
        return true;
    });
};
export default new TeamsAppPublishCommand();
//# sourceMappingURL=app-publish.js.map