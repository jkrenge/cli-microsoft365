var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageRemoveCommand_instances, _SpoPageRemoveCommand_initTelemetry, _SpoPageRemoveCommand_initOptions, _SpoPageRemoveCommand_initValidators, _SpoPageRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoPageRemoveCommand extends SpoCommand {
    get name() {
        return commands.PAGE_REMOVE;
    }
    get description() {
        return 'Removes a modern page';
    }
    constructor() {
        super();
        _SpoPageRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageRemoveCommand_instances, "m", _SpoPageRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPageRemoveCommand_instances, "m", _SpoPageRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageRemoveCommand_instances, "m", _SpoPageRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoPageRemoveCommand_instances, "m", _SpoPageRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removePage(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove page '${args.options.name}'?` });
            if (result) {
                await this.removePage(logger, args);
            }
        }
    }
    async removePage(logger, args) {
        try {
            // Remove leading slashes from the page name (page can be nested in folders)
            let pageName = urlUtil.removeLeadingSlashes(args.options.name);
            if (!pageName.toLowerCase().endsWith('.aspx')) {
                pageName += '.aspx';
            }
            if (this.verbose) {
                await logger.logToStderr(`Removing page ${pageName}...`);
            }
            const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            if (args.options.bypassSharedLock) {
                requestOptions.headers.Prefer = 'bypass-shared-lock';
            }
            if (args.options.recycle) {
                requestOptions.url += '/Recycle';
                await request.post(requestOptions);
            }
            else {
                await request.delete(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageRemoveCommand_instances = new WeakSet(), _SpoPageRemoveCommand_initTelemetry = function _SpoPageRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force,
            recycle: !!args.options.recycle,
            bypassSharedLock: !!args.options.bypassSharedLock
        });
    });
}, _SpoPageRemoveCommand_initOptions = function _SpoPageRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--recycle'
    }, {
        option: '--bypassSharedLock'
    }, {
        option: '-f, --force'
    });
}, _SpoPageRemoveCommand_initValidators = function _SpoPageRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoPageRemoveCommand_initTypes = function _SpoPageRemoveCommand_initTypes() {
    this.types.string.push('name', 'webUrl');
    this.types.boolean.push('force', 'bypassSharedLock', 'recycle');
};
export default new SpoPageRemoveCommand();
//# sourceMappingURL=page-remove.js.map