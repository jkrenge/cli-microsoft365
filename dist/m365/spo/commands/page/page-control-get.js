var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageControlGetCommand_instances, _SpoPageControlGetCommand_initOptions, _SpoPageControlGetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { getControlTypeDisplayName } from './pageMethods.js';
class SpoPageControlGetCommand extends SpoCommand {
    get name() {
        return commands.PAGE_CONTROL_GET;
    }
    get description() {
        return 'Gets information about the specific control on a modern page';
    }
    constructor() {
        super();
        _SpoPageControlGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageControlGetCommand_instances, "m", _SpoPageControlGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageControlGetCommand_instances, "m", _SpoPageControlGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let pageName = args.options.pageName;
        if (args.options.pageName.indexOf('.aspx') < 0) {
            pageName += '.aspx';
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const clientSidePage = await request.get(requestOptions);
            const canvasData = clientSidePage.CanvasContent1 ? JSON.parse(clientSidePage.CanvasContent1) : [];
            const control = canvasData.find(c => c.id?.toLowerCase() === args.options.id.toLowerCase());
            if (control) {
                const controlData = {
                    id: control.id,
                    type: getControlTypeDisplayName(control.controlType || 0),
                    title: control.webPartData?.title,
                    controlType: control.controlType,
                    order: control.position.sectionIndex,
                    controlData: {
                        ...control
                    }
                };
                await logger.log(controlData);
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Control with ID ${args.options.id} not found on page ${args.options.pageName}`);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageControlGetCommand_instances = new WeakSet(), _SpoPageControlGetCommand_initOptions = function _SpoPageControlGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-n, --pageName <pageName>'
    }, {
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoPageControlGetCommand_initValidators = function _SpoPageControlGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoPageControlGetCommand();
//# sourceMappingURL=page-control-get.js.map