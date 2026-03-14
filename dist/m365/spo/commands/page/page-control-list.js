var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageControlListCommand_instances, _SpoPageControlListCommand_initOptions, _SpoPageControlListCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { getControlTypeDisplayName } from './pageMethods.js';
class SpoPageControlListCommand extends SpoCommand {
    get name() {
        return commands.PAGE_CONTROL_LIST;
    }
    get description() {
        return 'Lists controls on the specific modern page';
    }
    defaultProperties() {
        return ['id', 'type', 'title'];
    }
    constructor() {
        super();
        _SpoPageControlListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageControlListCommand_instances, "m", _SpoPageControlListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageControlListCommand_instances, "m", _SpoPageControlListCommand_initValidators).call(this);
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
            const controls = canvasData.filter(c => c.position).map(c => {
                return {
                    id: c.id,
                    type: getControlTypeDisplayName(c.controlType || 0),
                    title: c.webPartData?.title,
                    controlType: c.controlType,
                    order: c.position.sectionIndex,
                    controlData: {
                        ...c
                    }
                };
            });
            await logger.log(JSON.parse(JSON.stringify(controls)));
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageControlListCommand_instances = new WeakSet(), _SpoPageControlListCommand_initOptions = function _SpoPageControlListCommand_initOptions() {
    this.options.unshift({
        option: '-n, --pageName <pageName>'
    }, {
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoPageControlListCommand_initValidators = function _SpoPageControlListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPageControlListCommand();
//# sourceMappingURL=page-control-list.js.map