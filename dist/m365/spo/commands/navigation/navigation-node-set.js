var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoNavigationNodeSetCommand_instances, _SpoNavigationNodeSetCommand_initTelemetry, _SpoNavigationNodeSetCommand_initTypes, _SpoNavigationNodeSetCommand_initOptions, _SpoNavigationNodeSetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoNavigationNodeSetCommand extends SpoCommand {
    get name() {
        return commands.NAVIGATION_NODE_SET;
    }
    get description() {
        return 'Adds a navigation node to the specified site navigation';
    }
    constructor() {
        super();
        _SpoNavigationNodeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeSetCommand_instances, "m", _SpoNavigationNodeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeSetCommand_instances, "m", _SpoNavigationNodeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeSetCommand_instances, "m", _SpoNavigationNodeSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeSetCommand_instances, "m", _SpoNavigationNodeSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Setting navigation node...`);
            }
            let url = args.options.url;
            if (url === '') {
                url = 'http://linkless.header/';
            }
            const requestBody = {
                Title: args.options.title,
                IsExternal: args.options.isExternal,
                Url: url
            };
            if (args.options.audienceIds !== undefined) {
                requestBody.AudienceIds = args.options.audienceIds === '' ? [] : args.options.audienceIds.split(',');
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const response = await request.patch(requestOptions);
            if (response['odata.null'] === true) {
                throw `Navigation node does not exist.`;
            }
            if (args.options.openInNewWindow !== undefined) {
                if (this.verbose) {
                    await logger.logToStderr(`Making sure that the navigation node opens in a new window.`);
                }
                let menuState = await spo.getQuickLaunchMenuState(args.options.webUrl);
                let menuStateItem = this.getMenuStateNode(menuState.Nodes, args.options.id);
                if (!menuStateItem) {
                    menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
                    menuStateItem = this.getMenuStateNode(menuState.Nodes, args.options.id);
                }
                menuStateItem.OpenInNewWindow = args.options.openInNewWindow;
                await spo.saveMenuState(args.options.webUrl, menuState);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getMenuStateNode(nodes, id) {
        let menuNode = nodes.find((node) => node.Key.toString() === id.toString());
        if (menuNode === undefined) {
            for (const node of nodes.filter(node => node.Nodes.length > 0)) {
                menuNode = this.getMenuStateNode(node.Nodes, id);
                if (menuNode) {
                    break;
                }
            }
        }
        return menuNode;
    }
}
_SpoNavigationNodeSetCommand_instances = new WeakSet(), _SpoNavigationNodeSetCommand_initTelemetry = function _SpoNavigationNodeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            isExternal: typeof args.options.isExternal !== 'undefined',
            audienceIds: typeof args.options.audienceIds !== 'undefined',
            openInNewWindow: typeof args.options.openInNewWindow !== 'undefined'
        });
    });
}, _SpoNavigationNodeSetCommand_initTypes = function _SpoNavigationNodeSetCommand_initTypes() {
    this.types.boolean.push('isExternal');
}, _SpoNavigationNodeSetCommand_initOptions = function _SpoNavigationNodeSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--id <id>'
    }, {
        option: '--title [title]'
    }, {
        option: '--url [url]'
    }, {
        option: '--audienceIds [audienceIds]'
    }, {
        option: '--isExternal [isExternal]'
    }, {
        option: '--openInNewWindow [openInNewWindow]'
    });
}, _SpoNavigationNodeSetCommand_initValidators = function _SpoNavigationNodeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.audienceIds === undefined && args.options.url === undefined && args.options.isExternal === undefined && !args.options.title && args.options.openInNewWindow === undefined) {
            return `Please specify at least one property to update.`;
        }
        if (args.options.audienceIds) {
            const audienceIdsSplitted = args.options.audienceIds.split(',');
            if (audienceIdsSplitted.length > 10) {
                return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.';
            }
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.audienceIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'audienceIds': ${isValidGUIDArrayResult}.`;
            }
        }
        return true;
    });
};
export default new SpoNavigationNodeSetCommand();
//# sourceMappingURL=navigation-node-set.js.map