var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoNavigationNodeAddCommand_instances, _SpoNavigationNodeAddCommand_initTelemetry, _SpoNavigationNodeAddCommand_initOptions, _SpoNavigationNodeAddCommand_initValidators, _SpoNavigationNodeAddCommand_initOptionSets;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoNavigationNodeAddCommand extends SpoCommand {
    get name() {
        return commands.NAVIGATION_NODE_ADD;
    }
    get description() {
        return 'Adds a navigation node to the specified site navigation';
    }
    constructor() {
        super();
        _SpoNavigationNodeAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeAddCommand_instances, "m", _SpoNavigationNodeAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeAddCommand_instances, "m", _SpoNavigationNodeAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeAddCommand_instances, "m", _SpoNavigationNodeAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeAddCommand_instances, "m", _SpoNavigationNodeAddCommand_initOptionSets).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding navigation node...`);
        }
        const nodesCollection = args.options.parentNodeId ?
            `GetNodeById(${args.options.parentNodeId})/Children` :
            args.options.location.toLowerCase();
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/navigation/${nodesCollection}`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                AudienceIds: args.options.audienceIds?.split(','),
                Title: args.options.title,
                Url: args.options.url ?? 'http://linkless.header/',
                IsExternal: args.options.isExternal === true
            }
        };
        try {
            const res = await request.post(requestOptions);
            if (args.options.openInNewWindow) {
                if (this.verbose) {
                    await logger.logToStderr(`Making sure that the newly added navigation node opens in a new window.`);
                }
                const id = res.Id.toString();
                let menuState = args.options.location === 'TopNavigationBar' ? await spo.getTopNavigationMenuState(args.options.webUrl) : await spo.getQuickLaunchMenuState(args.options.webUrl);
                let menuStateItem = this.getMenuStateNode(menuState.Nodes, id);
                if (args.options.parentNodeId && !menuStateItem) {
                    menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
                    menuStateItem = this.getMenuStateNode(menuState.Nodes, id);
                }
                menuStateItem.OpenInNewWindow = true;
                await spo.saveMenuState(args.options.webUrl, menuState);
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getMenuStateNode(nodes, id) {
        let menuNode = nodes.find((node) => node.Key !== null && node.Key === id);
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
_SpoNavigationNodeAddCommand_instances = new WeakSet(), _SpoNavigationNodeAddCommand_initTelemetry = function _SpoNavigationNodeAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            isExternal: args.options.isExternal,
            location: typeof args.options.location !== 'undefined',
            parentNodeId: typeof args.options.parentNodeId !== 'undefined',
            audienceIds: typeof args.options.audienceIds !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            openInNewWindow: !!args.options.openInNewWindow
        });
    });
}, _SpoNavigationNodeAddCommand_initOptions = function _SpoNavigationNodeAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --location [location]',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
    }, {
        option: '-t, --title <title>'
    }, {
        option: '--url [url]'
    }, {
        option: '--parentNodeId [parentNodeId]'
    }, {
        option: '--isExternal'
    }, {
        option: '--audienceIds [audienceIds]'
    }, {
        option: '--openInNewWindow'
    });
}, _SpoNavigationNodeAddCommand_initValidators = function _SpoNavigationNodeAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.parentNodeId) {
            if (isNaN(args.options.parentNodeId)) {
                return `${args.options.parentNodeId} is not a number`;
            }
        }
        else {
            if (args.options.location !== 'QuickLaunch' &&
                args.options.location !== 'TopNavigationBar') {
                return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
            }
        }
        if (args.options.audienceIds) {
            const audienceIdsSplitted = args.options.audienceIds.split(',');
            if (audienceIdsSplitted.length > 10) {
                return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of auciendeIds is 10.';
            }
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.audienceIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'audienceIds': ${isValidGUIDArrayResult}.`;
            }
        }
        return true;
    });
}, _SpoNavigationNodeAddCommand_initOptionSets = function _SpoNavigationNodeAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['location', 'parentNodeId'] });
};
export default new SpoNavigationNodeAddCommand();
//# sourceMappingURL=navigation-node-add.js.map