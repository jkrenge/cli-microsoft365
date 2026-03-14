var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMessageMoveCommand_instances, _OutlookMessageMoveCommand_initTelemetry, _OutlookMessageMoveCommand_initOptions, _OutlookMessageMoveCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
import { cli } from '../../../../cli/cli.js';
import DelegatedGraphCommand from '../../../base/GraphDelegatedCommand.js';
class OutlookMessageMoveCommand extends DelegatedGraphCommand {
    get name() {
        return commands.MESSAGE_MOVE;
    }
    get description() {
        return 'Moves message to the specified folder';
    }
    constructor() {
        super();
        _OutlookMessageMoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMessageMoveCommand_instances, "m", _OutlookMessageMoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMessageMoveCommand_instances, "m", _OutlookMessageMoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OutlookMessageMoveCommand_instances, "m", _OutlookMessageMoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let sourceFolder;
        let targetFolder;
        try {
            sourceFolder = await this.getFolderId(args.options.sourceFolderId, args.options.sourceFolderName);
            targetFolder = await this.getFolderId(args.options.targetFolderId, args.options.targetFolderName);
            const messageUrl = `mailFolders/${sourceFolder}/messages/${args.options.id}`;
            const requestOptions = {
                url: `${this.resource}/v1.0/me/${messageUrl}/move`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: {
                    destinationId: targetFolder
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFolderId(folderId, folderName) {
        if (folderId) {
            return folderId;
        }
        if (Outlook.wellKnownFolderNames.indexOf(folderName) > -1) {
            return folderName;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(folderName)}'&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.value.length === 0) {
            throw `Folder with name '${folderName}' not found`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple folders with name '${folderName}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return response.value[0].id;
    }
}
_OutlookMessageMoveCommand_instances = new WeakSet(), _OutlookMessageMoveCommand_initTelemetry = function _OutlookMessageMoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            sourceFolderId: typeof args.options.sourceFolderId !== 'undefined',
            sourceFolderName: typeof args.options.sourceFolderName !== 'undefined',
            targetFolderId: typeof args.options.targetFolderId !== 'undefined',
            targetFolderName: typeof args.options.targetFolderName !== 'undefined'
        });
    });
}, _OutlookMessageMoveCommand_initOptions = function _OutlookMessageMoveCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '--sourceFolderName [sourceFolderName]',
        autocomplete: Outlook.wellKnownFolderNames
    }, {
        option: '--sourceFolderId [sourceFolderId]',
        autocomplete: Outlook.wellKnownFolderNames
    }, {
        option: '--targetFolderName [targetFolderName]',
        autocomplete: Outlook.wellKnownFolderNames
    }, {
        option: '--targetFolderId [targetFolderId]',
        autocomplete: Outlook.wellKnownFolderNames
    });
}, _OutlookMessageMoveCommand_initOptionSets = function _OutlookMessageMoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['sourceFolderId', 'sourceFolderName'] }, { options: ['targetFolderId', 'targetFolderName'] });
};
export default new OutlookMessageMoveCommand();
//# sourceMappingURL=message-move.js.map