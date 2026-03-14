var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemBatchSetCommand_instances, _SpoListItemBatchSetCommand_initTelemetry, _SpoListItemBatchSetCommand_initOptions, _SpoListItemBatchSetCommand_initValidators, _SpoListItemBatchSetCommand_initOptionSets;
import fs from 'fs';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemBatchSetCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_BATCH_SET;
    }
    get description() {
        return 'Updates list items in a batch';
    }
    constructor() {
        super();
        _SpoListItemBatchSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemBatchSetCommand_instances, "m", _SpoListItemBatchSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchSetCommand_instances, "m", _SpoListItemBatchSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchSetCommand_instances, "m", _SpoListItemBatchSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchSetCommand_instances, "m", _SpoListItemBatchSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Starting to create batch items from csv at path ${args.options.filePath}`);
            }
            const csvContent = fs.readFileSync(args.options.filePath, 'utf8');
            const jsonContent = formatting.parseCsvToJson(csvContent);
            const amountOfRows = jsonContent.length;
            const idColumn = args.options.idColumn || 'ID';
            if (!Object.prototype.hasOwnProperty.call(jsonContent[0], idColumn)) {
                throw `The specified value for idColumn does not exist in the array. Specified idColumn is '${args.options.idColumn || 'ID'}'. Please specify the correct value.`;
            }
            const listId = args.options.listId ?
                args.options.listId :
                await spo.getListId(args.options.webUrl, args.options.listTitle, args.options.listUrl, logger, this.verbose);
            const fields = await this.getListFields(args.options, listId, jsonContent, idColumn, logger);
            const userFields = fields.filter(field => field.TypeAsString === 'UserMulti' || field.TypeAsString === 'User');
            const resolvedUsers = await this.getUsersFromCsv(args.options.webUrl, jsonContent, userFields);
            const formDigestValue = (await spo.getRequestDigest(args.options.webUrl)).FormDigestValue;
            const objectIdentity = (await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue)).objectIdentity;
            let counter = 0;
            while (jsonContent.length > 0) {
                const entriesToProcess = jsonContent.splice(0, 50);
                const objectPaths = [], actions = [];
                let index = 1;
                for (const row of entriesToProcess) {
                    counter += 1;
                    objectPaths.push(`<Identity Id="${index}" Name="${objectIdentity}:list:${listId}:item:${row[idColumn]},1" />`);
                    const [actionString, updatedIndex] = this.mapActions(index, row, fields, resolvedUsers, args.options.systemUpdate);
                    index = updatedIndex;
                    actions.push(actionString);
                }
                if (this.verbose) {
                    await logger.logToStderr(`Writing away batch of items, currently at: ${counter}/${amountOfRows}.`);
                }
                await this.sendBatchRequest(args.options.webUrl, this.getRequestBody(objectPaths, actions));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getRequestBody(objectPaths, actions) {
        return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${actions.join('')}</Actions><ObjectPaths>${objectPaths.join('')}</ObjectPaths></Request>`;
    }
    async sendBatchRequest(webUrl, requestBody) {
        const requestOptions = {
            url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'Content-Type': 'text/xml'
            },
            data: requestBody
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage + ' - ' + response.ErrorInfo.ErrorValue;
        }
    }
    mapActions(index, row, fields, users, systemUpdate) {
        const objectPathId = index;
        let actionString = '';
        fields.forEach((field) => {
            if (row[field.InternalName] === undefined || row[field.InternalName] === '') {
                actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String"></Parameter></Parameters></Method>`;
            }
            else {
                switch (field.TypeAsString) {
                    case 'User': {
                        const userDetail = users.find(us => us.email === row[field.InternalName]);
                        actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String">${userDetail.id}</Parameter></Parameters></Method>`;
                        break;
                    }
                    case 'UserMulti': {
                        const userMultiString = row[field.InternalName].toString().split(';').map((element) => {
                            const userDetail = users.find(us => us.email === element);
                            return `<Object TypeId="{c956ab54-16bd-4c18-89d2-996f57282a6f}"><Property Name="Email" Type="Null" /><Property Name="LookupId" Type="Int32">${userDetail.id}</Property><Property Name="LookupValue" Type="Null" /></Object>`;
                        });
                        actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="Array">${userMultiString.join('')}</Parameter></Parameters></Method>`;
                        break;
                    }
                    case 'Lookup': {
                        actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter TypeId="{f1d34cc0-9b50-4a78-be78-d5facfcccfb7}"><Property Name="LookupId" Type="Int32">${row[field.InternalName]}</Property><Property Name="LookupValue" Type="Null"/></Parameter></Parameters></Method>`;
                        break;
                    }
                    case 'LookupMulti': {
                        const lookupMultiString = row[field.InternalName].toString().split(';').map((element) => {
                            return `<Object TypeId="{f1d34cc0-9b50-4a78-be78-d5facfcccfb7}"><Property Name="LookupId" Type="Int32">${element}</Property><Property Name="LookupValue" Type="Null" /></Object>`;
                        });
                        actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="Array">${lookupMultiString.join('')}</Parameter></Parameters></Method>`;
                        break;
                    }
                    default:
                        actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String">${formatting.escapeXml(row[field.InternalName].toString())}</Parameter></Parameters></Method>`;
                        break;
                }
            }
        });
        actionString += `<Method Name="${systemUpdate ? 'System' : ''}Update" Id="${index += 1}" ObjectPathId="${objectPathId}"/>`;
        return [actionString, index];
    }
    async getListFields(options, listId, jsonContent, idColumn, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving fields for list with id ${listId}`);
        }
        const filterFields = [];
        const objectKeys = Object.keys(jsonContent[0]);
        const index = objectKeys.indexOf(idColumn, 0);
        if (index > -1) {
            objectKeys.splice(index, 1);
        }
        objectKeys.map(objectKey => {
            filterFields.push(`InternalName eq '${objectKey}'`);
        });
        const fields = await odata.getAllItems(`${options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`);
        if (fields.length !== objectKeys.length) {
            const fieldsThatDontExist = [];
            objectKeys.forEach(key => {
                const field = fields.find(field => field.InternalName === key);
                if (!field) {
                    fieldsThatDontExist.push(key);
                }
            });
            throw `Following fields specified in the csv do not exist on the list: ${fieldsThatDontExist.join(', ')}`;
        }
        return fields;
    }
    async getUsersFromCsv(webUrl, jsonContent, userFields) {
        if (userFields.length === 0) {
            return [];
        }
        const userFieldValues = [];
        const emailsToResolve = this.getEmailsToEnsure(jsonContent, userFields);
        for await (const email of emailsToResolve) {
            const requestOptions = {
                url: `${webUrl}/_api/web/ensureUser('${email}')?$select=Id`,
                headers: {
                    accept: 'application/json',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            userFieldValues.push({ email: email, id: response.Id });
        }
        return userFieldValues;
    }
    getEmailsToEnsure(jsonContent, userFields) {
        const emailsToResolve = [];
        userFields.forEach((userField) => {
            jsonContent.forEach(row => {
                const fieldValue = row[userField.InternalName];
                if (fieldValue !== undefined && fieldValue !== '') {
                    const emailsSplitted = fieldValue.split(';');
                    emailsSplitted.forEach((email) => {
                        if (!emailsToResolve.some(existingMail => existingMail === email)) {
                            emailsToResolve.push(email);
                        }
                    });
                }
            });
        });
        return emailsToResolve;
    }
}
_SpoListItemBatchSetCommand_instances = new WeakSet(), _SpoListItemBatchSetCommand_initTelemetry = function _SpoListItemBatchSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            idColumn: typeof args.options.idColumn !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            systemUpdate: !!args.options.systemUpdate
        });
    });
}, _SpoListItemBatchSetCommand_initOptions = function _SpoListItemBatchSetCommand_initOptions() {
    this.options.unshift({
        option: '-p, --filePath <filePath>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--idColumn [idColumn]'
    }, {
        option: '-s, --systemUpdate'
    });
}, _SpoListItemBatchSetCommand_initValidators = function _SpoListItemBatchSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (!fs.existsSync(args.options.filePath)) {
            return `File with path ${args.options.filePath} does not exist`;
        }
        return true;
    });
}, _SpoListItemBatchSetCommand_initOptionSets = function _SpoListItemBatchSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemBatchSetCommand();
//# sourceMappingURL=listitem-batch-set.js.map