var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewThreatAssessmentAddCommand_instances, _PurviewThreatAssessmentAddCommand_initTelemetry, _PurviewThreatAssessmentAddCommand_initOptions, _PurviewThreatAssessmentAddCommand_initValidators, _PurviewThreatAssessmentAddCommand_initOptionSets;
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import auth from '../../../../Auth.js';
import fs from 'fs';
import path from 'path';
class PurviewThreatAssessmentAddCommand extends GraphCommand {
    get name() {
        return commands.THREATASSESSMENT_ADD;
    }
    get description() {
        return 'Create a threat assessment';
    }
    constructor() {
        super();
        _PurviewThreatAssessmentAddCommand_instances.add(this);
        this.allowedTypes = ['file', 'url'];
        this.allowedExpectedAssessments = ['block', 'unblock'];
        this.allowedCategories = ['spam', 'phishing', 'malware'];
        __classPrivateFieldGet(this, _PurviewThreatAssessmentAddCommand_instances, "m", _PurviewThreatAssessmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentAddCommand_instances, "m", _PurviewThreatAssessmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentAddCommand_instances, "m", _PurviewThreatAssessmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentAddCommand_instances, "m", _PurviewThreatAssessmentAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)) {
                throw 'This command currently does not support app only permissions.';
            }
            if (this.verbose) {
                await logger.logToStderr(`Adding threat assessment of type ${args.options.type} with expected assessment ${args.options.expectedAssessment} and category ${args.options.category}`);
            }
            const requestBody = {
                expectedAssessment: args.options.expectedAssessment,
                category: args.options.category,
                url: args.options.url,
                contentData: args.options.path && fs.readFileSync(args.options.path).toString('base64'),
                fileName: args.options.path && path.basename(args.options.path)
            };
            switch (args.options.type) {
                case 'file':
                    requestBody['@odata.type'] = '#microsoft.graph.fileAssessmentRequest';
                    break;
                case 'url':
                    requestBody['@odata.type'] = '#microsoft.graph.urlAssessmentRequest';
                    break;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/informationProtection/threatAssessmentRequests`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: requestBody,
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataPromise(err);
        }
    }
}
_PurviewThreatAssessmentAddCommand_instances = new WeakSet(), _PurviewThreatAssessmentAddCommand_initTelemetry = function _PurviewThreatAssessmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            path: typeof args.options.path !== 'undefined',
            url: typeof args.options.url !== 'undefined'
        });
    });
}, _PurviewThreatAssessmentAddCommand_initOptions = function _PurviewThreatAssessmentAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type <type>',
        autocomplete: this.allowedTypes
    }, {
        option: '-e, --expectedAssessment <expectedAssessment>',
        autocomplete: this.allowedExpectedAssessments
    }, {
        option: '-c, --category <category>',
        autocomplete: this.allowedCategories
    }, {
        option: '-p, --path [path]'
    }, {
        option: '-u, --url [url]'
    });
}, _PurviewThreatAssessmentAddCommand_initValidators = function _PurviewThreatAssessmentAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!this.allowedTypes.some(type => type === args.options.type)) {
            return `${args.options.type} is not an allowed type. Allowed types are ${this.allowedTypes.join('|')}`;
        }
        if (!this.allowedExpectedAssessments.some(expectedAssessment => expectedAssessment === args.options.expectedAssessment)) {
            return `${args.options.expectedAssessment} is not an allowed expected assessment. Allowed expected assessments are ${this.allowedExpectedAssessments.join('|')}`;
        }
        if (!this.allowedCategories.some(category => category === args.options.category)) {
            return `${args.options.category} is not an allowed category. Allowed categories are ${this.allowedCategories.join('|')}`;
        }
        if (args.options.path && !fs.existsSync(args.options.path)) {
            return `File '${args.options.path}' not found. Please provide a valid path to the file.`;
        }
        return true;
    });
}, _PurviewThreatAssessmentAddCommand_initOptionSets = function _PurviewThreatAssessmentAddCommand_initOptionSets() {
    this.optionSets.push({
        options: ['path'],
        runsWhen: (args) => {
            return args.options.type === 'file';
        }
    }, {
        options: ['url'],
        runsWhen: (args) => {
            return args.options.type === 'url';
        }
    });
};
export default new PurviewThreatAssessmentAddCommand();
//# sourceMappingURL=threatassessment-add.js.map