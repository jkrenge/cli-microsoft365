#!/usr/bin/env node
import fs from 'fs';
import omelette from 'omelette';
import os from 'os';
import path from 'path';
import url from 'url';
import { cli } from './cli/cli.js';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
class Autocomplete {
    constructor() {
        this.commands = {};
        this.init();
    }
    init() {
        if (fs.existsSync(Autocomplete.autocompleteFilePath)) {
            try {
                const data = fs.readFileSync(Autocomplete.autocompleteFilePath, 'utf-8');
                this.commands = JSON.parse(data);
            }
            catch {
                // Do nothing
            }
        }
        this.omelette = omelette('m365_comp|m365|microsoft365');
        this.omelette.on('complete', this.handleAutocomplete.bind(this));
        this.omelette.init();
    }
    handleAutocomplete(fragment, data) {
        let replies = {};
        let allWords = [];
        if (data.fragment === 1) {
            replies = Object.keys(this.commands);
        }
        else {
            allWords = data.line.split(/\s+/).slice(1, -1);
            // build array of words to use as a path to retrieve completion
            // options from the commands tree
            const words = allWords
                .filter((e, i) => {
                if (e.indexOf('-') !== 0) {
                    // if the word is not an option check if it's not
                    // option's value, eg. --output json, in which case
                    // the suggestion should be command options
                    return i === 0 || allWords[i - 1].indexOf('-') !== 0;
                }
                else {
                    // remove all options but last one
                    return i === allWords.length - 1;
                }
            });
            let accessor = new Function('_', "return _['" + (words.join("']['")) + "']");
            try {
                replies = accessor(this.commands);
                // if the last word is an option without autocomplete
                // suggest other options from the same command
                if (words[words.length - 1].indexOf('-') === 0 &&
                    !Array.isArray(replies)) {
                    accessor = new Function('_', "return _['" + (words.filter(w => w.indexOf('-') !== 0).join("']['")) + "']");
                    replies = accessor(this.commands);
                    replies = Object.keys(replies);
                }
            }
            catch {
                // Do nothing
            }
        }
        if (!replies) {
            replies = [];
        }
        if (!Array.isArray(replies)) {
            replies = Object.keys(replies);
        }
        // remove options that already have been used
        replies = replies.filter(r => r.indexOf('-') !== 0 || allWords.indexOf(r) === -1);
        data.reply(replies);
    }
    generateShCompletion() {
        const commandsInfo = this.getCommandsInfo();
        fs.writeFileSync(Autocomplete.autocompleteFilePath, JSON.stringify(commandsInfo));
    }
    setupShCompletion() {
        this.omelette.setupShellInitFile();
    }
    getClinkCompletion() {
        const cmd = this.getCommandsInfo();
        const lua = ['local parser = clink.arg.new_parser'];
        const functions = {};
        this.buildClinkForBranch(cmd, functions, 'm365');
        Object.keys(functions).forEach(k => {
            functions[k] = functions[k].replace(/#([^#]+)#/g, (m, p1) => functions[p1]);
        });
        lua.push('local m365_parser = ' + functions['m365'], '', 'clink.arg.register_parser("m365", m365_parser)', 'clink.arg.register_parser("microsoft365", m365_parser)');
        return lua.join(os.EOL);
    }
    buildClinkForBranch(branch, functions, luaFunctionName) {
        if (!Array.isArray(branch)) {
            const keys = Object.keys(branch);
            keys.forEach(k => {
                if (Object.keys(branch[k]).length > 0) {
                    this.buildClinkForBranch(branch[k], functions, this.getLuaFunctionName(`${luaFunctionName}_${k}`));
                }
            });
        }
        const parser = [];
        parser.push(`parser({`);
        let printingArgs = false;
        if (Array.isArray(branch)) {
            branch.sort().forEach((c, i) => {
                const separator = i < branch.length - 1 ? ',' : '';
                parser.push(`"${c}"${separator}`);
            });
        }
        else {
            const keys = Object.keys(branch);
            if (keys.find(c => c.indexOf('-') === 0)) {
                printingArgs = true;
                const tmp = [];
                keys.sort().forEach((k) => {
                    if (Object.keys(branch[k]).length > 0) {
                        tmp.push(`"${k}"..#${this.getLuaFunctionName(`${luaFunctionName}_${k}`)}#`);
                    }
                    else {
                        tmp.push(`"${k}"`);
                    }
                });
                parser.push(`},${tmp.join(', ')}`);
            }
            else {
                keys.sort().forEach((k, i) => {
                    const separator = i < keys.length - 1 ? ',' : '';
                    parser.push(`"${k}"..#${this.getLuaFunctionName(`${luaFunctionName}_${k}`)}#${separator}`);
                });
            }
        }
        parser.push(`${printingArgs ? '' : '}'})`);
        functions[luaFunctionName] = parser.join('');
    }
    getLuaFunctionName(functionName) {
        return functionName.replace(/-/g, '_');
    }
    getCommandsInfo() {
        const commandsInfo = {};
        const commands = cli.commands;
        commands.forEach(c => {
            Autocomplete.processCommand(c.name, c, commandsInfo);
            if (c.aliases) {
                c.aliases.forEach(a => Autocomplete.processCommand(a, c, commandsInfo));
            }
        });
        return commandsInfo;
    }
    static processCommand(commandName, commandInfo, autocomplete) {
        const chunks = commandName.split(' ');
        let parent = autocomplete;
        for (let i = 0; i < chunks.length; i++) {
            const current = chunks[i];
            if (!parent[current]) {
                if (i < chunks.length - 1) {
                    parent[current] = {};
                }
                else {
                    // last chunk, add options
                    const optionsArr = commandInfo.options
                        .map(o => o.short)
                        .concat(commandInfo.options.map(o => o.long))
                        .filter(o => o)
                        .map(o => o.length === 1 ? `-${o}` : `--${o}`);
                    optionsArr.push('--help');
                    optionsArr.push('-h');
                    const optionsObj = {};
                    optionsArr.forEach(o => {
                        const optionName = o.replace(/^-+/, '');
                        const option = commandInfo.options.filter(opt => opt.long === optionName || opt.short === optionName)[0];
                        if (option && option.autocomplete) {
                            optionsObj[o] = option.autocomplete;
                        }
                        else {
                            optionsObj[o] = {};
                        }
                    });
                    parent[current] = optionsObj;
                }
            }
            parent = parent[current];
        }
    }
}
Autocomplete.autocompleteFilePath = path.join(__dirname, `..${path.sep}commands.json`);
export const autocomplete = new Autocomplete();
//# sourceMappingURL=autocomplete.js.map