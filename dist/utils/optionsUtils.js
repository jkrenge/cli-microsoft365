const longOptionRegex = /--([^\s]+)/;
const shortOptionRegex = /-([a-z])\b/;
export const optionsUtils = {
    getUnknownOptions(options, knownOptions) {
        const unknownOptions = JSON.parse(JSON.stringify(options));
        // remove minimist catch-all option
        delete unknownOptions._;
        knownOptions.forEach(o => {
            const longOptionName = longOptionRegex.exec(o.option)[1];
            delete unknownOptions[longOptionName];
            // short names are optional so we need to check if the current command has
            // one before continuing
            const shortOptionMatch = shortOptionRegex.exec(o.option);
            if (shortOptionMatch) {
                const shortOptionName = shortOptionMatch[1];
                delete unknownOptions[shortOptionName];
            }
        });
        return unknownOptions;
    },
    addUnknownOptionsToPayload(payload, unknownOptions) {
        const unknownOptionsNames = Object.getOwnPropertyNames(unknownOptions);
        unknownOptionsNames.forEach(o => {
            payload[o] = unknownOptions[o];
        });
    }
};
//# sourceMappingURL=optionsUtils.js.map