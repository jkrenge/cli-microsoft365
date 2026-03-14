export const misc = {
    getEnums(en) {
        return Object
            .keys(en)
            .filter(k => isNaN(parseInt(k)));
    }
};
//# sourceMappingURL=misc.js.map