export const basic = {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    asyncFilter: async (arr, predicate, thisArg) => {
        const results = await Promise.all(arr.map(predicate));
        return arr.filter((_v, index) => results[index]);
    }
};
//# sourceMappingURL=basic.js.map