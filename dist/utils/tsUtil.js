import ts from 'typescript';
export const tsUtil = {
    // wrapper needed to avoid the
    // "Descriptor for property createSourceFile is non-configurable and non-writable"
    // error in tests
    createSourceFile: (fileName, sourceText, languageVersionOrOptions, setParentNodes, scriptKind) => ts.createSourceFile(fileName, sourceText, languageVersionOrOptions, setParentNodes, scriptKind)
};
//# sourceMappingURL=tsUtil.js.map