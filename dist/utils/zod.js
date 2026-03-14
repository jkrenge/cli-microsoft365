import { z } from 'zod';
z.ZodType.prototype.alias = function (name) {
    this.def.alias = name;
    return this;
};
function parseObject(schema, options, _currentOption) {
    for (const key in schema.properties) {
        const property = schema.properties[key];
        const option = {
            name: key,
            long: key,
            short: property['x-alias'],
            required: schema.required?.includes(key) && property.default === undefined || false,
            type: 'string'
        };
        parseJSONSchema(property, options, option);
        options.push(option);
    }
    return;
}
function parseString(schema, _options, currentOption) {
    if (currentOption) {
        currentOption.type = 'string';
        if (schema.enum) {
            currentOption.autocomplete = schema.enum.map(e => String(e));
        }
    }
    return;
}
function parseNumber(schema, _options, currentOption) {
    if (currentOption) {
        currentOption.type = 'number';
    }
    return;
}
function parseBoolean(schema, _options, currentOption) {
    if (currentOption) {
        currentOption.type = 'boolean';
    }
    return;
}
function getParseFn(typeName) {
    switch (typeName) {
        case 'object':
            return parseObject;
        case 'string':
            return parseString;
        case 'number':
        case 'integer':
            return parseNumber;
        case 'boolean':
            return parseBoolean;
        default:
            return;
    }
}
function parseJSONSchema(jsonSchema, options, currentOption) {
    let parsedSchema = jsonSchema;
    do {
        if (parsedSchema.allOf) {
            parsedSchema.allOf.forEach(s => parseJSONSchema(s, options, currentOption));
        }
        const parse = getParseFn(parsedSchema.type);
        if (!parse) {
            break;
        }
        parsedSchema = parse(parsedSchema, options, currentOption);
        if (!parsedSchema) {
            break;
        }
    } while (parsedSchema);
}
function optionToString(optionInfo) {
    let s = '';
    if (optionInfo.short) {
        s += `-${optionInfo.short}, `;
    }
    s += `--${optionInfo.long}`;
    if (optionInfo.type !== 'boolean') {
        s += ' ';
        s += optionInfo.required ? '<' : '[';
        s += optionInfo.long;
        s += optionInfo.required ? '>' : ']';
    }
    return s;
}
;
export const zod = {
    schemaToOptionInfo(schema) {
        const jsonSchema = z.toJSONSchema(schema, {
            override: s => {
                const alias = s.zodSchema.def.alias;
                if (alias) {
                    s.jsonSchema['x-alias'] = alias;
                }
            },
            unrepresentable: 'any'
        });
        const options = [];
        parseJSONSchema(jsonSchema, options);
        return options;
    },
    schemaToOptions(schema) {
        const optionsInfo = this.schemaToOptionInfo(schema);
        const options = optionsInfo.map(option => {
            return {
                option: optionToString(option),
                autocomplete: option.autocomplete
            };
        });
        return options;
    },
    coercedEnum: (e) => z.preprocess(val => {
        const target = String(val)?.toLowerCase();
        for (const k of Object.values(e)) {
            if (String(k)?.toLowerCase() === target) {
                return k;
            }
        }
        return null;
    }, z.enum(e))
};
//# sourceMappingURL=zod.js.map