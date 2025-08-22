import globals from "globals";
import pluginJs from "@eslint/js";


/** @type {import('eslint').Linter.Config[]} */
export default
[
    {
        files: ["**/*.js",],
        languageOptions: {sourceType: "commonjs",},
    },
    {
        languageOptions:
        {
            globals:
            {
                console: "readonly",
                document: "readonly",
                expect: "readonly",
                it: "readonly",
                jest: "readonly",
                Office: "readonly",
                test: "readonly",
                Word: "readonly",
                Excel: "readonly",
                window: "readonly",
                initCustomProp: "readonly",
                fetch: "readonly",
                addCustomProperty: "readonly",
                readCustomProperty: "readonly",
            },
        },
    },
    pluginJs.configs.recommended,
    {
        rules:
        {
            'key-spacing': ["error", { "beforeColon": false, },],
            "object-shorthand": ["off",],
            'brace-style': ['error', 'allman',],
            'eol-last': ["error", "always",],
            indent:
            [
                'error', 4,
                {
                    outerIIFEBody: 1,
                    FunctionExpression: { body: 1, parameters: 2, },
                    SwitchCase: 1,
                },
            ],
            'no-unused-vars': ['off', { vars: 'local', },],
            'no-multi-spaces': ['off',],
            'no-trailing-spaces': "error",
            quotes: ["off",],
            "space-before-function-paren": ["off",],
            "comma-dangle":
            [
                "error",
                {
                    arrays: "always",
                    objects: "always",
                    imports: "always",
                    exports: "always",
                    functions: "never",
                },
            ],
            "keyword-spacing":
            [
                "error",
                {
                    overrides:
                    {
                        catch: { after: false, },
                    },
                },
            ],
        },
    },
];
