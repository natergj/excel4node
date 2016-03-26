module.exports = {
    "parserOptions": {
        "ecmaVersion": 6,
        "sourceType": "module"
    },
    'rules': {
        // TODO add a logger lib to the project
        // 'no-console': 2
        'brace-style': [2, '1tbs'],
        'camelcase': 1,
        'comma-dangle': [2, 'never'],
        'comma-spacing': [2, { 'before': false, 'after': true }],
        'comma-style': [2, 'last'],
        'eqeqeq': 2,
        'indent': [2, 4],
        'key-spacing': [2, { 'beforeColon': false, 'afterColon': true }],
        'keyword-spacing': 2,
        'linebreak-style': [2, 'unix'],
        'no-case-declarations': 0,
        'no-console': 0,
        'no-redeclare': 0,
        'no-underscore-dangle': 0,
        'no-unused-vars': 0,
        'object-curly-spacing': [2, 'always'],
        'quotes': [2, 'single'],
        'semi': [2, 'always'],
        'semi-spacing': [2, { 'before': false, 'after': true }],
        'space-before-blocks': [2, 'always'],
        'space-before-function-paren': [2, { 'anonymous': 'always', 'named': 'never' }],
        'space-in-parens': [2, 'never'],
        'space-infix-ops': 2,
        'strict': 0
    },
    'env': {
        'node': true,
        'es6': true
    },
    'extends': 'eslint:recommended'
};