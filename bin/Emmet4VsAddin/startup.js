/**
 * import underscore module
 * import core module
 */
context.require(context.Root + '\\javascript\\underscore.js');
context.require(context.Root + '\\javascript\\core.js');

/**
 * Set module loader and construct editor instance
 *
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Emmet4VsAddin
 */
var editor = emmet.exec(function(require, _) {
    
    function getModuleLoader() {
        var paths = _.toArray(arguments);
        return function (module) {
            _.find(paths, function (path) {
                var url = path + '\\' + module + '.js';
                return context.require(url);
            });
        };
    }

	emmet.setModuleLoader(getModuleLoader(
        context.Root,
        context.Root + '\\javascript',
        context.Root + '\\javascript\\loaders',
        context.Root + '\\javascript\\parsers',
        context.Root + '\\javascript\\parsers\\editTree',
        context.Root + '\\javascript\\resolvers'
    ));

	return require('editor');
});

/**
 * Load system snippets, load buildin actions
 *
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Emmet4VsAddin
 */
emmet.exec(function (require, _) {

    require('json');
    require('bootstrap').loadSystemSnippets(require('file').read(require('file').createPath(context.Root, 'snippets.json')));
    require('filters\\format');
    require('filters\\html');
    require('filters\\xsl');
    require('processors\\tag-name');
    require('processors\\pasted-content');
    require('processors\\resource-matcher');
    require('actions\\expandAbbreviation');
    require('actions\\wrapWithAbbreviation');
    require('actions\\matchPair');
    require('actions\\editPoints');
    require('actions\\selectItem');
    require('actions\\selectLine');
    require('actions\\lineBreak');
    require('actions\\mergeLines');
    require('actions\\toggleComment');
    require('actions\\splitJoinTag');
    require('actions\\removeTag');
    require('actions\\evaluateMath');
    require('actions\\increment_decrement');
    require('actions\\base64');
    require('actions\\reflectCSSValue');
    require('actions\\updateImageSize');
    context.startup(JSON.stringify(require('actions').getMenu()));

});
