// Replaces the Node "fs" module when bundling openxmlsdkts for browsers.
// ltxmlts only uses fs to parse XML from files, which alasql does not do.
module.exports = {};
