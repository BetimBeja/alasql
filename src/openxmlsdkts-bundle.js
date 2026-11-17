// Entry point used by build.sh to bundle openxmlsdkts into dist/openxmlsdkts.js and
// dist/openxmlsdkts.min.js. JSZip (already a dependency of openxmlsdkts) is exposed so
// alasql can repair XLSX files that openxmlsdkts refuses to open.
export * from 'openxmlsdkts';
export {default as JSZip} from 'jszip';
