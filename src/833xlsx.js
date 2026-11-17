/*
	XLSX import and export based on openxmlsdkts
	https://github.com/EricWhiteDev/OpenXmlSdkTs

	Workbooks are handled in memory with a small cell model:
		{SheetNames: ['Sheet1'], Sheets: {Sheet1: {'!ref': 'A1:B2', A1: {t: 's', v: 'a'}}}}
	Cell types: 's' string, 'n' number, 'b' boolean, 'd' date, 'e' error.
	openxmlsdkts is promise based, so reading and writing XLSX files is always asynchronous.
*/

var XLSX_NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
var XLSX_NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var XLSX_NS_PKGREL = 'http://schemas.openxmlformats.org/package/2006/relationships';
var XLSX_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Built-in number formats that represent dates or times
var XLSX_DATE_FMT_IDS = [
	14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
	52, 53, 54, 55, 56, 57, 58,
];

// Style indexes used by the writer (see xlsxStylesXml)
var XLSX_STYLE_DATE = 1;
var XLSX_STYLE_DATETIME = 2;

function xlsxEscapeXml(s) {
	return String(s)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

// Excel stores characters that are not allowed in XML as _xHHHH_
function xlsxEncodeText(s) {
	return xlsxEscapeXml(
		String(s)
			.replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
			.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, function (c) {
				return '_x' + ('000' + c.charCodeAt(0).toString(16).toUpperCase()).slice(-4) + '_';
			})
	);
}

function xlsxDecodeText(s) {
	return s.replace(/_x([0-9A-Fa-f]{4})_/g, function (m, hex) {
		return String.fromCharCode(parseInt(hex, 16));
	});
}

/**
	Convert an Excel serial date number to a JavaScript Date (local time)
	@param {number} serial
	@param {boolean} date1904 Workbook uses the 1904 date system
	@return {Date}
*/
function xlsxSerialToDate(serial, date1904) {
	if (date1904) serial += 1462;
	var days = Math.floor(serial);
	var ms = Math.round((serial - days) * 864e5);
	// Excel treats 1900 as a leap year, so serials before 1900-03-01 are shifted by one day
	if (days >= 1 && days < 61) days += 1;
	return new Date(1899, 11, 30 + days, 0, 0, 0, ms);
}

/**
	Convert a JavaScript Date (local time) to an Excel serial date number
	@param {Date} d
	@return {number}
*/
function xlsxDateToSerial(d) {
	var days = Math.round(
		(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5
	);
	if (days >= 2 && days <= 60) days -= 1;
	var ms =
		((d.getHours() * 60 + d.getMinutes()) * 60 + d.getSeconds()) * 1000 + d.getMilliseconds();
	return days + ms / 864e5;
}

function xlsxIsDateFormat(code) {
	var s = String(code)
		.split(';')[0]
		.replace(/"[^"]*"/g, '')
		.replace(/\\./g, '')
		.replace(/[_*]./g, '')
		.replace(/\[(h+|m+|s+)\]/gi, 'h')
		.replace(/\[[^\]]*\]/g, '');
	if (/^\s*general\s*$/i.test(s)) return false;
	return /[dmyhs]/i.test(s);
}

/**
	Convert data returned by loadBinaryFile to an input accepted by openxmlsdkts
*/
function xlsxToPackageInput(data) {
	if (typeof Blob !== 'undefined' && data instanceof Blob) {
		return data;
	}
	if (data instanceof ArrayBuffer || ArrayBuffer.isView(data)) {
		return new Blob([data]);
	}
	if (typeof data === 'string') {
		// Base64 content (React Native) is accepted as is
		if (data.slice(0, 2) !== 'PK') {
			return data;
		}
		// Binary string (one char per byte)
		var bytes = new Uint8Array(data.length);
		for (var i = 0; i < data.length; i++) bytes[i] = data.charCodeAt(i) & 0xff;
		return new Blob([bytes]);
	}
	throw new Error('Unsupported input for XLSX file');
}

/**
	Open an XLSX package. openxmlsdkts rejects files with zip entries that have no content type
	(for example "[trash]/0000.dat" left by some editors), so such entries are removed and the
	file is opened again.
*/
async function xlsxOpenPackage(ox, input) {
	try {
		return await ox.SmlPackage.open(input);
	} catch (err) {
		if (!ox.JSZip || typeof input === 'string' || !/Content type not found/.test(err.message)) {
			throw err;
		}
	}

	var zip = await ox.JSZip.loadAsync(await input.arrayBuffer());
	var ctRoot = ox.XDocument.parse(await zip.file('[Content_Types].xml').async('string')).root;
	var overrides = {};
	var defaults = {};
	ctRoot.elements(ox.CT.Override).forEach(function (el) {
		overrides[el.attribute('PartName').value] = true;
	});
	ctRoot.elements(ox.CT.Default).forEach(function (el) {
		defaults[el.attribute('Extension').value] = true;
	});
	Object.keys(zip.files).forEach(function (name) {
		if (name.endsWith('/') || name === '[Content_Types].xml') return;
		if (!overrides['/' + name] && !defaults[name.split('.').pop()]) {
			zip.remove(name);
		}
	});

	return ox.SmlPackage.open(await zip.generateAsync({type: 'blob', compression: 'STORE'}));
}

function xlsxText(el, S) {
	var t = el.element(S.t);
	if (t) return xlsxDecodeText(t.value);
	return xlsxDecodeText(
		el
			.elements(S.r)
			.map(function (r) {
				var rt = r.element(S.t);
				return rt ? rt.value : '';
			})
			.join('')
	);
}

/**
	Read XLSX content into the in-memory workbook model
	@param {object} ox openxmlsdkts module
	@param {string|ArrayBuffer|Uint8Array|Blob} data File content
	@param {object} opts Options. cellDates: return date formatted cells as Date objects
	@return {Promise<object>} Workbook
*/
async function xlsxReadWorkbook(ox, data, opts) {
	opts = opts || {};
	var S = ox.S;
	var pkg = await xlsxOpenPackage(ox, xlsxToPackageInput(data));
	var wbPart = await pkg.workbookPart();
	if (!wbPart) {
		throw new Error('Invalid XLSX file: no workbook part found');
	}
	var wbRoot = (await wbPart.getXDocument()).root;

	var wbPr = wbRoot.element(S.workbookPr);
	var date1904 = !!wbPr && /^(1|true)$/.test((wbPr.attribute('date1904') || {}).value);

	var sst = [];
	var sstPart = await wbPart.sharedStringTablePart();
	if (sstPart) {
		sst = (await sstPart.getXDocument()).root.elements(S.si).map(function (si) {
			return xlsxText(si, S);
		});
	}

	// Style index -> is date format
	var dateStyles = [];
	if (opts.cellDates) {
		var stylesPart = await wbPart.workbookStylesPart();
		if (stylesPart) {
			var stRoot = (await stylesPart.getXDocument()).root;
			var customFmts = {};
			var numFmts = stRoot.element(S.numFmts);
			if (numFmts) {
				numFmts.elements(S.numFmt).forEach(function (nf) {
					customFmts[nf.attribute('numFmtId').value] = nf.attribute('formatCode').value;
				});
			}
			var cellXfs = stRoot.element(S.cellXfs);
			if (cellXfs) {
				dateStyles = cellXfs.elements(S.xf).map(function (xf) {
					var id = +((xf.attribute('numFmtId') || {}).value || 0);
					if (customFmts[id] !== undefined) return xlsxIsDateFormat(customFmts[id]);
					return XLSX_DATE_FMT_IDS.indexOf(id) > -1;
				});
			}
		}
	}

	var wb = {SheetNames: [], Sheets: {}};
	var sheetsEl = wbRoot.element(S.sheets);
	var sheetEls = sheetsEl ? sheetsEl.elements(S.sheet) : [];

	for (var i = 0; i < sheetEls.length; i++) {
		var name = sheetEls[i].attribute('name').value;
		var rid = sheetEls[i].attribute(ox.R.id);
		var sheetPart = rid ? await wbPart.getPartById(rid.value) : undefined;
		wb.SheetNames.push(name);
		wb.Sheets[name] = {};
		if (sheetPart && sheetPart.getContentType() === XLSX_CT + 'worksheet+xml') {
			readSheet((await sheetPart.getXDocument()).root, wb.Sheets[name]);
		}
	}

	return wb;

	function readSheet(wsRoot, cells) {
		var sheetData = wsRoot.element(S.sheetData);
		if (!sheetData) return;
		var minCol = Infinity,
			maxCol = -1,
			minRow = Infinity,
			maxRow = -1;
		var rowNum = 0;

		sheetData.elements(S.row).forEach(function (rowEl) {
			var r = rowEl.attribute('r');
			rowNum = r ? +r.value : rowNum + 1;
			var colNum = -1;

			rowEl.elements(S.c).forEach(function (c) {
				var ref = c.attribute('r');
				var m = ref && ref.value.match(/^([A-Z]+)(\d+)$/);
				colNum = m ? alasql.utils.xlscn(m[1]) : colNum + 1;
				var row = m ? +m[2] : rowNum;

				var cell = readCell(c);
				if (!cell) return;

				cells[alasql.utils.xlsnc(colNum) + row] = cell;
				if (colNum < minCol) minCol = colNum;
				if (colNum > maxCol) maxCol = colNum;
				if (row < minRow) minRow = row;
				if (row > maxRow) maxRow = row;
			});
		});

		if (maxRow > -1) {
			cells['!ref'] =
				alasql.utils.xlsnc(minCol) + minRow + ':' + alasql.utils.xlsnc(maxCol) + maxRow;
		}
	}

	function readCell(c) {
		var t = (c.attribute('t') || {}).value || 'n';
		var vEl = c.element(S.v);
		var fEl = c.element(S.f);
		var cell;

		if (t === 'inlineStr') {
			var is = c.element(S.is);
			cell = {t: 's', v: is ? xlsxText(is, S) : ''};
		} else if (!vEl) {
			return;
		} else if (t === 's') {
			cell = {t: 's', v: sst[+vEl.value]};
		} else if (t === 'str') {
			cell = {t: 's', v: xlsxDecodeText(vEl.value)};
		} else if (t === 'b') {
			cell = {t: 'b', v: vEl.value === '1' || vEl.value === 'true'};
		} else if (t === 'e') {
			cell = {t: 'e', v: vEl.value};
		} else if (t === 'd') {
			var d = new Date(vEl.value);
			cell = opts.cellDates ? {t: 'd', v: d} : {t: 'n', v: xlsxDateToSerial(d)};
		} else {
			cell = {t: 'n', v: +vEl.value};
			var s = c.attribute('s');
			if (opts.cellDates && s && dateStyles[+s.value]) {
				cell = {t: 'd', v: xlsxSerialToDate(cell.v, date1904)};
			}
		}

		// Keep plain formulas so they survive a round trip (sourcefilename option)
		if (fEl && fEl.value && !fEl.attribute('si')) {
			cell.f = fEl.value;
		}
		return cell;
	}
}

function xlsxStylesXml() {
	return (
		'<styleSheet xmlns="' +
		XLSX_NS_MAIN +
		'">' +
		'<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
		'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
		'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
		'<cellXfs count="3">' +
		'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
		'<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
		'<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
		'</cellXfs>' +
		'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
		'</styleSheet>'
	);
}

function xlsxCheckSheetName(name) {
	if (!name || name.length > 31) {
		throw new Error('Sheet name must be between 1 and 31 characters: ' + name);
	}
	if (/[\\\/?*\[\]:]/.test(name)) {
		throw new Error('Sheet name cannot contain : \\ / ? * [ ]: ' + name);
	}
}

/**
	Write the in-memory workbook model as XLSX
	@param {object} ox openxmlsdkts module
	@param {object} wb Workbook
	@return {Promise<Uint8Array>} XLSX file content
*/
async function xlsxWriteWorkbook(ox, wb) {
	var sst = [];
	var sstIndex = new Map();

	function sharedString(s) {
		if (!sstIndex.has(s)) {
			sstIndex.set(s, sst.length);
			sst.push(s);
		}
		return sstIndex.get(s);
	}

	function cellXml(ref, cell) {
		var v = cell.v;
		var t = cell.t;
		var f = cell.f ? '<f>' + xlsxEscapeXml(cell.f) + '</f>' : '';

		if (v === undefined || v === null) {
			return f ? '<c r="' + ref + '">' + f + '</c>' : '';
		}
		if (v instanceof Date) {
			if (isNaN(v.getTime())) return '';
			var serial = xlsxDateToSerial(v);
			var style = serial % 1 === 0 ? XLSX_STYLE_DATE : XLSX_STYLE_DATETIME;
			return '<c r="' + ref + '" s="' + style + '">' + f + '<v>' + serial + '</v></c>';
		}
		if (t === 'e') {
			return '<c r="' + ref + '" t="e">' + f + '<v>' + xlsxEscapeXml(v) + '</v></c>';
		}
		if (typeof v === 'number') {
			if (!isFinite(v)) {
				var err = isNaN(v) ? '#NUM!' : '#DIV/0!';
				return '<c r="' + ref + '" t="e">' + f + '<v>' + err + '</v></c>';
			}
			return '<c r="' + ref + '">' + f + '<v>' + v + '</v></c>';
		}
		if (typeof v === 'boolean') {
			return '<c r="' + ref + '" t="b">' + f + '<v>' + (v ? 1 : 0) + '</v></c>';
		}
		if (f) {
			return '<c r="' + ref + '" t="str">' + f + '<v>' + xlsxEncodeText(v) + '</v></c>';
		}
		return '<c r="' + ref + '" t="s"><v>' + sharedString(String(v)) + '</v></c>';
	}

	function sheetXml(cells) {
		var rows = {};
		var minCol = Infinity,
			maxCol = -1,
			minRow = Infinity,
			maxRow = -1;

		Object.keys(cells).forEach(function (ref) {
			var m = ref.match(/^([A-Z]+)(\d+)$/);
			if (!m) return;
			var col = alasql.utils.xlscn(m[1]);
			var row = +m[2];
			(rows[row] = rows[row] || []).push({col: col, ref: ref});
			if (col < minCol) minCol = col;
			if (col > maxCol) maxCol = col;
			if (row < minRow) minRow = row;
			if (row > maxRow) maxRow = row;
		});

		var dimension =
			maxRow > -1
				? alasql.utils.xlsnc(minCol) + minRow + ':' + alasql.utils.xlsnc(maxCol) + maxRow
				: 'A1';

		var xml =
			'<worksheet xmlns="' +
			XLSX_NS_MAIN +
			'" xmlns:r="' +
			XLSX_NS_REL +
			'"><dimension ref="' +
			dimension +
			'"/><sheetData>';

		Object.keys(rows)
			.map(Number)
			.sort(function (a, b) {
				return a - b;
			})
			.forEach(function (row) {
				var cs = rows[row]
					.sort(function (a, b) {
						return a.col - b.col;
					})
					.map(function (c) {
						return cellXml(c.ref, cells[c.ref]);
					})
					.join('');
				if (cs) xml += '<row r="' + row + '">' + cs + '</row>';
			});

		return xml + '</sheetData></worksheet>';
	}

	function part(name, contentType, xml) {
		return (
			'<pkg:part pkg:name="' +
			name +
			'" pkg:contentType="' +
			contentType +
			'"><pkg:xmlData>' +
			xml +
			'</pkg:xmlData></pkg:part>'
		);
	}

	var relsCT = 'application/vnd.openxmlformats-package.relationships+xml';
	var sheetsXml = '';
	var wbRelsXml = '';
	var dataParts = [];

	wb.SheetNames.forEach(function (name, idx) {
		xlsxCheckSheetName(name);
		var n = idx + 1;
		sheetsXml +=
			'<sheet name="' + xlsxEscapeXml(name) + '" sheetId="' + n + '" r:id="rId' + n + '"/>';
		wbRelsXml +=
			'<Relationship Id="rId' +
			n +
			'" Type="' +
			XLSX_NS_REL +
			'/worksheet" Target="worksheets/sheet' +
			n +
			'.xml"/>';
		dataParts.push({
			uri: '/xl/worksheets/sheet' + n + '.xml',
			contentType: XLSX_CT + 'worksheet+xml',
			xml: sheetXml(wb.Sheets[name] || {}),
		});
	});

	var n = wb.SheetNames.length;
	wbRelsXml +=
		'<Relationship Id="rId' +
		(n + 1) +
		'" Type="' +
		XLSX_NS_REL +
		'/styles" Target="styles.xml"/>' +
		'<Relationship Id="rId' +
		(n + 2) +
		'" Type="' +
		XLSX_NS_REL +
		'/sharedStrings" Target="sharedStrings.xml"/>';

	var sstXml =
		'<sst xmlns="' +
		XLSX_NS_MAIN +
		'" count="' +
		sst.length +
		'" uniqueCount="' +
		sst.length +
		'">' +
		sst
			.map(function (s) {
				return '<si><t xml:space="preserve">' + xlsxEncodeText(s) + '</t></si>';
			})
			.join('') +
		'</sst>';

	var flatOpc =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
		part(
			'/_rels/.rels',
			relsCT,
			'<Relationships xmlns="' +
				XLSX_NS_PKGREL +
				'"><Relationship Id="rId1" Type="' +
				XLSX_NS_REL +
				'/officeDocument" Target="xl/workbook.xml"/></Relationships>'
		) +
		part(
			'/xl/workbook.xml',
			XLSX_CT + 'sheet.main+xml',
			'<workbook xmlns="' +
				XLSX_NS_MAIN +
				'" xmlns:r="' +
				XLSX_NS_REL +
				'"><sheets>' +
				sheetsXml +
				'</sheets></workbook>'
		) +
		part(
			'/xl/_rels/workbook.xml.rels',
			relsCT,
			'<Relationships xmlns="' + XLSX_NS_PKGREL + '">' + wbRelsXml + '</Relationships>'
		) +
		part('/xl/styles.xml', XLSX_CT + 'styles+xml', xlsxStylesXml()) +
		'</pkg:package>';

	dataParts.push({
		uri: '/xl/sharedStrings.xml',
		contentType: XLSX_CT + 'sharedStrings+xml',
		xml: sstXml,
	});

	var pkg = await ox.SmlPackage.open(flatOpc);

	// Worksheets and shared strings can be large, so they are added as unparsed XML.
	// openxmlsdkts stores parts it has not parsed yet as objects with an async('string')
	// method (JSZip entries) and writes their content as is when saving.
	dataParts.forEach(function (p) {
		var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + p.xml;
		pkg.addPart(p.uri, p.contentType, 'xml', {
			async: function () {
				return Promise.resolve(xml);
			},
		});
	});

	// saveToBlobAsync() always compresses (DEFLATE level 6), which is the slowest part of an
	// export. The zip is generated here without compression when openxmlsdkts allows it
	// (saveToZip is private in its typings, so other versions of openxmlsdkts fall back).
	if (typeof pkg.saveToZip === 'function') {
		var zip = await pkg.saveToZip();
		return zip.generateAsync({type: 'uint8array', compression: 'STORE'});
	}
	var blob = await pkg.saveToBlobAsync();
	return new Uint8Array(await blob.arrayBuffer());
}

/**
	Export to XLSX function
	@function
	@param {string|object} filename Filename or options
	@param {object|undefined} opts Options or undefined
	@param {array} data Data
	@param {array} columns Columns
	@param {callback} cb Callback function
	@return {number} Number of files processed
*/

alasql.into.XLSX = function (filename, opts, data, columns, cb) {
	if (!cb || cb.sync) {
		throw new Error(
			'Export to XLSX is asynchronous: use alasql.promise() or pass a callback to alasql()'
		);
	}

	/** @type {number} result */
	var res = 1;
	opts = opts || {};

	if (deepEqual(columns, [{columnid: '_'}])) {
		data = data.map(function (dat) {
			return dat._;
		});
		columns = undefined;
	}

	filename = alasql.utils.autoExtFilename(filename, 'xlsx', opts);

	var ox = getXLSX();

	/* If called without filename, use opts */
	if (typeof filename == 'object') {
		opts = filename;
		filename = undefined;
	}

	/** @type {object} Workbook */
	var wb = {SheetNames: [], Sheets: {}};

	function fail(err) {
		cb(null, err);
	}

	if (opts.sourcefilename) {
		alasql.utils.loadBinaryFile(
			opts.sourcefilename,
			true,
			function (content) {
				xlsxReadWorkbook(ox, content, {}).then(function (sourceWb) {
					wb = sourceWb;
					doExport();
				}, fail);
			},
			fail
		);
	} else {
		doExport();
	}

	return res;

	/**
		Export workbook
		@function
	*/
	function doExport() {
		try {
			/*
				If opts is array of arrays then this is a
				multisheet workboook, else it is a singlesheet
			*/
			if (typeof opts == 'object' && Array.isArray(opts)) {
				if (data && data.length > 0) {
					data.forEach(function (dat, idx) {
						prepareSheet(opts[idx], dat, undefined, idx + 1);
					});
				}
			} else {
				prepareSheet(opts, data, columns, 1);
			}
		} catch (err) {
			return fail(err);
		}

		// The callback is called outside of the promise chain so errors thrown by it are not reported twice
		xlsxWriteWorkbook(ox, wb).then(saveWorkbook, fail);
	}

	/**
		Prepare sheet
		@params {object} opts
		@params {array|object} data
		@params {array} columns Columns
	*/
	function prepareSheet(opts, data, columns, idx) {
		/** Default options for sheet */
		var opt = {sheetid: 'Sheet ' + idx, headers: true};
		alasql.utils.extend(opt, opts);

		var dataLength = Object.keys(data).length;

		// Generate columns if they are not defined
		if (!columns || columns.length == 0) {
			if (dataLength > 0) {
				columns = Object.keys(data[0]).map(function (columnid) {
					return {columnid: columnid};
				});
			} else {
				columns = [];
			}
		}

		if (wb.SheetNames.indexOf(opt.sheetid) === -1) {
			wb.SheetNames.push(opt.sheetid);
			wb.Sheets[opt.sheetid] = {};
		}
		var cells = wb.Sheets[opt.sheetid];

		var range = 'A1';
		if (opt.range) range = opt.range;

		var col0 = alasql.utils.xlscn(range.match(/[A-Z]+/)[0]);
		var i = +range.match(/[0-9]+/)[0];

		if (opt.headers) {
			columns.forEach(function (col, idx) {
				cells[alasql.utils.xlsnc(col0 + idx) + '' + i] = {
					t: 's',
					v: col.columnid.trim(),
				};
			});
			i++;
		}

		for (var j = 0; j < dataLength; j++) {
			columns.forEach(function (col, idx) {
				cells[alasql.utils.xlsnc(col0 + idx) + '' + i] = {v: data[j][col.columnid]};
			});
			i++;
		}
	}

	/**
		Save Workbook
		@params {Uint8Array} content XLSX file content
	*/
	function saveWorkbook(content) {
		if (typeof filename == 'undefined') {
			// No filename: return the file content
			cb(content);
			return;
		}

		if (utils.isNode || utils.isMeteorServer) {
			//*not-for-browser/*
			var fs = require('fs');
			fs.writeFile(filename, content, function (err) {
				if (err) return fail(err);
				cb(res);
			});
			//*/
		} else {
			saveAs(new Blob([content], {type: XLSX_MIME}), filename);
			cb(res);
		}
	}
};
