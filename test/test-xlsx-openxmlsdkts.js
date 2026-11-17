if (typeof exports === 'object') {
	var assert = require('assert');
	var alasql = require('..');
	var fs = require('fs');
}

describe('Test XLSX import and export with openxmlsdkts', function () {
	if (typeof exports !== 'object') return;

	var file = __dirname + '/restest-xlsx-openxmlsdkts.xlsx';

	after(function () {
		if (fs.existsSync(file)) fs.unlinkSync(file);
	});

	it('A) Round trip keeps strings, numbers, booleans and dates', async function () {
		var data = [
			{s: 'plain', n: 1, b: true, d: new Date(2024, 0, 15), dt: new Date(2024, 6, 1, 13, 45, 30)},
			{s: 'ünïcødé & <tag> "q"', n: -1234.5678, b: false, d: new Date(1900, 1, 1)},
			{s: '  spaces  ', n: 1e21},
			{s: 'line1\nline2\u0001', n: 0.1 + 0.2},
			{s: '_x0041_ stays literal', n: 0},
		];
		var res = await alasql.promise('SELECT * INTO XLSX("' + file + '") FROM ?', [data]);
		assert.equal(res, 1);

		res = await alasql.promise('SELECT * FROM XLSX("' + file + '")');
		assert.deepStrictEqual(res, data);
	});

	it('B) cellDates:false returns Excel serial numbers', async function () {
		await alasql.promise('SELECT * INTO XLSX("' + file + '") FROM ?', [
			[{d: new Date(2024, 0, 15, 12, 0, 0)}, {d: new Date(1900, 1, 1)}],
		]);
		var res = await alasql.promise('SELECT * FROM XLSX("' + file + '", {cellDates:false})');
		assert.deepStrictEqual(res, [{d: 45306.5}, {d: 32}]);
	});

	it('C) Multiple sheets with special characters in the sheet name', async function () {
		await alasql.promise('SELECT * INTO XLSX("' + file + '", ?) FROM ?', [
			[{sheetid: 'One'}, {sheetid: 'Two & <B>', headers: false}],
			[[{a: 1}], [{b: 'x'}]],
		]);
		var res = await alasql.promise('SELECT * FROM XLSX("' + file + '", {sheetid:"*"})');
		assert.deepStrictEqual(res, [{a: 1, _sheet: 'One'}]);
		res = await alasql.promise(
			'SELECT * FROM XLSX("' + file + '", {sheetid:"Two & <B>", headers:false})'
		);
		assert.deepStrictEqual(res, [{A: 'x'}]);
	});

	it('D) Invalid sheet names are reported', async function () {
		await assert.rejects(
			alasql.promise('SELECT * INTO XLSX("' + file + '", {sheetid:"a/b"}) FROM ?', [[{a: 1}]]),
			/Sheet name cannot contain/
		);
	});

	it('E) Export without a callback throws a clear error', function () {
		assert.throws(function () {
			alasql('SELECT * INTO XLSX("' + file + '") FROM ?', [[{a: 1}]]);
		}, /Export to XLSX is asynchronous/);
	});

	it('F) setXLSX rejects libraries other than openxmlsdkts', function () {
		alasql.setXLSX({read: function () {}, utils: {}});
		try {
			assert.throws(function () {
				alasql('SELECT * INTO XLSX("' + file + '") FROM ?', [[{a: 1}]], function () {});
			}, /must be openxmlsdkts/);
		} finally {
			alasql.setXLSX(null);
		}
	});
});
