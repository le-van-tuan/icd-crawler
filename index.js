const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const Excel = require('exceljs');

(async () => {
	try {
		console.log("=======> Opening Browser...");
		let icdResults = [];
		let groups = [];

		const browser = await puppeteer.launch({headless: false});
		const page = await browser.newPage();
		await page.goto("https://icd.kcb.vn/icd-10/icd10-dual");

		

	    console.log("=======> Total results: " + icdResults.length);
		console.log("=======> Total groups: " + groups.length);
		// await browser.close();
		console.log("=======> DONE <=========");
	} catch (e) {
		console.log("========== FAILED =========", e);
	}
})();

const waitUntilContentLoaded = async (page) => {
	return await page.waitForSelector('#divMain > div > div.row.form-inline > div.col-xs-8 > div.page-refresh', {hidden : true}, 0);
}

const exportResults = async (icdResults, groups) => {
	try {
		console.log("=======> Begin export results to file...");
		let workbook = new Excel.Workbook();
		let worksheet = workbook.addWorksheet('Ma ICD');
		worksheet.columns = [
			{header: 'Mã', key: 'id', width: 10},
			{header: 'Tên', key: 'name', width: 60},
			{header: 'Tên Tiếng Anh', key: 'engName', width: 60},
			{header: 'Tên Nhóm', key: 'groupName', width: 60},
			{header: 'Mô Tả', key: 'desc', width: 100},
			{header: 'Mô Tả Tiếng Anh', key: 'engDesc', width: 100}
		];
		icdResults.forEach((e, index) => {
			worksheet.addRow({
				...e
			});
		});
		worksheet.getRow(1).eachCell((cell) => {
			cell.font = {bold: true};
		});

		let groupsWs = workbook.addWorksheet('Nhom');
		groupsWs.columns = [
			{header: 'Tên Nhóm', key: 'name', width: 60},
			{header: 'Tên Nhóm Tiếng Anh', key: 'enName', width: 60},
			{header: 'Nhóm Cha', key: 'parent', width: 60},
			{header: 'Nhóm Cha Tiếng Anh', key: 'enParent', width: 60},
			{header: 'Mã Bắt Đầu', key: 'startID', width: 50},
			{header: 'Mã Kết Thúc', key: 'endID', width: 50}
		];
		groups.forEach((e, index) => {
			groupsWs.addRow({
				...e
			});
		});
		groupsWs.getRow(1).eachCell((cell) => {
			cell.font = {bold: true};
		});

		await workbook.xlsx.writeFile('ICD-results.xlsx');
		console.log("=======> Finished export result to file...");
	} catch (e) {
		console.log("=======> Error while exporting results to file...: ", e);
	}
}
