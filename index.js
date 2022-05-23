const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const Excel = require('exceljs');

(async () => {
	try {
		console.log("=======> Opening Browser...");
		let results = [];
		const browser = await puppeteer.launch({headless: false});
		const page = await browser.newPage();
		await page.goto('http://icd.kcb.vn/');

		const chapters = await page.$$('#uLeftMenu > ul > li');
	    for (let i = 0, length = chapters.length; i < length; i++) {
	        let dropDownMenu = await page.$$(`#uLeftMenu > ul > li:nth-child(${i + 1}) > div > i`);
	        await dropDownMenu[0].click();
	        await page.waitForSelector(`#uLeftMenu > ul > li:nth-child(${i + 1}) > div > i.glyphicon-chevron-down`);

	        const chapters = await page.$$(`#uLeftMenu > ul > li:nth-child(${i + 1}) > div > div`);
	        await chapters[0].click();

			const bodyHandle = await page.$(`#uLeftMenu > ul > li:nth-child(${i + 1}) > div`);
			const html = await page.evaluate(body => body.innerHTML, bodyHandle);
	        const $ = cheerio.load(html);
	        let title = $("a").text();
	        console.log("=======> Begin crawl: " + title);

	        await waitUntilContentLoaded(page);

	        const subChapters = await page.$$(`#uLeftMenu > ul > li:nth-child(${i + 1}) > ul > li`);
			for (let subChapterElement of subChapters) {
			    await subChapterElement.click();
			    await waitUntilContentLoaded(page);

			    const content = await page.content();
			    const $ = cheerio.load(content);

				let rows = $('div#dContent').children();

			    let id = "";
			    let name = "";
			    let engName = "";
			    let groupName = $('div#dContent').children(".group-header").children("div").first().text();
			    let desc = "";
			    let engDesc = "";

			    for (let i = 0; i < rows.length; i++) {
			    	let element = rows[i];
			    	let clazzName = await $(element).attr("class");
			    	if (clazzName.includes("line-item")) {
			    		if (id) {
			    			results.push({id, name, engName, groupName, desc, engDesc});
			    		}
			    		id = $(element).children(clazzName.includes("group") ? ".type-header" : ".item-header").text();
			    		$(element).find(".item-name").each((index, el) => {
			    			let value = $(el).text();
			    			if ($(el).attr("style").includes("left")) {
			    				name = value;
			    			} else {
			    				engName = value;
			    			}
			    		});
			    	} else if (clazzName.includes("item-detail")) {
			    		$(element).find(".item-include").each((index, el) => {
			    			let value = $(el).text();
			    			if ($(el).attr("style") && $(el).attr("style").includes("left")) {
			    				desc = value;
			    			} else if ($(el).attr("style") && $(el).attr("style").includes("right")) {
			    				engDesc = value;
			    			} else {
			    				desc = value;
			    			}
			    		});
			    	}
			    	if ((i == (rows.length - 1))) {
			    		results.push({id, name, engName, groupName, desc, engDesc});
			    	};
			    }
			}
			console.log("=======> Finished crawl: " + title);
	    }
	    console.log("=======> Total results: " + results.length);
	    await exportResults(results);
		await browser.close();
		console.log("=======> DONE <=========");
	} catch (e) {
		console.log("========== FAILED =========", e);
	}
})();

const waitUntilContentLoaded = async (page) => {
	return await page.waitForSelector('#divMain > div > div.row.form-inline > div.col-xs-8 > div.page-refresh', {hidden : true}, 0);
}

const exportResults = async (results) => {
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
		results.forEach((e, index) => {
			worksheet.addRow({
				...e
			});
		});
		worksheet.getRow(1).eachCell((cell) => {
			cell.font = {bold: true};
		});
		await workbook.xlsx.writeFile('ICD-results.xlsx');
		console.log("=======> Finished export result to file...");
	} catch (e) {
		console.log("=======> Error while exporting results to file...: ", e);
	}
}
