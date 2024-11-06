const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const Excel = require('exceljs');

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
};

function getMatches(string, regex, index) {
    index || (index = 1);
    var matches = [];
    var match;
    while ((match = regex.exec(string))) {
        matches.push(match[index]);
    }
    return matches;
};

(async () => {
    console.log("=======> Opening Browser...");
    let itemIdCounter = 1;
    let icdItems = [];

    let groupIdCounter = 1;
    let groups = [];
    

    const browser = await puppeteer.launch({
        headless: false,
        ignoreHTTPSErrors: true,
        defaultViewport: {
            width: 1920,
            height: 1080,
        },
        args: [`--window-size=1920,1080`],
    });
    const page = await browser.newPage();
    await page.goto("https://icd.kcb.vn/icd-10/icd10-dual");
    await page.waitForNetworkIdle();

    /**
     * Root Level
     */
    const rootLevelSelector = "div.card-body > .cdk-tree > .mat-nested-tree-node";
    const chapters = await page.$$(rootLevelSelector);
    await page.waitForNetworkIdle();
    await sleep(1000);

    for (let index = 1; index < chapters.length; index++) {
        let chapter = chapters[index];

        let span = await chapter.$("span.cursor-pointer");
        await page.evaluate((element) => {
            element.click();
        }, span);
        console.log("chapter clicked");
        await page.waitForNetworkIdle();

        let title = await page.evaluate((el) => el.textContent, span);
        console.log("=> crawling at: " + title.trim());
        let regex = /\((.*)\)\s*(.*)/g;
        let code = getMatches(title, regex, 1)[0];

        let group = {
            id: groupIdCounter,
            name: getMatches(title, regex, 2)[0].trim(),
            startCode: code.split("-")[0],
            endCode: code.split("-")[1],
            parentId: null,
            system: true,
        };
        groups.push(group);
        groupIdCounter += 1;

        let expandButton = await chapter.$$("button.mat-focus-indicator");
        await page.evaluate((element) => {
            element.click();
        }, expandButton[0]);
        await page.waitForNetworkIdle();
        await sleep(1000);

        /**
         * Second Level
         */
        let secondLevelSelector = `${rootLevelSelector}:nth-child(${index + 1}) > div[role='group'] > .mat-nested-tree-node`;
        let secondLevelItems = await page.$$(secondLevelSelector);
        for (let secondLevelIndex = 0; secondLevelIndex < secondLevelItems.length; secondLevelIndex++) {
            let secondLevelElement = secondLevelItems[secondLevelIndex];
            let span = await secondLevelElement.$("span.cursor-pointer");
            await page.evaluate((element) => {
                element.click();
            }, span);
            console.log("click me...");
            await page.waitForNetworkIdle();
            await page.waitForSelector("div#detail", {visible: true}, 0);

            const detailContent = await page.$(`div#detail`);
            const dContentHtml = await page.evaluate((body) => body.innerHTML, detailContent);
            const $ = await cheerio.load(dContentHtml, null, false);

            let startRowIndex = 2;
            let rows = $("div.row-layout");
            if (rows.length < 3) {
                let span = await secondLevelElement.$("span.cursor-pointer");
                let title = await page.evaluate((el) => el.textContent, span);
                console.log();
                console.error("=== WE SHOULD NEVER GET HERE ===");
                console.log("AT:", title);
                console.log();
                continue;
            }

            let secondLevelGroup = null; 
            for (let i = startRowIndex; i < rows.length; i++) {
                let element = rows[i];
                let clCode = await $(element).find(".column-code div.content").first().text();

                let columnContent = await $(element).find(".column-content").first().children();
                let vnText = await $(columnContent[0]).find("div.hover-container h2").first().text();
                let enText = await $(columnContent[1]).find("div.hover-container h2").first().text();

                if (i == startRowIndex) {
                    secondLevelGroup = {
                        id: groupIdCounter,
                        name: vnText.trim(),
                        startCode: clCode.trim().split("-")[0],
                        endCode: clCode.trim().split("-")[1],
                        parentId: group.id,
                        system: true,
                    };
                    groups.push(secondLevelGroup);
                    groupIdCounter += 1;
                } else {
                    let vnText = await $(columnContent[0]).find("div.hover-container div.content").first().text();
                    let enText = await $(columnContent[1]).find("div.hover-container div.content").first().text();
                    let icdItem = {
                        id: itemIdCounter,
                        vnName: vnText.trim(),
                        enName: enText.trim(),
                        code: clCode.trim(),
                        groupId: secondLevelGroup.id,
                        system: true,
                    };
                    icdItems.push(icdItem);
                    itemIdCounter += 1;
                }
            }
        }

        if (index == 3) break;
    }

    console.log("=======> Total items: " + icdItems.length);
    console.log("=======> Total groups: " + groups.length);
    // console.log("=========== GROUPS ==========");
    // console.log(groups);
    // console.log("=========== ITEMS ==========");
    // console.log(icdItems);
    
    await exportResults(icdItems, groups);
    // await browser.close();
    console.log("=======> DONE <> NICK NỢ 1 THÙNG BIA =========");
})();

const exportResults = async (icdResults, groups) => {
	try {
		console.log("=======> Begin export results to file...");
		let workbook = new Excel.Workbook();
		let worksheet = workbook.addWorksheet('Ma ICD');
		worksheet.columns = [
            {header: "Mã", key: "code", width: 10},
            {header: "Tên", key: "vnName", width: 60},
            {header: "Tên Tiếng Anh", key: "enName", width: 60},
            {header: "Thuộc Nhóm", key: "groupId", width: 60}
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
            {header: "Tên Nhóm", key: "name", width: 60},
            {header: "Mã Bắt Đầu", key: "startCode", width: 50},
            {header: "Mã Kết Thúc", key: "endCode", width: 50},
            {header: "Nhóm Cha", key: "parentId", width: 60},
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
