const Excel = require('exceljs');
const {Client, Pool} = require("pg");

const client = new Client({
    user: "postgres",
    password: "123456",
    host: "192.168.100.15",
    port: 5432,
    database: "openpractice_demo",
});

function getMatches(string, regex, index) {
    index || (index = 1);
    var matches = [];
    var match;
    while ((match = regex.exec(string))) {
        matches.push(match[index]);
    }
    return matches;
};

async function connectPostgresql() {
    await client.connect();
};

async function persistResult(groups, items) {
    try {
        const insertGroupSql = "INSERT INTO icd10_group_temp (id, name, start_code, end_code, parent_id, is_system) VALUES($1, $2, $3, $4, $5, $6)";
        for (const gr of groups) {
            await client.query(insertGroupSql, [gr.id, gr.name, gr.startCode, gr.endCode, gr.parentId, gr.system]);
        }
        
        const insertItemSql = "INSERT INTO icd10_item_temp (id, name_vi, name_en, code, group_id, is_system) VALUES($1, $2, $3, $4, $5, $6)";
        for (const it of items) {
            await client.query(insertItemSql, [it.id, it.vnName, it.enName, it.code, it.groupId, it.system]);
        }
    } catch (e) {
        console.error("****** FAILED TO PERSIST RESULT TO DB ******", e);
    }
};

let itemIdCounter = 1;
let icdItems = [];

let groupIdCounter = 1;
let groupHodler = {};

(async () => {
    await connectPostgresql();
    console.log("=== postgresql connected =====");

    let workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("ICD-10-input.xlsx");
    let workSheet = await workbook.getWorksheet("ICD10");
    let itemCodeRegex = /[A-Za-z0-9.]*/;
    
    workSheet.eachRow({includeEmpty: true}, function (row, rowNumber) {
        if (rowNumber >= 4) {
            let values = row.values;
            let parentCode = values[2].trim();

            let group = {
                code: values[2].trim(),
                startCode: values[2].trim().split("-")[0],
                endCode: values[2].trim().split("-")[1],
                name: values[4],
                parentId: null,
                system: true,
            };
            if (!groupHodler[group.code]) {
                group.id = groupIdCounter;
                groupHodler[group.code] = group;
                groupIdCounter += 1;
            }

            let subGroup = {
                code: values[5].trim(),
                startCode: values[5].trim().split("-")[0],
                endCode: values[5].trim().split("-")[1],
                name: values[7],
                parentCode: values[2].trim(),
                parentId: null,
                system: true,
            };
            if (!groupHodler[subGroup.code]) {
                subGroup.id = groupIdCounter;
                groupHodler[subGroup.code] = subGroup;
                groupIdCounter += 1;
            }
            if (groupHodler[subGroup.parentCode]) groupHodler[subGroup.code].parentId = groupHodler[subGroup.parentCode].id;

            
            let item = {
                id: itemIdCounter,
                code: itemCodeRegex.exec(values[17])[0],
                vnName: values[20],
                enName: values[19],
                groupId: groupHodler[subGroup.code].id,
                system: true,
            };
            icdItems.push(item);
            itemIdCounter += 1;
        }
    });

    let groups = Object.keys(groupHodler).map((key) => groupHodler[key]);

    console.log("=======> Total items: " + icdItems.length);
    console.log("=======> Total groups: " + groups.length);
    // console.log("=========== GROUPS ==========");
    // console.log(groups);
    // console.log("=========== ITEMS ==========");
    // console.log(icdItems);

    await persistResult(groups, icdItems);
    console.log("=======> DONE <> NICK NỢ 1 THÙNG BIA =========");
})().catch((err) => console.error(err));