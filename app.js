const fs = require('fs');

const xlsx = require('node-xlsx');
const json2xls = require('json2xls');

const filePath = process.argv[2];
const sheetName = process.argv[3];

const confluenceBaseUrl = 'https://interlink.atlassian.net/wiki'; // modify this to client's url 

const getWorksheet = (path, sheetName) => {
    try {
        const worksheet = xlsx.parse(path);

        if(!sheetName) {
            return worksheet[0].data;
        } 

        const sheetIndex = worksheet.findIndex(el => el.name === sheetName);

        if(sheetIndex === -1) {
            throw 'Invalid Sheet Name!';
        }
        
        return worksheet[sheetIndex].data;
    } catch(err) {
        throw `Error: ${err}`;
    }
};

const buildJSON = (filePath, sheetName) => {
    try {
        const data = getWorksheet(filePath, sheetName);

        const jsonArray = [];
        const iterationDataset = data.slice(1,data.length);

        const keyArray = data[0];

        const pageIdIndex = keyArray.findIndex(el => el === 'Page ID');
        const attachmentTitleIndex = keyArray.findIndex(el => el === 'Attachment Title');

        for (const record of iterationDataset) {
            let currentObject = {};

            const pageId = record[pageIdIndex];
            const attachmentTitle = record[attachmentTitleIndex];

            keyArray.forEach((el, index) => {
                currentObject[el] = record[index];
            });

            currentObject['Download Link'] = `${confluenceBaseUrl}/download/attachments/${pageId}/${encodeURIComponent(attachmentTitle)}`;

            jsonArray.push(currentObject);
        }

        fs.writeFileSync('attachmentsData.json', JSON.stringify(jsonArray))

        const excelData = json2xls(jsonArray);

        fs.writeFileSync('attachmentsData.xlsx', excelData, 'binary');
    } catch(err) {
        console.error(err)
    }
}

buildJSON(filePath, sheetName);
