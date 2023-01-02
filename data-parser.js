const reader = require('xlsx');
const fs = require('fs');

const year = "2023";
const inputFilePath = `./${year}/personal-data-extraction-test.xlsx`;
const outputFilePath = `./${year}/output.json`;
const organizedFilePath = `./${year}/organized.json`;

function main() {
    let data = parseSpreadsheet();
    if(data.length > 0) {
        writeOutputToFile(data);
    }
}

main();

function parseSpreadsheet() {
    let data = [];
    try {
        console.log(`Reading from file: ${inputFilePath}...`);
        const file = reader.readFile(inputFilePath);
        const sheets = file.SheetNames;
        
        console.log("Parsing data...");
        for(let i = 0; i < sheets.length; i++) {
            const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], {raw: false});
            temp.forEach((res) => {
                data.push({
                    date: res.Date,
                    usedWorkHour: res['Used dedicated work hour? (yes/no)'],
                    continuedAfter: res['Continued after hour? (yes/no)'],
                    projectPlanning: res['Project Planning (m)'],
                    gameDev: res['Game Dev (m)'],
                    generalProgramming: res['General Programming (m)'],
                    vfx: res['VFX / Art (m)'],
                    sfx: res['SFX (m)'],
                    writing: res['Writing (m)'],
                    researchAndStudying: res['Research / Studying (m)'],
                    happiness: res['Happiness (1-5)'],
                    weight: res['Weight (lb)'] || null,
                });
            });
        }

        console.log(`Records created: ${data.length}`);
        return data;
    } catch(e) {
        console.log(e);
        console.log("ERROR PARSING DATA FROM: ", inputFilePath);
        return [];
    }
}

async function writeOutputToFile(data) {
    console.log(`Writing to file: ${outputFilePath}...`);
    fs.writeFile(outputFilePath, JSON.stringify(data), (err) => {
        if(err) {
            console.error(err);
            console.log("Error writing to file: ", outputFilePath);
        } else {
            organizeData(data);
        }
    });
}

function organizeData(data) {
    let weights = [];

    let prevMonth = -1;

    let projectPlanning = [];
    let gameDev = [];
    let generalProgramming = [];
    let vfx = [];
    let sfx = [];
    let writing = [];
    let researchAndStudying = [];

    let happiness = [];
    let happyValue = -1;

    let recordDate;
    
    data.forEach(record => {
        if(record.weight !== null) { weights.push({date: record.date, weight: record.weight}); }

        recordDate = new Date(record.date);
        if(recordDate.getMonth() === prevMonth) {
            projectPlanning[prevMonth] += parseInt(record.projectPlanning);
            gameDev[prevMonth] += parseInt(record.gameDev);
            generalProgramming[prevMonth] += parseInt(record.generalProgramming);
            vfx[prevMonth] += parseInt(record.vfx);
            sfx[prevMonth] += parseInt(record.sfx);
            writing[prevMonth] += parseInt(record.writing);
            researchAndStudying[prevMonth] += parseInt(record.researchAndStudying);
        } else {
            prevMonth = recordDate.getMonth();
            projectPlanning.push(parseInt(record.projectPlanning));
            gameDev.push(parseInt(record.gameDev));
            generalProgramming.push(parseInt(record.generalProgramming));
            vfx.push(parseInt(record.vfx));
            sfx.push(parseInt(record.sfx));
            writing.push(parseInt(record.writing));
            researchAndStudying.push(parseInt(record.researchAndStudying));
            happiness.push({one: 0, two: 0, three: 0, four: 0, five:0});
        }

        happyValue = parseInt(record.happiness);
        if(happyValue === 1) {
            happiness[prevMonth].one += 1;
        } else if(happyValue === 2) {
            happiness[prevMonth].two += 1;
        } else if(happyValue === 3) {
            happiness[prevMonth].three += 1;
        } else if(happyValue === 4) {
            happiness[prevMonth].four += 1;
        } else if(happyValue === 5) {
            happiness[prevMonth].five += 1;
        }


    });

    let organized = {
        weights,
        timeSpent: {
            projectPlanning,
            gameDev,
            generalProgramming,
            vfx,
            sfx,
            writing,
            researchAndStudying,  
        },
        happiness
    }

    writeOrganizedToFile(organized);
}

async function writeOrganizedToFile(organized) {
    console.log(`Writing to file: ${organizedFilePath}...`);
    fs.writeFile(organizedFilePath, JSON.stringify(organized), (err) => {
        if(err) {
            console.error(err);
            console.log("Error writing to file: ", organizedFilePath);
        } else {
            console.log("Done!");
        }
    });
}