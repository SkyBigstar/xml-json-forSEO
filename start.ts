import * as XLSX from 'xlsx';

const workbook: XLSX.WorkBook = XLSX.readFile('data/1.xlsx');

const worksheet_roadmap: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];

const data: any[] = XLSX.utils.sheet_to_json(worksheet_roadmap, { header: 1 });

interface Result {
    name: string
    toAdd: number
    tooAddType: string
    difficulty: string
    type: string
    isTop200Factor: boolean
    isHighUsageRate: boolean
    strongFactor: boolean // is green?
    strongFactorExcluding: boolean // is red?
}

interface Phase {
    title: string;
    results: Result[];
}

interface Roadmap_data {
    phases: Phase[];
}

const ret_data: Roadmap_data = {phases: []}; // the return value of this method

let cnt_phases = 0;

// Loop through each row of data
for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    if(!rowData.length) continue;
    // When meet the first line for Phase paragraph
    if (rowData[0].startsWith("Phase")) {
        const _data : Phase = { title: rowData[0], results: [] }
        ret_data.phases.push(_data);
        cnt_phases ++;
        continue;
    }

    if(!cnt_phases) continue;
    
    // Get the cell address
    const cellAddress = XLSX.utils.encode_cell({ r: i, c: 1 })
    // Get the cell style
    const cellStyle = worksheet_roadmap[cellAddress]?.s;
    // If the cell has a style, get the font color
    let _strongFactor: boolean = false;
    let _strongFactorExcluding: boolean = false;

    if (cellStyle !== undefined && cellStyle !== null && cellStyle.font !== undefined && cellStyle.font !== null) {
        const fontColor = cellStyle.font.color;
        // If the font color is not black, add it to the cell data
        if (fontColor && fontColor.rgb && (fontColor.rgb === 'ff0000' || fontColor.rgb === 'FF0000')) {
            _strongFactor = true;
        }
        if (fontColor && fontColor.rgb && (fontColor.rgb === '00ff00' || fontColor.rgb === '00FF00')) {
            _strongFactorExcluding = true;
        }
    }
    
    const _result: Result = {
        name: rowData[0],
        toAdd: parseInt(rowData[1].match(/Add (\d+) more\. (.*)/)[1]),
        tooAddType: rowData[1].match(/Add (\d+) more\. (.*)/)[2],
        difficulty: rowData[2],
        type: rowData[3],
        isTop200Factor: rowData[4] == "Top 200 Factor",
        isHighUsageRate: rowData[5] == "High Usage Rate",
        strongFactor: _strongFactor, // is green?
        strongFactorExcluding: _strongFactorExcluding, // is red?
    }
    ret_data.phases[cnt_phases - 1].results.push(_result);
}

console.log(ret_data);