const mammoth = require("mammoth");
const ExcelJS = require("exceljs");

async function readDocx(filePath) {
  const result = await mammoth.extractRawText({ path: filePath });
  return result.value; 
}

function processText(text) {
  const blocks = text.split(/\n{3,}/).filter(block => block.trim() !== '');

  const groupedData = {
    chapter: [],
    section: [],
    hadith: []
  };

  const chapterRegex = /^অধ্যায়:\s*(.*)/;
  const hadithRegex = /^\s*\[[০-৯]+\][\s\S]*?/;

  blocks.forEach(block => {
    if (block.match(chapterRegex)) {
      groupedData.chapter.push(block);
    } else if(block.match(hadithRegex)) {
        groupedData.hadith.push(block);
    } else{
        groupedData.section.push(block);
    }
  });

  return groupedData;
}

async function writeToExcel(groupedData) {
    let i = 0, j = 0, k = 0;
    const workbook = new ExcelJS.Workbook();

    const chapterSheet = workbook.addWorksheet('Chapter');
    const sectionSheet = workbook.addWorksheet('Section');
    const hadithSheet = workbook.addWorksheet('Hadith');

    chapterSheet.addRow(['ID', 'Chapter Name']);
    chapterSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center' };
    })

    sectionSheet.addRow(['ID', 'Section Name']);
    sectionSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center' };
    })

    hadithSheet.addRow(['ID', 'Hadith']);
    hadithSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center' };
    })

    groupedData.chapter.forEach(item => {
        i++;
        chapterSheet.addRow([i, item]);
    });

    groupedData.section.forEach(item => {
        j++;
        sectionSheet.addRow([j, item]);
    });

    groupedData.hadith.forEach(item => {
        k++;
        hadithSheet.addRow([k, item]);
    });

    await workbook.xlsx.writeFile('output.xlsx');

    console.log('Excel file created successfully!');
}


async function main() {
    const text = await readDocx("Merged.docx");
    const processedData = processText(text);
    

    writeToExcel(processedData);

}

main();