
document.getElementById('selectTarget').addEventListener('click', function () {
    let fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx, .xls';
    fileInput.onchange = e => {
        let file = e.target.files[0];
        let reader = new FileReader();
        reader.onload = function (event) {
            let data = new Uint8Array(event.target.result);
            let workbook = XLSX.read(data, { type: 'array' });
            processWorkbook(workbook);
        };
        reader.readAsArrayBuffer(file);
    };
    fileInput.click();
});

function processWorkbook(workbook) {
    let numberOfSplits = parseInt(document.getElementById('splitParts').value);
    let firstSheetName = workbook.SheetNames[0];
    let worksheet = workbook.Sheets[firstSheetName];
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    let rowsPerSplit = Math.ceil(jsonData.length / numberOfSplits);

    for (let i = 0; i < numberOfSplits; i++) {
        let splitData = jsonData.slice(i * rowsPerSplit, (i + 1) * rowsPerSplit);
        let newWorkbook = XLSX.utils.book_new();
        let newWorksheet = XLSX.utils.aoa_to_sheet(splitData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, `Split_${i + 1}`);
        XLSX.writeFile(newWorkbook, `Split_${i + 1}.xlsx`);
    }

    updateProgressBar(100);
}

document.getElementById('splitFile').addEventListener('click', function () {
    processWorkbook(globalWorkbook);
});

function updateProgressBar(percentage) {
    let progressBar = document.getElementById('progressBar');
    document.getElementById('progressBarContainer').style.display = 'block';
    progressBar.style.width = percentage + '%';
}
