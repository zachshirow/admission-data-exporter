const fs = require("fs");
const XLSX = require("xlsx");

//to do
// 1) get the list of files in the directory
// 2) in a map, rename the files to the format Degree - UserID - RequestID
// 3) create another array with the new file names and append each file name that is created to this array
// 4) create a new object or excel workbook where you can add the needed data later on
// 5) loop over the new file names and get the needed fields from them. add the fields based on their degree to three different sheets

const excelFileNames = fs.readdirSync("./excel-data");
const excelFileNamesRevised = [];

// rename the files and push to the array of filenames
excelFileNames.map((f) => {
	// I first need to rename the file as follows: Degree - UserID - RequestID
	// read the file
	const workbook = XLSX.readFile(`./excel-data/${f}`);

	// the available sheets
	const personalInfo = workbook.SheetNames[1];
	const step2 = workbook.SheetNames[2];
	const step3 = workbook.SheetNames[3];
	const step4 = workbook.SheetNames[4];

	const personalInfoSheet = workbook.Sheets[personalInfo];
	const step2Sheet = workbook.Sheets[step2];
	const step3Sheet = workbook.Sheets[step3];
	const step4Sheet = workbook.Sheets[step4];

	// required cells for doing the rename
	const degree = step2Sheet["D2"] ? step2Sheet["D2"].v : undefined;
	const userID = step2Sheet["B2"] ? step2Sheet["B2"].v : undefined;
	const requestID = personalInfoSheet["D2"]
		? personalInfoSheet["D2"].v
		: undefined;

	const newFileName = `./excel-data/${degree} - ${userID} - ${requestID}.xlsx`;

	fs.renameSync(`./excel-data/${f}`, newFileName);

	excelFileNamesRevised.push(newFileName);
});

excelFileNamesRevised.map((f) => {
	const workbook = XLSX.readFile(f);
	console.log(f);
	// the available sheets
	const personalInfo = workbook.SheetNames[1];
	const step2 = workbook.SheetNames[2];
	const step3 = workbook.SheetNames[3];
	const step4 = workbook.SheetNames[4];

	const personalInfoSheet = workbook.Sheets[personalInfo];
	const step2Sheet = workbook.Sheets[step2];
	const step3Sheet = workbook.Sheets[step3];
	const step4Sheet = workbook.Sheets[step4];

	// required cells for doing the rename
	const degree = step2Sheet["D2"] ? step2Sheet["D2"].v : undefined;
	const userID = step2Sheet["B2"] ? step2Sheet["B2"].v : undefined;
	const requestID = personalInfoSheet["D2"]
		? personalInfoSheet["D2"].v
		: undefined;
});
