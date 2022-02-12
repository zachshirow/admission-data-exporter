const fs = require("fs");
const XLSX = require("xlsx");

//to do
// 1) get the list of files in the directory
// 2) in a map, rename the files to the format Degree - UserID - RequestID
// 3) create another array with the new file names and append each file name that is created to this array
// 4) create a new object or excel workbook where you can add the needed data later on
// 5) loop over the new file names and get the needed fields from them. add the fields based on their degree to three different sheets

const fileNames = fs.readdirSync("./data-import");

const excelFileNames = fileNames.filter((file) => {
	return file.includes(".xlsx");
});

function findCell(sheet, cell) {
	return sheet[cell] ? sheet[cell].v : undefined;
}

const excelFileNamesRevised = [];

// rename the files and push to the array of filenames
excelFileNames.map((f) => {
	// I first need to rename the file as follows: Degree - UserID - RequestID
	// read the file
	const workbook = XLSX.readFile(`./data-import/${f}`);

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
	const degree = findCell(step2Sheet, "D2");
	const userID = findCell(step2Sheet, "B2");
	const requestID = findCell(personalInfoSheet, "D2");

	const newFileName = `./data-import/${degree} - ${userID} - ${requestID}.xlsx`;

	fs.renameSync(`./data-import/${f}`, newFileName);

	excelFileNamesRevised.push(newFileName);
});

const headings = [
	"userID",
	"requestID",
	"lastName",
	"middleName",
	"firstName",
	"firstNameFa",
	"lastNameFa",
	"gender",
	"maritalStatus",
	"spouseName",
	"numberOfChildren",
	"religion",
	"fathersName",
	"fathersNameFa",
	"mothersName",
	"mothersNameFa",
	"birthDate",
	"country",
	"city",
	"nationality1",
	"nationality2",
	"nationality3",
	"passportNumber",
	"passportIssueDate",
	"passportExpiryDate",
	"passportIssuePlace",
	"addressLine1",
	"addressLine2",
	"addressLine3",
	"phone",
	"fax",
	"mobile",
	"email",
	"programId",
	"programTitle",
	"programTitleFa",
	"programDisciplineFa",
	"degree",
	"highSchoolProgramId",
	"highSchoolProgramTitle",
	"highSchoolProgramTitleFa",
	"highSchoolInstitutionId",
	"highSchoolInstitutionTitle",
	"highSchoolInstitutionTitleFa",
	"highSchoolGPA",
	"highSchoolIssueDate",
	"bachelorProgramId",
	"bachelorProgramTitle",
	"bachelorProgramTitleFa",
	"bachelorInstitutionId",
	"bachelorInstitutionTitle",
	"bachelorInstitutionTitleFa",
	"bachelorGPA",
	"bachelorIssueDate",
	"masterProgramId",
	"masterProgramTitle",
	"masterProgramTitleFa",
	"masterInstitutionId",
	"masterInstitutionTitle",
	"masterInstitutionTitleFa",
	"masterGPA",
	"masterIssueDate",
];

const records = [];

//map through the excel files and get the data in the cells
excelFileNamesRevised.map((f) => {
	const workbook = XLSX.readFile(f);

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
	const userID = findCell(personalInfoSheet, "B2");
	const requestID = findCell(personalInfoSheet, "D2");
	const lastName = findCell(personalInfoSheet, "H2");
	const middleName = findCell(personalInfoSheet, "G2");
	const firstName = findCell(personalInfoSheet, "F2");
	const firstNameFa = findCell(personalInfoSheet, "I2");
	const lastNameFa = findCell(personalInfoSheet, "J2");
	const gender = findCell(personalInfoSheet, "P2");
	const maritalStatus = findCell(personalInfoSheet, "Q2");
	const spouseName = findCell(personalInfoSheet, "R2");
	const numberOfChildren = findCell(personalInfoSheet, "S2");
	const religion = findCell(personalInfoSheet, "T2");

	const fathersName = findCell(personalInfoSheet, "K2");
	const fathersNameFa = undefined;
	const mothersName = findCell(personalInfoSheet, "L2");
	const mothersNameFa = undefined;

	const birthDate = findCell(personalInfoSheet, "M2");
	const country = findCell(personalInfoSheet, "N2");
	const city = findCell(personalInfoSheet, "O2");
	const nationality1 = findCell(personalInfoSheet, "Y2");
	const nationality2 = findCell(personalInfoSheet, "Z2");
	const nationality3 = findCell(personalInfoSheet, "AA2");

	const passportNumber = findCell(personalInfoSheet, "U2");
	const passportIssueDate = findCell(personalInfoSheet, "V2");
	const passportExpiryDate = findCell(personalInfoSheet, "W2");
	const passportIssuePlace = findCell(personalInfoSheet, "X2");

	const addressLine1 = findCell(personalInfoSheet, "AB2");
	const addressLine2 = findCell(personalInfoSheet, "AC2");
	const addressLine3 = findCell(personalInfoSheet, "AD2");
	const phone = findCell(personalInfoSheet, "AE2");
	const fax = findCell(personalInfoSheet, "AF2");
	const mobile = findCell(personalInfoSheet, "AG2");
	const email = findCell(personalInfoSheet, "AH2");

	const programId = findCell(step2Sheet, "F2");
	const programTitle = findCell(step2Sheet, "X2");
	const programTitleFa = undefined;
	const programDisciplineFa = undefined;

	const degree = findCell(step2Sheet, "D2");

	const highSchoolProgramId = undefined;
	const highSchoolProgramTitle = findCell(step3Sheet, "G2");
	const highSchoolProgramTitleFa = undefined;
	const highSchoolInstitutionId = undefined;
	const highSchoolInstitutionTitle = findCell(step3Sheet, "E2");
	const highSchoolInstitutionTitleFa = undefined;
	const highSchoolGPA = findCell(step3Sheet, "D2");
	const highSchoolIssueDate = findCell(step3Sheet, "C2");

	const bachelorProgramId = undefined;
	const bachelorProgramTitle = findCell(step3Sheet, "G3");
	const bachelorProgramTitleFa = undefined;
	const bachelorInstitutionId = undefined;
	const bachelorInstitutionTitle = findCell(step3Sheet, "E3");
	const bachelorInstitutionTitleFa = undefined;
	const bachelorGPA = findCell(step3Sheet, "D3");
	const bachelorIssueDate = findCell(step3Sheet, "C3");

	const masterProgramId = undefined;
	const masterProgramTitle = findCell(step3Sheet, "G4");
	const masterProgramTitleFa = undefined;
	const masterInstitutionId = undefined;
	const masterInstitutionTitle = findCell(step3Sheet, "E4");
	const masterInstitutionTitleFa = undefined;
	const masterGPA = findCell(step3Sheet, "D4");
	const masterIssueDate = findCell(step3Sheet, "C4");

	const newRecord = [
		userID,
		requestID,
		lastName,
		middleName,
		firstName,
		firstNameFa,
		lastNameFa,
		gender,
		maritalStatus,
		spouseName,
		numberOfChildren,
		religion,
		fathersName,
		fathersNameFa,
		mothersName,
		mothersNameFa,
		birthDate,
		country,
		city,
		nationality1,
		nationality2,
		nationality3,
		passportNumber,
		passportIssueDate,
		passportExpiryDate,
		passportIssuePlace,
		addressLine1,
		addressLine2,
		addressLine3,
		phone,
		fax,
		mobile,
		email,
		programId,
		programTitle,
		programTitleFa,
		programDisciplineFa,
		degree,
		highSchoolProgramId,
		highSchoolProgramTitle,
		highSchoolProgramTitleFa,
		highSchoolInstitutionId,
		highSchoolInstitutionTitle,
		highSchoolInstitutionTitleFa,
		highSchoolGPA,
		highSchoolIssueDate,
		bachelorProgramId,
		bachelorProgramTitle,
		bachelorProgramTitleFa,
		bachelorInstitutionId,
		bachelorInstitutionTitle,
		bachelorInstitutionTitleFa,
		bachelorGPA,
		bachelorIssueDate,
		masterProgramId,
		masterProgramTitle,
		masterProgramTitleFa,
		masterInstitutionId,
		masterInstitutionTitle,
		masterInstitutionTitleFa,
		masterGPA,
		masterIssueDate,
	];

	records.push(newRecord);
});

// create the workbook
const outputWorkbook = XLSX.utils.book_new();

// to do

// 1. filter the records based on the
const bachelorRecords = records.filter((record) => {
	return record[37] === "Bachelor";
});
const masterRecords = records.filter((record) => {
	return record[37] === "Master";
});
const phdRecords = records.filter((record) => {
	return record[37] === "Phd";
});

// 2. add the heading cells
bachelorRecords.unshift(headings);
masterRecords.unshift(headings);
phdRecords.unshift(headings);
// 3. create the worksheets based on the arrays of data
var bachelorWorksheet = XLSX.utils.aoa_to_sheet(bachelorRecords);
var masterWorksheet = XLSX.utils.aoa_to_sheet(masterRecords);
var phdWorksheet = XLSX.utils.aoa_to_sheet(phdRecords);

// 4. append the worksheets to the workbook
XLSX.utils.book_append_sheet(outputWorkbook, bachelorWorksheet, "bachelor");
XLSX.utils.book_append_sheet(outputWorkbook, masterWorksheet, "master");
XLSX.utils.book_append_sheet(outputWorkbook, phdWorksheet, "phd");

/* Add the worksheet to the workbook */

XLSX.writeFile(outputWorkbook, "./data-export/applications.xlsx");
