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
	"شماره درخواست",
	"نام خانوادگی انگلیسی",
	"نام انگلیسی",
	"نام فارسی",
	"نام خانوادگی فارسی",
	"جنسیت",
	"وضعیت تاهل",
	"دین",
	"مذهب",
	"نام انگلیسی پدر",
	"نام فارسی پدر",
	"تاریخ تولد",
	"محل تولد",
	"تابعیت",
	"شماره گذرنامه",
	"آدرس",
	"موبایل",
	"ایمیل",
	"نوع پذیرش",
	"نوع ورود به آموزش عالی",
	"دوره",
	"شماره رشته",
	"عنوان رشته",
	"گرایش رشته",
	"دانشکده",
	"مقطع",
	"مقطع قبلی",
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
	const requestID = findCell(personalInfoSheet, "D2");
	const lastName = findCell(personalInfoSheet, "H2");
	const firstName = findCell(personalInfoSheet, "F2");
	const middleName =
		findCell(personalInfoSheet, "G2") == undefined
			? ""
			: findCell(personalInfoSheet, "G2");
	const name = `${firstName} ${middleName}`;
	const firstNameFa = findCell(personalInfoSheet, "I2");
	const lastNameFa = findCell(personalInfoSheet, "J2");

	let gender = findCell(personalInfoSheet, "P2");

	switch (gender) {
		case "Male":
			gender = "مرد";
		case "Female":
			gender = "زن";
	}

	let maritalStatus = findCell(personalInfoSheet, "Q2");

	switch (maritalStatus) {
		case "Married":
			maritalStatus = "متاهل";
		case "Single":
			maritalStatus = "مجرد";
	}

	const religion = findCell(personalInfoSheet, "T2");
	const mazhab = undefined;

	const fathersName = findCell(personalInfoSheet, "K2");
	const fathersNameFa = undefined;

	const birthDate = findCell(personalInfoSheet, "M2");
	const country = findCell(personalInfoSheet, "N2");
	const nationality = findCell(personalInfoSheet, "Y2");

	const passportNumber = findCell(personalInfoSheet, "U2");

	const addressLine1 =
		findCell(personalInfoSheet, "AB2") == undefined
			? ""
			: findCell(personalInfoSheet, "AB2");
	const addressLine2 =
		findCell(personalInfoSheet, "AC2") == undefined
			? ""
			: findCell(personalInfoSheet, "AC2");
	const addressLine3 =
		findCell(personalInfoSheet, "AD2") == undefined
			? ""
			: findCell(personalInfoSheet, "AD2");

	const address = `${addressLine1} ${addressLine2} ${addressLine3} `;
	const mobile =
		findCell(personalInfoSheet, "AG2") || findCell(personalInfoSheet, "AE2");
	const email = findCell(personalInfoSheet, "AH2");

	const applicationType = "غیربورسیه";
	const entranceType = "غیربورسیه";
	const programType = "غیربورسیه";

	const programId = findCell(step2Sheet, "F2");
	const programTitleFa = undefined;
	const programDisciplineFa = undefined;
	const faculty = undefined;

	let degree = findCell(step2Sheet, "D2");

	let previousDegree = undefined;

	switch (degree) {
		case "Bachelor":
			degree = "کارشناسی";
			previousDegree = "دبیرستان";
		case "Master":
			degree = "کارشناسی ارشد";
			previousDegree = "کارشناسی";
		case "Phd":
			degree = "دکتری";
			previousDegree = "کارشناسی ارشد";
	}

	const newRecord = [
		requestID,
		lastName,
		name,
		firstNameFa,
		lastNameFa,
		gender,
		maritalStatus,
		religion,
		mazhab,
		fathersName,
		fathersNameFa,
		birthDate,
		country,
		nationality,
		passportNumber,
		address,
		mobile,
		email,
		applicationType,
		entranceType,
		programType,
		programId,
		programTitleFa,
		programDisciplineFa,
		faculty,
		degree,
		previousDegree,
	];

	records.push(newRecord);
});

// create the workbook
const outputWorkbook = XLSX.utils.book_new();

// to do

// 1. filter the records based on the
const bachelorRecords = records.filter((record) => {
	return record[25] == "کارشناسی";
});
const masterRecords = records.filter((record) => {
	return record[25] == "کارشناسی ارشد";
});
const phdRecords = records.filter((record) => {
	return record[25] == "دکتری";
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
