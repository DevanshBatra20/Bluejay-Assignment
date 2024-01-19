const fs = require("fs");
const xlsx = require("xlsx");

//Specifying excel file path.
const excelFilePath = "./Assignment_Timecard.xlsx";

const logFilePath = "./usersWithShiftOfMoreThan14Hours.txt";

//Reading the excel file.
const workbook = xlsx.readFile(excelFilePath);

//Since the file has only one sheet we are using 0 index to get the sheet.
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

//Converting the sheet to JSON object.
const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 2 });

//Filtering the users with more than 7 days of work.
const usersWithSevenDays = jsonData.filter((user) =>
  hasConsecutiveDays(user, 7)
);

//Filtering the users with 1 to 10 hours of shift.
const usersWithGivenShiftHours = jsonData.filter((user) =>
  usersWithShiftBetweenTimes(user, 1, 10)
);

//Filtering the users with more than 14 hours of shift.
const usersWithMoreThanFourteenHoursShift = jsonData.filter((user) =>
  usersWithShiftMoreThanHours(user, 14)
);

nameAndPoistion = usersWithMoreThanFourteenHoursShift.map((user) => ({
  name: user["Employee Name"],
  position: user["Position Status"],
}));

console.log(nameAndPoistion);

const logString = nameAndPoistion
  .map((obj) => JSON.stringify(obj, null, 2))
  .join("\n\n");
fs.writeFileSync(logFilePath, logString);

console.log(`Log saved to ${logFilePath}`);

//Function to check if the user has consecutive days of work.
function hasConsecutiveDays(user, consecutiveDays) {
  const startDateKey = "Pay Cycle Start Date";
  const endDateKey = "Pay Cycle End Date";

  if (!user[startDateKey] || !user[endDateKey]) {
    return false;
  }

  const startDate = excelDateToJsonDate(user[startDateKey]);
  const endDate = excelDateToJsonDate(user[endDateKey]);

  const dayInMillis = 24 * 60 * 60 * 1000;
  const differenceInDays = (endDate - startDate) / dayInMillis;

  return differenceInDays >= consecutiveDays;
}

//Function to check if the user has shift between given hours.
function usersWithShiftBetweenTimes(user, minHours, maxHours) {
  const startTimeKey = "Time";
  const endTimeKey = "Time Out";

  if (!user[startTimeKey] || !user[endTimeKey]) {
    return false;
  }

  const startTime = excelDateToJsonDate(user[startTimeKey]);
  const endTime = excelDateToJsonDate(user[endTimeKey]);

  const differenceInHours = (endTime - startTime) / (1000 * 60 * 60);

  if (differenceInHours >= minHours && differenceInHours <= maxHours) {
    return true;
  }

  return false;
}

//Function to check if the user has shift more than given hours.
function usersWithShiftMoreThanHours(user, minHours) {
  const startTimeKey = "Time";
  const endTimeKey = "Time Out";

  if (!user[startTimeKey] || !user[endTimeKey]) {
    return false;
  }

  const startTime = excelDateToJsonDate(user[startTimeKey]);
  const endTime = excelDateToJsonDate(user[endTimeKey]);

  const differenceInHours = (endTime - startTime) / (1000 * 60 * 60);

  if (differenceInHours >= minHours) {
    return true;
  }

  return false;
}

//Function to convert excel date to json date.
function excelDateToJsonDate(excelDate) {
  const baseDate = new Date(1899, 11, 30);

  const milliSeconds = (excelDate - 1) * 24 * 60 * 60 * 1000;
  return new Date(baseDate.getTime() + milliSeconds);
}
