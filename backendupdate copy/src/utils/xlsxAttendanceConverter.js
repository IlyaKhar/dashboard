import * as XLSX from "xlsx";

export default async function loadAttendance() {
   try {
      const response = await fetch("/attendance.xlsx");
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const rawData = XLSX.utils.sheet_to_json(worksheet, {
         header: 1,
         defval: "",
         raw: false,
      });

      let currentDepartment = "";
      let currentGroup = "";
      let currentStudent = "";
      const records = [];

      for (const row of rawData) {
         const firstCell = row[0];
         const hours = row[5];

         if (
            firstCell !== undefined &&
            firstCell !== null &&
            firstCell !== ""
         ) {
            const cellStr = String(firstCell).trim();

            if (cellStr.startsWith("Отделение")) {
               currentDepartment = cellStr;
               currentGroup = "";
               currentStudent = "";
            } else if (
               cellStr.length <= 10 &&
               /^\d/.test(cellStr) &&
               cellStr.match(/[а-яa-z]/i)
            ) {
               currentGroup = cellStr.toLowerCase();
               currentStudent = "";
            } else if (
               cellStr.split(/\s+/).filter((w) => w.length > 1).length === 3
            ) {
               currentStudent = cellStr;
            } else {
               let parsedDate = null;

               const dateMatch = cellStr.match(
                  /^(\d{1,2})[./](\d{1,2})[./](\d{4})/,
               );
               if (dateMatch) {
                  const [_, day, month, year] = dateMatch;
                  parsedDate = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
               }

               if (
                  parsedDate &&
                  hours !== undefined &&
                  hours !== "" &&
                  currentDepartment &&
                  currentGroup &&
                  currentStudent
               ) {
                  records.push({
                     department: currentDepartment,
                     group: currentGroup,
                     student: currentStudent,
                     date: parsedDate,
                     missed: parseInt(hours) || 0,
                  });
               }
            }
         }
      }

      const departments = {};

      for (const r of records) {
         const dep = r.department;
         const grp = r.group;
         const stu = r.student;

         if (!departments[dep]) {
            departments[dep] = {
               department: dep,
               groups: [],
            };
         }

         let groupObj = departments[dep].groups.find((g) => g.group === grp);
         if (!groupObj) {
            groupObj = {
               group: grp,
               students: [],
            };
            departments[dep].groups.push(groupObj);
         }

         let studentObj = groupObj.students.find((s) => s.student === stu);
         if (!studentObj) {
            studentObj = {
               student: stu,
               attendance: [],
            };
            groupObj.students.push(studentObj);
         }

         studentObj.attendance.push({
            date: r.date,
            missed: r.missed,
         });
      }

      const result = Object.values(departments);

      return result;
   } catch (error) {
      console.error("Ошибка загрузки:", error);
      return [];
   }
}
