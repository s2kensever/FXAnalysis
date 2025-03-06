const fs = require("fs");
const ExcelJS = require("exceljs");

// Đọc dữ liệu từ file JSON
fs.readFile("data.json", "utf8", async (err, jsonData) => {
  if (err) {
    console.error("Lỗi khi đọc file:", err);
    return;
  }

  const data = JSON.parse(jsonData);

  // 1️⃣ Đếm số lần lost theo type
  let lostCount = { N: 0, C: 0 };
  data.forEach((entry) => {
    entry.lost.forEach((lostItem) => {
      lostCount[lostItem.type]++;
    });
  });

  // 2️⃣ Xác suất lost theo khung giờ 1h
  const hourlyDistribution = Array.from({ length: 24 }, () => ({
    total: 0,
    N: 0,
    C: 0,
  }));

  data.forEach((entry) => {
    entry.lost.forEach((lostItem) => {
      const hour = parseInt(lostItem.time.substring(0, 2)); // Giờ Exness
      const realHour = (hour + 7) % 24; // Chuyển sang Giờ Real
      hourlyDistribution[hour].total++;
      hourlyDistribution[hour][lostItem.type]++;
    });
  });

  // 3️⃣ Xác suất lost theo khung giờ 2h
  const twoHourlyDistribution = Array.from({ length: 12 }, () => ({
    total: 0,
    N: 0,
    C: 0,
  }));

  for (let i = 0; i < 24; i++) {
    const index = Math.floor(i / 2);
    twoHourlyDistribution[index].total += hourlyDistribution[i].total;
    twoHourlyDistribution[index].N += hourlyDistribution[i].N;
    twoHourlyDistribution[index].C += hourlyDistribution[i].C;
  }

  // 4️⃣ Đếm số lần lost theo các ngày trong tuần
  const daysOfWeek = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
  ];
  let dailyLost = Array.from({ length: 7 }, () => ({ total: 0, N: 0, C: 0 }));

  data.forEach((entry) => {
    entry.lost.forEach((lostItem) => {
      const dateParts = entry.date.split("/"); // Giả sử định dạng "DD/MM/YYYY"
      const dateObj = new Date(
        `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`
      ); // Chuyển về Date object
      const dayIndex = dateObj.getDay();
      const correctedDayIndex = (dayIndex + 6) % 7; // Đổi 0 (CN) thành 6, 1 (T2) thành 0, ...
      dailyLost[correctedDayIndex].total++;
      dailyLost[correctedDayIndex][lostItem.type]++;
    });
  });

  // **Tạo file Excel**
  const workbook = new ExcelJS.Workbook();

  // Sheet 1: Số lần lost theo type
  const typeSheet = workbook.addWorksheet("Lost by Type");
  typeSheet.addRow(["Type", "Total"]);
  typeSheet.addRow(["N", lostCount.N]);
  typeSheet.addRow(["C", lostCount.C]);

  // Sheet 2: Xác suất lost theo khung giờ 1h
  const hourlySheet = workbook.addWorksheet("Hourly Lost (1h)");
  hourlySheet.addRow(["Range (Exness)", "Range (Real)", "Total", "N", "C"]);
  hourlyDistribution.forEach((item, hour) => {
    const realHour = (hour + 7) % 24;
    hourlySheet.addRow([
      `${hour}-${hour + 1}`,
      `${realHour}-${(realHour + 1) % 24}`,
      item.total,
      item.N,
      item.C,
    ]);
  });

  // Sheet 3: Xác suất lost theo khung giờ 2h
  const twoHourlySheet = workbook.addWorksheet("Hourly Lost (2h)");
  twoHourlySheet.addRow(["Range (Exness)", "Range (Real)", "Total", "N", "C"]);
  twoHourlyDistribution.forEach((item, index) => {
    const exnessStart = index * 2;
    const exnessEnd = exnessStart + 2;
    const realStart = (exnessStart + 7) % 24;
    const realEnd = (exnessEnd + 7) % 24;
    twoHourlySheet.addRow([
      `${exnessStart}-${exnessEnd}`,
      `${realStart}-${realEnd}`,
      item.total,
      item.N,
      item.C,
    ]);
  });

  // Sheet 4: Số lần lost theo ngày trong tuần
  const dailySheet = workbook.addWorksheet("Lost by Day");
  dailySheet.addRow(["Day", "Total", "N", "C"]);
  daysOfWeek.forEach((day, index) => {
    dailySheet.addRow([
      day,
      dailyLost[index].total,
      dailyLost[index].N,
      dailyLost[index].C,
    ]);
  });

  // Lưu file Excel
  const fileName = "Lost_Analysis.xlsx";
  await workbook.xlsx.writeFile(fileName);
  console.log(`✅ Xuất file Excel thành công: ${fileName}`);
});
