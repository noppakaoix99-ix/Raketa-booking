function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Raketa - Phuket Motorbike Rental')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getVehicleModels() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('vehicles');
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return values.map(row => row[0]).filter(item => item !== "");
  } catch (e) { return []; }
}

function searchBooking(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('rental logs');
    if (!sheet) return "ERROR: Sheet name 'rental logs' not found";

    // ใช้ getDisplayValues เพื่อให้อ่านค่าได้ตรงกับที่ตาเห็น
    const data = sheet.getDataRange().getDisplayValues(); 
    const searchId = id.toString().trim().toLowerCase();
    
    // วนลูปทุกแถว (เริ่มจากแถวที่ 2)
    for (let i = 1; i < data.length; i++) {
      // ตรวจสอบข้อมูลในแถวนั้นๆ (เน้นคอลัมน์ A แต่ถ้าไม่เจอจะเช็คคอลัมน์อื่นด้วย)
      // rowId คือค่าในคอลัมน์ A (Index 0)
      let rowId = data[i][0].toString().trim().toLowerCase();
      
      // ถ้าคอลัมน์ A ตรง หรือ มีคอลัมน์ใดคอลัมน์หนึ่งในแถวนั้นที่มีรหัสนี้อยู่
      if (rowId === searchId || data[i].includes(id.toString().trim())) {
        return {
          row: i + 1,
          clientName: data[i][2],    // C
          passport: data[i][3],      // D
          nationality: data[i][4],   // E
          model: data[i][5],         // F
          startMileage: data[i][6],  // G
          amount: data[i][11],       // L
          deposit: data[i][12],      // M
          bikeCost: data[i][13],     // N
          bikeId: data[i][14],       // O
          planReturn: data[i][15]    // P
        };
      }
    }
    return null; // หาไม่เจอจริงๆ
  } catch (e) { 
    return "ERROR: " + e.toString(); 
  }
}

function saveBooking(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('rental logs');
    const now = new Date();
    const timestampNow = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");

    if (formData.row) {
      const rowNum = parseInt(formData.row);
      sheet.getRange(rowNum, 9).setValue(timestampNow); // I: Actual Return
      sheet.getRange(rowNum, 10).setValue(formData.returnMileage); // J: Return Mileage
      const startMil = sheet.getRange(rowNum, 7).getValue();
      if (startMil !== "" && formData.returnMileage) {
        sheet.getRange(rowNum, 11).setValue(Number(formData.returnMileage) - Number(startMil)); // K: Total distance
      }
      return "SUCCESS_UPDATE";
    } else {
      const bookingId = "RK-" + now.getTime().toString().slice(-6);
      sheet.appendRow([
        bookingId,            // A
        timestampNow,         // B
        formData.clientName,  // C
        formData.passport,    // D
        formData.nationality, // E
        formData.model,       // F
        formData.startMileage,// G
        timestampNow,         // H (Rent Start)
        "", "", "",           // I, J, K
        formData.amount,      // L
        formData.deposit,     // M
        formData.bikeCost,    // N
        formData.bikeId,      // O
        formData.planReturn   // P
      ]);
      return bookingId;
    }
  } catch (err) { return "ERROR: " + err.toString(); }
}

function getGanttChartData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rentalSheet = ss.getSheetByName('rental logs');
  const vehicleSheet = ss.getSheetByName('vehicles');
  const vData = vehicleSheet.getRange(2, 1, vehicleSheet.getLastRow()-1, 3).getValues();
  const rData = rentalSheet.getDataRange().getValues();
  
  const today = new Date();
  today.setHours(0,0,0,0);
  const dateHeaders = [];
  for(let i=0; i<14; i++) {
    let d = new Date(today);
    d.setDate(today.getDate() + i);
    dateHeaders.push(Utilities.formatDate(d, "GMT+7", "dd/MM"));
  }

  const ganttData = vData.map(vRow => {
    const modelName = vRow[0];
    const days = [];
    for(let i=0; i<14; i++) {
      let checkDate = new Date(today);
      checkDate.setDate(today.getDate() + i);
      let status = 'available';
      for(let j=1; j<rData.length; j++) {
        let rentStart = rData[j][7] ? new Date(rData[j][7]) : null;
        let actualReturn = rData[j][8] ? new Date(rData[j][8]) : null;
        let planReturn = rData[j][15] ? new Date(rData[j][15]) : null;
        let rowBike = rData[j][5]; 

        if(rowBike === modelName && rentStart) {
          rentStart.setHours(0,0,0,0);
          let rentEnd = actualReturn ? new Date(actualReturn) : (planReturn ? new Date(planReturn) : new Date(today.getTime() + 100 * 86400000));
          rentEnd.setHours(0,0,0,0);
          if(checkDate >= rentStart && checkDate <= rentEnd) {
            status = 'rented';
            break;
          }
        }
      }
      days.push(status);
    }
    return { model: modelName, days: days };
  });
  return { headers: dateHeaders, data: ganttData };
}
