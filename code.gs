const SPREADSHEET_ID = '1EduuMVZfiXJQkKGB9jgwKxUDVKn0zyCS8jwYgw3EJ2Q'; // ตรวจสอบ ID ของ Google Sheets ให้ถูกต้อง
const FOLDER_ID = '1IH3qRwY9qpJZb2RaX3Ru650RrFfQ6BCk'; 

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const sheetMap = {
      'student': 'Students',
      'teacher': 'Teachers',
      'parent': 'Parents'
    };
    
    // -----------------------------------------
    // ACTION: ตรวจสอบสถานะผู้ใช้เมื่อเข้าหน้าเว็บ
    // -----------------------------------------
    if (action === 'checkUser') {
      const roles = ['student', 'teacher', 'parent'];
      for (let r = 0; r < roles.length; r++) {
        let role = roles[r];
        let sheet = ss.getSheetByName(sheetMap[role]);
        if (!sheet) continue;
        
        let sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) {
          if (sheetData[i][1] === data.userId) { 
            let responseObj = { 
              isRegistered: true, 
              role: role, 
              id: sheetData[i][0].toString().trim(), 
              name: sheetData[i][3] + sheetData[i][4] + " " + sheetData[i][5] // คำนำหน้า + ชื่อ + นามสกุล
            };

            // ✅ เพิ่มการส่งห้องเรียน (คอลัมน์ G / Index 6) กลับไปด้วยสำหรับครูและนักเรียน
            if (role === 'teacher' || role === 'student') {
              responseObj.assignedClass = sheetData[i][6] ? sheetData[i][6].toString().trim() : "";
            }
            return createResponse(responseObj);
          }
        }
      }
      return createResponse({ isRegistered: false });
    }
    
    // -----------------------------------------
    // ACTION: ยืนยันรหัสประจำตัวเพื่อลงทะเบียน
    // -----------------------------------------
    else if (action === 'verifyId') {
      const roles = ['student', 'teacher', 'parent'];
      const targetId = data.systemId.toString().trim();
      
      for (let r = 0; r < roles.length; r++) {
        let role = roles[r];
        let sheet = ss.getSheetByName(sheetMap[role]);
        if (!sheet) continue;
        
        let sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) {
          if (sheetData[i][0].toString().trim() === targetId) { 
            if (sheetData[i][1] && sheetData[i][1] !== "") {
              return createResponse({ success: false, message: 'รหัสนี้ถูกผูกกับ LINE อื่นไปแล้ว!' });
            }
            let fullName = sheetData[i][3] + sheetData[i][4] + " " + sheetData[i][5];
            
            let responseObj = { 
              success: true, 
              role: role, 
              name: fullName,
              systemId: targetId,
              rowNumber: i + 1
            };

            // ✅ เพิ่มการส่งห้องเรียนกลับไปด้วยตอนยืนยันสิทธิ์
            if (role === 'teacher' || role === 'student') {
              responseObj.assignedClass = sheetData[i][6] ? sheetData[i][6].toString().trim() : "";
            }
            return createResponse(responseObj);
          }
        }
      }
      return createResponse({ success: false, message: 'ไม่พบรหัสนี้ในระบบ โปรดตรวจสอบให้ถูกต้อง' });
    }
    
    // -----------------------------------------
    // ACTION: บันทึกข้อมูลผูกบัญชี LINE ID
    // -----------------------------------------
    else if (action === 'registerLine') {
      let sheetName = sheetMap[data.role];
      let sheet = ss.getSheetByName(sheetName);
      
      sheet.getRange(data.rowNumber, 2).setValue(data.userId);      
      sheet.getRange(data.rowNumber, 3).setValue(data.displayName); 
      
      if (data.role === 'student' && data.imageBase64) {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        let base64Data = data.imageBase64;
        let mimeType = data.mimeType || 'image/jpeg';
        if (base64Data.indexOf(';base64,') !== -1) {
          const parts = base64Data.split(';base64,');
          mimeType = parts[0].split(':')[1];
          base64Data = parts[1];
        }
        const decodedData = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedData, mimeType, data.systemId + '_profile_' + new Date().getTime() + '.jpg');
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
        sheet.getRange(data.rowNumber, 8).setValue(file.getUrl());
      }
      return createResponse({ success: true });
    }

    // -----------------------------------------
    // ACTION: บันทึกประวัติการลงชื่อแบบกลุ่ม (สำหรับครู)
    // -----------------------------------------
    // ✅ เปลี่ยนชื่อหรือรองรับ 'bulkAttendance' เพื่อให้ตรงกับหน้าจอ teacher.html
    else if (action === 'bulkAttendance' || action === 'submitBulkAttendance') {
      const attSheet = ss.getSheetByName('Attendance');
      if (!attSheet) return createResponse({ success: false, message: 'ไม่พบชีต Attendance' });
      
      const currentTime = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
      const selectedDate = data.date; 
      
      let attData = attSheet.getDataRange().getValues();
      let rowsToDelete = [];
      let studentIdsToUpdate = data.records.map(r => r.id.toString().trim()); // โค้ดหน้าบ้านส่งมาเป็นคีย์ id หรือ studentId

      for (let i = attData.length - 1; i >= 1; i--) {
        try {
           let rDate = attData[i][4] ? Utilities.formatDate(new Date(attData[i][4]), "Asia/Bangkok", "yyyy-MM-dd") : Utilities.formatDate(new Date(attData[i][0]), "Asia/Bangkok", "yyyy-MM-dd");
           let rId = attData[i][1].toString().trim();
           if (rDate === selectedDate && studentIdsToUpdate.indexOf(rId) !== -1) {
             rowsToDelete.push(i + 1); 
           }
        } catch(e) {}
      }
      // ลบข้อมูลเก่าของวันนั้นออกก่อนเพื่อป้องกันข้อมูลซ้ำซ้อนเวลาลงชื่อซ้ำ
      rowsToDelete.forEach(row => attSheet.deleteRow(row));
      
      data.records.forEach(function(record) {
        attSheet.appendRow([
          currentTime,
          record.id, // รองรับ id จากโค้ดหน้าจอครู
          record.status, 
          'ครูประจำชั้น',
          selectedDate 
        ]);
      });
      return createResponse({ status: 'success', success: true }); // ส่งกลับทั้ง status และ success เพื่อความปลอดภัย
    }

    // -----------------------------------------
    // ACTION: ดึงข้อมูลประวัติการเช็คชื่อ (รองรับทั้ง ครู และ ผู้ปกครอง) 🌟 ปรับปรุงใหม่ 🌟
    // -----------------------------------------
    else if (action === 'getHistory') {
      const pSheet = ss.getSheetByName('Parents');
      const sSheet = ss.getSheetByName('Students');
      const attSheet = ss.getSheetByName('Attendance');
      
      let students = [];
      let records = [];
      let targetStudentIds = [];

      // A. ถ้าเป็นความต้องการของฝั่ง "ครูที่ปรึกษา" (ดูจากตัวแปร type หรือ assignedClass)
      if (data.type === 'teacher' || data.assignedClass) {
        let selectedClass = data.assignedClass;
        if (sSheet) {
          let sRows = sSheet.getDataRange().getValues();
          for (let i = 1; i < sRows.length; i++) {
            let className = sRows[i][6].toString().trim();
            let sId = sRows[i][0].toString().trim();
            // ถ้าเลือกห้องเรียนตรงกัน หรือเลือก 'ทั้งหมด' (ฝ่ายกิจการ)
            if (selectedClass === 'ทั้งหมด' || className === selectedClass) {
              students.push({
                id: sId,
                name: sRows[i][3] + sRows[i][4] + " " + sRows[i][5],
                className: className
              });
              targetStudentIds.push(sId);
            }
          }
        }
      } 
      // B. ถ้าเป็นฝั่ง "ผู้ปกครอง" (ดึงจาก parentId ตามเดิมของคุณ)
      else if (data.parentId) {
        if (pSheet) {
          let pRows = pSheet.getDataRange().getValues();
          for (let i = 1; i < pRows.length; i++) {
            if (pRows[i][0].toString().trim() === data.parentId.toString().trim()) {
              let childIdRaw = pRows[i][6] ? pRows[i][6].toString().trim() : "";
              if (childIdRaw) {
                let ids = childIdRaw.split(',').map(id => id.trim());
                ids.forEach(id => { if(id !== "") targetStudentIds.push(id); });
              }
            }
          }
        }
        if (sSheet && targetStudentIds.length > 0) {
          let sRows = sSheet.getDataRange().getValues();
          for (let i = 1; i < sRows.length; i++) {
            let sId = sRows[i][0].toString().trim();
            if (targetStudentIds.indexOf(sId) !== -1) {
              students.push({
                id: sId,
                name: sRows[i][3] + sRows[i][4] + " " + sRows[i][5],
                className: sRows[i][6]
              });
            }
          }
        }
      }

      // ดึงประวัติการเช็คชื่อตามรายชื่อนักเรียน (ใช้ร่วมกันได้ทั้งครูและผู้ปกครอง)
      if (attSheet && targetStudentIds.length > 0) {
        let attRows = attSheet.getDataRange().getValues();
        let startDate = new Date(data.startDate);
        let endDate = new Date(data.endDate);

        for (let i = 1; i < attRows.length; i++) {
          try {
            let rawDate = attRows[i][4] ? attRows[i][4] : attRows[i][0];
            let recDate = new Date(rawDate);
            let recDateStr = Utilities.formatDate(recDate, "Asia/Bangkok", "yyyy-MM-dd");
            let currentSId = attRows[i][1].toString().trim();

            if (recDate >= startDate && recDate <= endDate && targetStudentIds.indexOf(currentSId) !== -1) {
              records.push({
                studentId: currentSId,
                date: recDateStr,
                status: attRows[i][2].toString().trim()
              });
            }
          } catch(err) { continue; }
        }
      }
      
      // ดึงข้อมูลผู้อำนวยการและชื่อโรงเรียนจาก Setting นำส่งหน้าพิมพ์ใบเช็คชื่อ
      let schoolName = 'โรงเรียนแม่อ้อวิทยาคม';
      let directorName = '.......................';
      const settingsSheet = ss.getSheetByName('Settings');
      if (settingsSheet) {
        const sData = settingsSheet.getDataRange().getValues();
        for (let i = 0; i < sData.length; i++) {
          if (sData[i][0] === 'SCHOOL_NAME') schoolName = sData[i][1];
          if (sData[i][0] === 'DIRECTOR_NAME') directorName = sData[i][1];
        }
      }

      return createResponse({ 
        success: true, 
        students: students, 
        records: records, 
        sysConf: { SCHOOL_NAME: schoolName, DIRECTOR_NAME: directorName } 
      });
    }

    // -----------------------------------------
    // ACTIONอื่น ๆ คงเดิมตามโค้ดของคุณ...
    // -----------------------------------------
    else if (action === 'getSettingsAndUser') {
      // (ส่วนนี้คงโค้ดเดิมของคุณไว้ได้เลย เพราะใช้สำหรับหน้า Check-In พิกัดของนักเรียน)
      const settingsSheet = ss.getSheetByName('Settings');
      let config = { schoolLat: 19.677237, schoolLng: 99.85024, maxDist: 100 }; 
      if (settingsSheet) {
        const sData = settingsSheet.getDataRange().getValues();
        for (let i = 0; i < sData.length; i++) {
          if (sData[i][0] === 'SCHOOL_LAT') config.schoolLat = parseFloat(sData[i][1]);
          if (sData[i][0] === 'SCHOOL_LNG') config.schoolLng = parseFloat(sData[i][1]);
          if (sData[i][0] === 'MAX_DIST') config.maxDist = parseFloat(sData[i][1]);
        }
      }

      let profileData = null;
      const sSheet = ss.getSheetByName('Students');
      const tSheet = ss.getSheetByName('Teachers');
      const pSheet = ss.getSheetByName('Parents');

      if (sSheet) {
        let dataRows = sSheet.getDataRange().getValues();
        for (let i = 1; i < dataRows.length; i++) {
          if (dataRows[i][1] === data.userId) {
            profileData = { role: 'student', id: dataRows[i][0], prefix: dataRows[i][3], name: dataRows[i][4], surname: dataRows[i][5], className: dataRows[i][6] }; break;
          }
        }
      }
      if (!profileData && tSheet) {
        let dataRows = tSheet.getDataRange().getValues();
        for (let i = 1; i < dataRows.length; i++) {
          if (dataRows[i][1] === data.userId) {
            profileData = { role: 'teacher', id: dataRows[i][0], prefix: dataRows[i][3], name: dataRows[i][4], surname: dataRows[i][5], assignedClass: dataRows[i][6] }; break; // แนบ assignedClass
          }
        }
      }
      if (!profileData && pSheet) {
        let dataRows = pSheet.getDataRange().getValues();
        for (let i = 1; i < dataRows.length; i++) {
          if (dataRows[i][1] === data.userId) {
            let childId = dataRows[i][6];
            let childInfo = { prefix: '', name: 'ไม่พบข้อมูลเด็ก', surname: '', className: '-' };
            if (sSheet && childId) {
              let sRows = sSheet.getDataRange().getValues();
              for (let j = 1; j < sRows.length; j++) {
                if (sRows[j][0].toString().trim() === childId.toString().trim()) {
                  childInfo = { prefix: sRows[j][3], name: sRows[j][4], surname: sRows[j][5], className: sRows[j][6] }; break;
                }
              }
            }
            profileData = { role: 'parent', id: dataRows[i][0], prefix: dataRows[i][3], name: dataRows[i][4], surname: dataRows[i][5], child: childInfo }; break;
          }
        }
      }

      let hasCheckedInToday = false;
      const attSheet = ss.getSheetByName('Attendance');
      if (attSheet && profileData) {
        let todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
        let attData = attSheet.getDataRange().getValues();
        for (let i = attData.length - 1; i >= 1; i--) {
          try {
            let recordDate = Utilities.formatDate(new Date(attData[i][0]), "Asia/Bangkok", "yyyy-MM-dd");
            let recordId = attData[i][1].toString().trim();
            if (recordDate === todayStr && recordId === profileData.id.toString().trim()) {
              hasCheckedInToday = true;
              break;
            }
          } catch(e) { continue; }
        }
      }

      if (profileData) {
        return createResponse({ success: true, settings: config, user: profileData, hasCheckedInToday: hasCheckedInToday });
      } else {
        return createResponse({ success: false, message: 'ไม่พบสิทธิ์ผู้ใช้งาน' });
      }
    }
    else if (action === 'checkIn') {
      const sheet = ss.getSheetByName('Attendance');
      if (!sheet) return createResponse({ success: false, message: 'ไม่พบหน้าชีต Attendance' });
      
      const currentTime = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
      let recordStatus = data.status ? data.status : 'มาเรียน';

      sheet.appendRow([
        currentTime, 
        data.roleId, 
        recordStatus, 
        data.roleText
      ]);
      return createResponse({ success: true });
    }
    else if (action === 'getAttendanceSummary') {
      const sSheet = ss.getSheetByName('Students');
      const attSheet = ss.getSheetByName('Attendance');
      
      let students = sSheet ? sSheet.getDataRange().getValues() : [];
      let attendance = attSheet ? attSheet.getDataRange().getValues() : [];
      let todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
      
      let todayAttendanceMap = {};
      for (let i = 1; i < attendance.length; i++) {
        let attDate = Utilities.formatDate(new Date(attendance[i][0]), "Asia/Bangkok", "yyyy-MM-dd");
        if (attDate === todayStr) {
          todayAttendanceMap[attendance[i][1]] = attendance[i][2]; 
        }
      }
      
      let levels = ['ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'];
      let summaryData = {};
      levels.forEach(lvl => {
        summaryData[lvl] = { totalM: 0, totalF: 0, presentM: 0, presentF: 0, absentM: 0, absentF: 0 };
      });
      
      for (let i = 1; i < students.length; i++) {
        let currentLvl = students[i][6].toString().trim();
        let gender = students[i][2].toString().trim();
        let sId = students[i][0];
        
        if (summaryData[currentLvl]) {
          let status = todayAttendanceMap[sId] || 'ขาด'; 
          
          if (gender === 'ชาย') {
            summaryData[currentLvl].totalM++;
            if (status === 'มาเรียน') summaryData[currentLvl].presentM++;
            else summaryData[currentLvl].absentM++; 
          } else {
            summaryData[currentLvl].totalF++;
            if (status === 'มาเรียน') summaryData[currentLvl].presentF++;
            else summaryData[currentLvl].absentF++;
          }
        }
      }
      
      return createResponse({ success: true, dateText: todayStr, summary: summaryData });
    }

  } catch (error) {
    return createResponse({ success: false, message: 'ข้อผิดพลาดของระบบ: ' + error.message });
  }
}

function createResponse(responseObject) {
  return ContentService.createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}
