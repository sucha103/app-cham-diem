/**
 * Google Apps Script - Web App Proxy & Sheet Logger
 * Dự án: Mini App Chấm Bài 28 Ngày
 * Phiên bản: V4 (Hỗ trợ nhiều người dùng, Đồng bộ đám mây & Nhớ lịch sử học viên)
 */

// Tên các tab dữ liệu
const TAB_NAME = "Dữ liệu chấm bài";
const ACCOUNTS_TAB = "Tài khoản";
const STUDENTS_TAB = "Danh sách học viên";

// Tiêu đề các cột dữ liệu
const SHEET_HEADERS = ["Thời gian chấm", "Người chấm", "Mã học viên", "Tên học viên", "Ngày học", "Tên ngày", "Bài làm", "Nhận xét AI"];
const ACCOUNTS_HEADERS = ["Thời gian tạo", "Tên đăng nhập", "Mật khẩu", "Họ và tên", "Cấu hình cá nhân"];
const STUDENTS_HEADERS = ["Thời gian tạo", "Mã học viên", "Tên học viên", "AI xưng là", "Học viên xưng là", "Thông tin thêm / Ghi chú"];

/**
 * Khởi tạo tự động các tab dữ liệu nếu chưa tồn tại
 */
function ensureTabsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Script chưa được gắn (bind) với Google Sheet. Vui lòng mở Apps Script từ Google Sheet của bạn.");
  }

  // 1. Tab Dữ liệu chấm bài
  let sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(SHEET_HEADERS);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#F3F4F6")
      .setHorizontalAlignment("center");
  } else {
    // Tự động nâng cấp cấu trúc nếu là phiên bản cũ (7 cột) sang phiên bản mới (8 cột có cột Người chấm)
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders.length === 7) {
      sheet.insertColumnBefore(2);
      sheet.getRange(1, 2).setValue("Người chấm");
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
        .setFontWeight("bold")
        .setBackground("#F3F4F6")
        .setHorizontalAlignment("center");
    }
  }

  // 2. Tab Tài khoản
  let accSheet = ss.getSheetByName(ACCOUNTS_TAB);
  if (!accSheet) {
    accSheet = ss.insertSheet(ACCOUNTS_TAB);
    accSheet.appendRow(ACCOUNTS_HEADERS);
    accSheet.getRange(1, 1, 1, ACCOUNTS_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#F3F4F6")
      .setHorizontalAlignment("center");
    
    // Tạo tài khoản admin mặc định
    accSheet.appendRow([new Date(), "admin", "123456", "Giảng viên Admin", "{}"]);
  }

  // 3. Tab Danh sách học viên
  let studSheet = ss.getSheetByName(STUDENTS_TAB);
  if (!studSheet) {
    studSheet = ss.insertSheet(STUDENTS_TAB);
    studSheet.appendRow(STUDENTS_HEADERS);
    studSheet.getRange(1, 1, 1, STUDENTS_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#F3F4F6")
      .setHorizontalAlignment("center");
  }
}

/**
 * Xử lý yêu cầu GET
 */
function doGet(e) {
  try {
    ensureTabsExist();
  } catch(_) {}
  return ContentService.createTextOutput("Mini App Chấm Bài 28 Ngày - Apps Script Web App V4 đang hoạt động!")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Xử lý yêu cầu POST gửi từ Frontend
 */
function doPost(e) {
  let sheetUrl = "";
  let sheetName = "";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      sheetUrl = ss.getUrl();
      sheetName = ss.getName();
    }
  } catch (_) {}

  try {
    ensureTabsExist();
    
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Không tìm thấy dữ liệu yêu cầu.");
    }
    
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || "grade_and_save";
    const username = payload.username || "";

    // 1. XỬ LÝ ĐĂNG NHẬP
    if (action === "login") {
      const loginUser = payload.loginUsername || "";
      const loginPass = payload.loginPassword || "";
      
      if (!loginUser || !loginPass) {
        throw new Error("Thiếu thông tin đăng nhập.");
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const accSheet = ss.getSheetByName(ACCOUNTS_TAB);
      const data = accSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] == loginUser && data[i][2] == loginPass) {
          const fullname = data[i][3];
          const configJson = data[i][4];
          const students = getStudentList();
          return createJsonResponse({
            success: true,
            fullname: fullname,
            config: JSON.parse(configJson || "{}"),
            students: students,
            sheetUrl: sheetUrl,
            sheetName: sheetName
          });
        }
      }
      return createJsonResponse({ success: false, error: "Tên đăng nhập hoặc mật khẩu không chính xác." });
    }

    // 2. XỬ LÝ ĐĂNG KÝ
    if (action === "register") {
      const regUser = payload.regUsername || "";
      const regPass = payload.regPassword || "";
      const regName = payload.regFullname || "";
      
      if (!regUser || !regPass || !regName) {
        throw new Error("Thiếu thông tin đăng ký.");
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const accSheet = ss.getSheetByName(ACCOUNTS_TAB);
      const data = accSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] == regUser) {
          return createJsonResponse({ success: false, error: "Tên đăng nhập đã tồn tại trên hệ thống." });
        }
      }
      
      accSheet.appendRow([new Date(), regUser, regPass, regName, "{}"]);
      return createJsonResponse({ success: true, msg: "Đăng ký tài khoản thành công!" });
    }

    // 3. XỬ LÝ LƯU CẤU HÌNH CÁ NHÂN
    if (action === "save_user_config") {
      const configUser = payload.configUsername || "";
      const userConfig = payload.userConfig || {};
      
      if (!configUser) {
        throw new Error("Thiếu thông tin người dùng để lưu cấu hình.");
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const accSheet = ss.getSheetByName(ACCOUNTS_TAB);
      const data = accSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] == configUser) {
          accSheet.getRange(i + 1, 5).setValue(JSON.stringify(userConfig));
          return createJsonResponse({ success: true, msg: "Lưu cấu hình cá nhân thành công!" });
        }
      }
      throw new Error("Không tìm thấy thông tin tài khoản người dùng.");
    }

    // 4. LẤY DANH SÁCH HỌC VIÊN
    if (action === "get_students") {
      const students = getStudentList();
      return createJsonResponse({ success: true, students: students });
    }

    // 5. TẠO HỒ SƠ HỌC VIÊN MỚI
    if (action === "create_student") {
      const sId = payload.studentId || "";
      const sName = payload.studentName || "";
      const tPron = payload.teacherPronoun || "chị";
      const sPron = payload.studentPronoun || "em";
      const sNotes = payload.notes || "";
      
      if (!sName) {
        throw new Error("Tên học viên không được để trống.");
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const studSheet = ss.getSheetByName(STUDENTS_TAB);
      const data = studSheet.getDataRange().getValues();
      
      // Kiểm tra xem đã có học viên này chưa (nếu khớp cả ID hoặc Tên)
      let existIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if ((sId && data[i][1] == sId) || (data[i][2] == sName)) {
          existIndex = i;
          break;
        }
      }
      
      if (existIndex !== -1) {
        // Cập nhật thông tin học viên cũ
        studSheet.getRange(existIndex + 1, 2).setValue(sId);
        studSheet.getRange(existIndex + 1, 3).setValue(sName);
        studSheet.getRange(existIndex + 1, 4).setValue(tPron);
        studSheet.getRange(existIndex + 1, 5).setValue(sPron);
        studSheet.getRange(existIndex + 1, 6).setValue(sNotes);
      } else {
        // Tạo dòng mới
        studSheet.appendRow([new Date(), sId, sName, tPron, sPron, sNotes]);
      }
      
      const students = getStudentList();
      return createJsonResponse({ success: true, msg: "Lưu hồ sơ học viên thành công!", students: students });
    }

    // 5.5. LẤY CHI TIẾT LỊCH SỬ KÝ ỨC CỦA HỌC VIÊN
    if (action === "get_student_history") {
      const sName = payload.studentName || "";
      if (!sName) {
        throw new Error("Thiếu Tên học viên để lấy lịch sử.");
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(TAB_NAME);
      if (!sheet) {
        return createJsonResponse({ success: true, history: [] });
      }
      
      const data = sheet.getDataRange().getValues();
      const historyList = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][3] == sName) {
          historyList.push({
            timestamp: data[i][0],
            teacherUsername: data[i][1],
            studentId: data[i][2],
            studentName: data[i][3],
            day: data[i][4],
            dayTitle: data[i][5],
            essay: data[i][6],
            aiReview: data[i][7]
          });
        }
      }
      historyList.sort((a, b) => parseInt(a.day) - parseInt(b.day));
      return createJsonResponse({ success: true, history: historyList });
    }

    // Lấy thông tin đầu vào phục vụ chấm bài/sửa bài
    const apiKey = payload.apiKey;
    const studentId = payload.studentId || "";
    const studentName = payload.studentName || "";
    const day = payload.day || "";
    const dayTitle = payload.dayTitle || "";
    const promptQuestion = payload.promptQuestion || "";
    const essay = payload.essay || "";
    const aiReviewPreGenerated = payload.aiReview;
    const clientSystemPrompt = payload.systemPrompt;
    const feedback = payload.feedback || "";
    const teacherPronoun = payload.teacherPronoun || "chị";
    const studentPronoun = payload.studentPronoun || "em";
    const notes = payload.notes || "";

    // 6. HÀNH ĐỘNG AI VIẾT LẠI NHẬN XÉT (REWRITE)
    if (action === "rewrite_review") {
      if (!apiKey) {
        return createJsonResponse({ success: false, errorType: "CONFIG_ERROR", error: "Thiếu Gemma API Key. Vui lòng cấu hình trong app.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!studentName || !day) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Thiếu Tên học viên hoặc Ngày học để viết lại.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!feedback) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Vui lòng nhập phản hồi/góp ý của bạn.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!aiReviewPreGenerated) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Không tìm thấy nhận xét cũ để chỉnh sửa.", sheetUrl: sheetUrl, sheetName: sheetName });
      }

      // Lấy lịch sử cũ làm ngữ cảnh khi AI viết lại
      const historyContext = getStudentHistoryContext(studentName);
      
      let newReview = "";
      try {
        newReview = callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt, aiReviewPreGenerated, feedback, teacherPronoun, studentPronoun, historyContext, notes);
      } catch (aiError) {
        return createJsonResponse({
          success: false,
          errorType: "AI_ERROR",
          error: "Lỗi kết nối Gemma AI khi viết lại: " + aiError.message,
          sheetUrl: sheetUrl,
          sheetName: sheetName
        });
      }

      let sheetSyncFailed = false;
      let sheetErrorMsg = "";
      try {
        updateSheetReview(studentName, day, newReview);
      } catch (sheetError) {
        sheetSyncFailed = true;
        sheetErrorMsg = sheetError.message;
      }

      return createJsonResponse({
        success: true,
        aiReview: newReview,
        sheetSyncFailed: sheetSyncFailed,
        sheetError: sheetSyncFailed ? "Cập nhật Sheet thất bại: " + sheetErrorMsg : null,
        sheetUrl: sheetUrl,
        sheetName: sheetName
      });
    }

    // 7. HÀNH ĐỘNG CHỈ LƯU (SYNC QUEUE)
    if (action === "save_only") {
      if (!aiReviewPreGenerated) {
        throw new Error("Không có nhận xét AI để lưu.");
      }
      saveToSheet(username, studentId, studentName, day, dayTitle, essay, aiReviewPreGenerated);
      return createJsonResponse({ success: true, msg: "Đã lưu vào Google Sheet thành công.", sheetUrl: sheetUrl, sheetName: sheetName });
    }

    // 8. HÀNH ĐỘNG CHẤM BÀI & TỰ ĐỘNG LƯU (MẶC ĐỊNH)
    if (!apiKey) {
      return createJsonResponse({ success: false, errorType: "CONFIG_ERROR", error: "Thiếu Gemma API Key. Vui lòng cấu hình trong app.", sheetUrl: sheetUrl, sheetName: sheetName });
    }
    if (!studentName) {
      return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Tên học viên không được để trống.", sheetUrl: sheetUrl, sheetName: sheetName });
    }
    if (!essay) {
      return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Nội dung bài làm không được để trống.", sheetUrl: sheetUrl, sheetName: sheetName });
    }

    // Lấy lịch sử cũ làm ngữ cảnh khi AI chấm bài mới
    const historyContext = getStudentHistoryContext(studentName);

    let aiReview = "";
    try {
      aiReview = callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt, null, null, teacherPronoun, studentPronoun, historyContext, notes);
    } catch (aiError) {
      return createJsonResponse({
        success: false,
        errorType: "AI_ERROR",
        error: "Lỗi kết nối Gemma AI: " + aiError.message,
        sheetUrl: sheetUrl,
        sheetName: sheetName
      });
    }

    let sheetSyncFailed = false;
    let sheetErrorMsg = "";
    try {
      saveToSheet(username, studentId, studentName, day, dayTitle, essay, aiReview);
    } catch (sheetError) {
      sheetSyncFailed = true;
      sheetErrorMsg = sheetError.message;
    }

    return createJsonResponse({
      success: true,
      aiReview: aiReview,
      sheetSyncFailed: sheetSyncFailed,
      sheetError: sheetSyncFailed ? "Ghi Sheet thất bại: " + sheetErrorMsg : null,
      sheetUrl: sheetUrl,
      sheetName: sheetName
    });

  } catch (err) {
    return createJsonResponse({
      success: false,
      errorType: "SYSTEM_ERROR",
      error: "Lỗi hệ thống Apps Script: " + err.message,
      sheetUrl: sheetUrl,
      sheetName: sheetName
    });
  }
}

/**
 * Lấy danh sách học viên hiện có từ tab "Danh sách học viên"
 */
function getStudentList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    list.push({
      studentId: data[i][1] || "",
      studentName: data[i][2] || "",
      teacherPronoun: data[i][3] || "chị",
      studentPronoun: data[i][4] || "em",
      notes: data[i][5] || ""
    });
  }
  return list;
}

/**
 * Trích xuất toàn bộ lịch sử viết bài phản tư và nhận xét cũ của học viên
 */
function getStudentHistoryContext(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) return "";
  
  const data = sheet.getDataRange().getValues();
  const history = [];
  
  for (let i = 1; i < data.length; i++) {
    // Cột D là Tên học viên (index 3), Cột E là Ngày học (index 4)
    if (data[i][3] == studentName) {
      history.push({
        day: data[i][4],
        dayTitle: data[i][5],
        essay: data[i][6],
        aiReview: data[i][7]
      });
    }
  }
  
  if (history.length === 0) return "";
  
  // Sắp xếp lịch sử theo ngày học tăng dần
  history.sort((a, b) => parseInt(a.day) - parseInt(b.day));
  
  let contextStr = "\n=== LỊCH SỬ PHẢN TƯ VÀ LỜI PHÊ CŨ CỦA HỌC VIÊN NÀY (ĐỂ ĐẢM BẢO TÍNH NHẤT QUÁN): ===\n";
  history.forEach(item => {
    contextStr += `Ngày thứ ${item.day} (${item.dayTitle}):\n`;
    contextStr += `- Bài làm phản tư: "${item.essay.trim()}"\n`;
    contextStr += `- Nhận xét của giảng viên (bạn): "${item.aiReview.trim()}"\n\n`;
  });
  contextStr += "=== HẾT PHẦN LỊCH SỬ HỌC TẬP ===\n\n";
  return contextStr;
}

/**
 * Helper tạo phản hồi JSON
 */
function createJsonResponse(data) {
  const jsonString = JSON.stringify(data);
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gọi Google AI Studio API để chấm bài viết bằng model gemma-4-26b-a4b-it
 */
function callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt, previousReview, feedback, teacherPronoun, studentPronoun, studentHistoryContext, studentNotes) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemma-4-26b-a4b-it:generateContent?key=" + apiKey;
  
  const tPronoun = (teacherPronoun && teacherPronoun.trim() !== "") ? teacherPronoun.trim() : "chị";
  const sPronoun = (studentPronoun && studentPronoun.trim() !== "") ? studentPronoun.trim() : "em";

  // Xây dựng system instruction mặc định
  const defaultSystemInstruction = 
    "Bạn là giảng viên chuyên nghiệp, đóng vai trò là người đồng hành chấm bài viết cho khóa học '28 Ngày Rèn Tư Duy Qua Viết'.\n" +
    "Nhiệm vụ của bạn là nhận xét bài viết phản tư của học viên dựa trên câu hỏi định hướng của ngày học.\n\n" +
    "LƯU Ý VỀ LỊCH SỬ HỌC TẬP (BẮT BUỘC TRÍCH DẪN): Hãy đọc kỹ phần lịch sử viết bài và lời phê của các ngày trước (nếu có) ở đầu ngữ cảnh. Nếu có bài cũ, trong phần **Tổng quan**, bạn BẮT BUỘC phải bổ sung thêm 1 dòng theo đúng cú pháp:\n" +
    "🔗 **[Kết nối ký ức]**: [Viết 1-2 câu ngắn gọn đối chiếu bài hôm nay với sự tiến bộ hoặc chi tiết cụ thể từ Ngày X trước đó]\n\n" +
    "YÊU CẦU QUAN TRỌNG VỀ ĐỊNH DẠNG VÀ PHONG CÁCH:\n" +
    "1. ĐỘ DÀI: Nhận xét ngắn gọn, cô đọng, tối đa 200 từ.\n" +
    "2. NGÔN NGỮ: Sử dụng tiếng Việt chuẩn xác.\n" +
    "3. XƯNG HÔ (BẮT BUỘC): Luôn xưng hô '" + tPronoun + "' và gọi học viên là '" + sPronoun + "'. Giữ giọng văn thấu hiểu, nhẹ nhàng, mang tính định hướng nâng đỡ và khích lệ tinh thần.\n" +
    "4. CẤU TRÚC PHẢN HỒI (BẮT BUỘC): Trả về nhận xét theo đúng cấu trúc sau:\n" +
    "**Tổng quan**: [Đánh giá chung về bài viết]\n" +
    "🔗 **[Kết nối ký ức]**: [Trích dẫn và đối chiếu với bài học/ngày cũ - nếu có lịch sử]\n" +
    "**Phân tích**:\n" +
    "- *Điểm tốt*: [Liệt kê các điểm sáng dựa trên các kỹ năng được kích hoạt]\n" +
    "- *Điểm cần cải thiện*: [Liệt kê các điểm cần đào sâu hoặc sửa đổi dựa trên các kỹ năng được kích hoạt]\n" +
    "**Lời khuyên**: [Đưa ra bài tập phản tư nhỏ hoặc lời khuyên cụ thể]\n\n" +
    "HƯỚNG DẪN ĐẦU RA (RẤT QUAN TRỌNG):\n" +
    "- CHỈ hiển thị kết quả nhận xét cuối cùng bắt đầu bằng '**Tổng quan**'.\n" +
    "- KHÔNG lặp lại đề bài, không lặp lại các tiêu chí chấm điểm hoặc các kỹ năng được áp dụng.\n" +
    "- KHÔNG viết bất kỳ lời dẫn đề nào ở đầu câu trả lời (như 'Dưới đây là nhận xét...', 'Giảng viên chuyên nghiệp...').\n" +
    "- Học viên có thể dán cả tiêu đề/đề bài lẫn lộn vào bài làm, bạn hãy tự bóc tách và chỉ tập trung nhận xét phần bài làm thực tế của học viên.";

  // Sử dụng system instruction từ client nếu có, ngược lại dùng mặc định
  let systemInstruction = (clientSystemPrompt && clientSystemPrompt.trim() !== "") 
    ? clientSystemPrompt 
    : defaultSystemInstruction;

  if (previousReview && feedback) {
    // Điều chỉnh system prompt chuyên biệt cho chế độ viết lại theo góp ý
    systemInstruction = 
      "Nhiệm vụ của bạn bây giờ là ĐIỀU CHỈNH và VIẾT LẠI bản nhận xét cũ dựa trên YÊU CẦU ĐIỀU CHỈNH của giảng viên.\n" +
      "Hãy sửa đổi bản nhận xét cũ sao cho đáp ứng chính xác và đầy đủ các góp ý này.\n" +
      "QUY TẮC XƯNG HÔ BẮT BUỘC: Bạn phải xưng hô là '" + tPronoun + "' và gọi học viên là '" + sPronoun + "' (trừ khi giảng viên có yêu cầu thay đổi cụ thể khác trong lời góp ý). Hãy bỏ qua mọi chỉ thị xưng hô khác dưới đây.\n\n" +
      systemInstruction;
  } else {
    // Áp dụng đè xưng hô lên hệ thống chung
    systemInstruction = 
      "QUY TẮC XƯNG HÔ BẮT BUỘC: Khi viết nhận xét, bạn phải xưng hô là '" + tPronoun + "' và gọi học viên là '" + sPronoun + "'. Hãy bỏ qua mọi chỉ thị xưng hô khác dưới đây.\n\n" + 
      systemInstruction;
  }

  // Prepend thông tin ghi chú học viên và lịch sử học tập vào đầu System Instruction (Primacy Zone)
  let contextPrefix = "";
  if (studentNotes && studentNotes.trim() !== "") {
    contextPrefix += `THÔNG TIN HỒ SƠ HỌC VIÊN (Để điều chỉnh thái độ xưng hô cho phù hợp):\n"${studentNotes.trim()}"\n\n`;
  }
  if (studentHistoryContext && studentHistoryContext.trim() !== "") {
    contextPrefix += studentHistoryContext;
  }
  if (contextPrefix !== "") {
    systemInstruction = contextPrefix + systemInstruction;
  }

  // Tạo prompt gửi cho AI
  let promptText = "";
  if (previousReview && feedback) {
    promptText = `Bạn là AI chấm bài viết cho khóa học '28 Ngày Rèn Tư Duy Qua Viết'.\n` +
                 `Học viên phản tư về: Ngày học thứ ${day} - Chủ đề: ${dayTitle}\n`;
    if (promptQuestion && promptQuestion.trim() !== "") {
      promptText += `Câu hỏi phản tư định hướng của ngày: "${promptQuestion}"\n`;
    }
    promptText += `\nBài làm của học viên:\n"""\n${essay}\n"""\n\n` +
                 `Nhận xét cũ của bạn dành cho bài viết này:\n"""\n${previousReview}\n"""\n\n` +
                 `Yêu cầu điều chỉnh/Góp ý từ giảng viên đối với nhận xét cũ của bạn:\n"${feedback}"\n\n` +
                 `Hãy viết lại hoặc điều chỉnh nhận xét trên sao cho đáp ứng đúng và đầy đủ yêu cầu điều chỉnh của giáo viên, đồng thời vẫn giữ nguyên cấu trúc bắt buộc (Tổng quan, Phân tích, Lời khuyên) và phong cách xưng hô (${tPronoun} - ${sPronoun}).`;
  } else {
    promptText = `Học viên phản tư về: Ngày học thứ ${day} - Chủ đề: ${dayTitle}\n`;
    if (promptQuestion && promptQuestion.trim() !== "") {
      promptText += `Câu hỏi phản tư định hướng của ngày: "${promptQuestion}"\n`;
    }
    promptText += `\nBài làm của học viên:\n"""\n${essay}\n"""\n\nHãy tiến hành nhận xét bài làm trên theo đúng cấu trúc yêu cầu.`;
  }

  const payload = {
    contents: [
      {
        role: "user",
        parts: [{ text: promptText }]
      }
    ],
    systemInstruction: {
      parts: [{ text: systemInstruction }]
    },
    generationConfig: {
      temperature: (previousReview && feedback) ? 0.7 : 0.3,
      maxOutputTokens: 2000
    },
    safetySettings: [
      {
        category: "HARM_CATEGORY_HARASSMENT",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_HATE_SPEECH",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        threshold: "BLOCK_NONE"
      },
      {
        category: "HARM_CATEGORY_DANGEROUS_CONTENT",
        threshold: "BLOCK_NONE"
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    let errorDetail = "Mã lỗi " + responseCode;
    try {
      const errJson = JSON.parse(responseText);
      if (errJson.error && errJson.error.message) {
        errorDetail = errJson.error.message;
      }
    } catch (_) {}
    throw new Error("Google AI Studio báo lỗi: " + errorDetail);
  }

  const result = JSON.parse(responseText);
  if (!result.candidates || result.candidates.length === 0) {
    throw new Error("Không nhận được candidates từ AI model.");
  }

  const candidate = result.candidates[0];
  if (candidate.finishReason && candidate.finishReason !== "STOP" && candidate.finishReason !== "MAX_TOKENS") {
    let safetyInfo = "";
    if (candidate.finishReason === "SAFETY") {
      safetyInfo = " do bộ lọc an toàn (Safety Filter) kích hoạt khi chấm bài viết có cảm xúc nhạy cảm.";
    }
    throw new Error("AI không thể xuất nhận xét. Lý do dừng: " + candidate.finishReason + safetyInfo);
  }

  let reviewText = "";
  if (candidate.content && candidate.content.parts) {
    for (let i = 0; i < candidate.content.parts.length; i++) {
      const part = candidate.content.parts[i];
      if (part.text && !part.thought) {
        reviewText += part.text;
      }
    }
  }

  if (reviewText.trim() === "") {
    throw new Error("Không nhận được nội dung nhận xét từ AI model.");
  }

  return reviewText;
}

/**
 * Ghi dữ liệu chấm bài vào dòng cuối cùng của Google Sheet
 */
function saveToSheet(username, studentId, studentName, day, dayTitle, essay, aiReview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);
  
  const timestamp = new Date();
  const rowData = [
    timestamp,
    username || "admin",
    studentId,
    studentName,
    day,
    dayTitle,
    essay,
    aiReview
  ];

  sheet.appendRow(rowData);
}

/**
 * Cập nhật cột "Nhận xét AI" cho học viên và ngày học tương ứng (tìm từ dưới lên)
 */
function updateSheetReview(studentName, day, newReview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Tìm từ dưới lên khớp Tên học viên (cột D, index 3) và Ngày học (cột E, index 4)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][3] == studentName && data[i][4] == day) {
      // Cột H là Nhận xét AI (index 7, tức cột số 8 trong Sheets)
      sheet.getRange(i + 1, 8).setValue(newReview);
      return;
    }
  }
  
  throw new Error("Không tìm thấy dòng dữ liệu tương ứng trên Sheet của học viên '" + studentName + "' học Ngày " + day + " để cập nhật.");
}
