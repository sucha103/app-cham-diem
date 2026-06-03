/**
 * Google Apps Script - Web App Proxy & Sheet Logger
 * Dự án: Mini App Chấm Bài 28 Ngày
 * 
 * LƯU Ý TRIỂN KHAI:
 * 1. Mở Google Sheet của bạn.
 * 2. Chọn Tiện ích mở rộng -> Apps Script.
 * 3. Xóa code cũ và dán toàn bộ đoạn code này vào.
 * 4. Bấm Lưu (Ctrl+S).
 * 5. Bấm "Triển khai" -> "Triển khai mới".
 * 6. Chọn loại triển khai: "Ứng dụng web".
 * 7. Cấu hình:
 *    - Thực thi với tư cách: "Tôi" (Tài khoản Google của bạn)
 *    - Ai có quyền truy cập: "Mọi người" (để ứng dụng HTML chạy cục bộ gọi được)
 * 8. Bấm "Triển khai", cấp quyền và sao chép URL Web App được tạo ra để dán vào cài đặt trên ứng dụng.
 */

// Tiêu đề các cột dữ liệu
const SHEET_HEADERS = ["Thời gian chấm", "Mã học viên", "Tên học viên", "Ngày học", "Tên ngày", "Bài làm", "Nhận xét AI"];
const TAB_NAME = "Dữ liệu chấm bài";

/**
 * Xử lý yêu cầu CORS preflight (OPTIONS)
 */
function doGet(e) {
  return ContentService.createTextOutput("Mini App Chấm Bài 28 Ngày - Apps Script Web App đang chạy tốt!")
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
    // 1. Kiểm tra và parse payload
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Không tìm thấy dữ liệu yêu cầu.");
    }
    
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || "grade_and_save"; // Mặc định là chấm bài và lưu

    // Lấy thông tin đầu vào
    const apiKey = payload.apiKey;
    const studentId = payload.studentId || "";
    const studentName = payload.studentName || "";
    const day = payload.day || "";
    const dayTitle = payload.dayTitle || "";
    const promptQuestion = payload.promptQuestion || "";
    const essay = payload.essay || "";
    const aiReviewPreGenerated = payload.aiReview; // Dùng khi đồng bộ lại, cập nhật hoặc điều chỉnh
    const clientSystemPrompt = payload.systemPrompt; // System prompt gửi động từ client
    const feedback = payload.feedback || ""; // Góp ý chỉnh sửa từ giáo viên

    // A. Hành động chỉ lưu (Sync Queue)
    if (action === "save_only") {
      if (!aiReviewPreGenerated) {
        throw new Error("Không có nhận xét AI để lưu.");
      }
      saveToSheet(studentId, studentName, day, dayTitle, essay, aiReviewPreGenerated);
      return createJsonResponse({ success: true, msg: "Đã lưu vào Google Sheet thành công.", sheetUrl: sheetUrl, sheetName: sheetName });
    }

    // B. Hành động cập nhật nhận xét (khi Giáo viên chỉnh sửa trên giao diện)
    if (action === "update_review") {
      if (!studentName || !day) {
        throw new Error("Thiếu Tên học viên hoặc Ngày học để cập nhật.");
      }
      if (!aiReviewPreGenerated) {
        throw new Error("Nhận xét mới không được để trống.");
      }
      updateSheetReview(studentName, day, aiReviewPreGenerated);
      return createJsonResponse({ success: true, msg: "Đã cập nhật nhận xét thành công trên Google Sheet.", sheetUrl: sheetUrl, sheetName: sheetName });
    }

    // B2. Hành động AI viết lại nhận xét dựa trên phản hồi của Giáo viên
    if (action === "rewrite_review") {
      if (!apiKey) {
        return createJsonResponse({ success: false, errorType: "CONFIG_ERROR", error: "Thiếu Gemma API Key. Vui lòng nhập trong phần Cài đặt ứng dụng.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!studentName || !day) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Thiếu Tên học viên hoặc Ngày học để cập nhật.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!feedback) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Vui lòng nhập góp ý/yêu cầu điều chỉnh của bạn.", sheetUrl: sheetUrl, sheetName: sheetName });
      }
      if (!aiReviewPreGenerated) {
        return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Không tìm thấy nhận xét cũ để chỉnh sửa.", sheetUrl: sheetUrl, sheetName: sheetName });
      }

      let newReview = "";
      try {
        newReview = callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt, aiReviewPreGenerated, feedback);
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
        sheetError: sheetSyncFailed ? "Ghi đè Sheet thất bại: " + sheetErrorMsg : null,
        sheetUrl: sheetUrl,
        sheetName: sheetName
      });
    }

    // C. Hành động Chấm Bài & Tự Động Lưu (Mặc định)
    // Xác thực đầu vào cho tính năng Chấm Bài
    if (!apiKey) {
      return createJsonResponse({ success: false, errorType: "CONFIG_ERROR", error: "Thiếu Gemma API Key. Vui lòng nhập trong phần Cài đặt ứng dụng.", sheetUrl: sheetUrl, sheetName: sheetName });
    }
    if (!studentName) {
      return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Tên học viên không được để trống.", sheetUrl: sheetUrl, sheetName: sheetName });
    }
    if (!essay) {
      return createJsonResponse({ success: false, errorType: "VALIDATION_ERROR", error: "Nội dung bài làm không được để trống.", sheetUrl: sheetUrl, sheetName: sheetName });
    }

    // 2. Gọi Google AI Studio API để chấm bài
    let aiReview = "";
    try {
      aiReview = callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt);
    } catch (aiError) {
      return createJsonResponse({
        success: false,
        errorType: "AI_ERROR",
        error: "Lỗi kết nối Gemma AI: " + aiError.message,
        sheetUrl: sheetUrl,
        sheetName: sheetName
      });
    }

    // 3. Ghi dữ liệu vào Google Sheet
    let sheetSyncFailed = false;
    let sheetErrorMsg = "";
    try {
      saveToSheet(studentId, studentName, day, dayTitle, essay, aiReview);
    } catch (sheetError) {
      sheetSyncFailed = true;
      sheetErrorMsg = sheetError.message;
    }

    // 4. Trả kết quả về cho client
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
function callGemmaAPI(apiKey, day, dayTitle, promptQuestion, essay, clientSystemPrompt, previousReview, feedback) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemma-4-26b-a4b-it:generateContent?key=" + apiKey;
  
  // Xây dựng system instruction mặc định
  const defaultSystemInstruction = 
    "Bạn là giảng viên chuyên nghiệp, đóng vai trò là người đồng hành chấm bài viết cho khóa học '28 Ngày Rèn Tư Duy Qua Viết'.\n" +
    "Nhiệm vụ của bạn là nhận xét bài viết phản tư của học viên dựa trên câu hỏi định hướng của ngày học.\n" +
    "YÊU CẦU QUAN TRỌNG VỀ ĐỊNH DẠNG VÀ PHONG CÁCH:\n" +
    "1. XƯNG HÔ: Xưng là 'chị' và gọi học viên là 'em' (thể hiện sự thân thiện, ấm áp và nâng đỡ).\n" +
    "2. ĐỘ DÀI: Nhận xét ngắn gọn, cô đọng, tối đa 200 từ.\n" +
    "3. NGÔN NGỮ: Sử dụng tiếng Việt chuẩn xác, giàu tính khích lệ nhưng vẫn thẳng thắn chỉ ra điểm cần cải thiện.\n" +
    "4. CẤU TRÚC PHẢN HỒI (BẮT BUỘC): Trả về nhận xét theo đúng cấu trúc sau:\n" +
    "**Tổng quan**: [Nhận xét chung về bài viết, đánh giá xem em đã làm rõ ràng chưa, đã trả lời đúng câu hỏi cốt lõi của ngày chưa]\n" +
    "**Phân tích**:\n" +
    "- *Điểm tốt*: [Liệt kê ngắn gọn 1-2 điểm sáng tư duy hoặc cách hành văn tốt]\n" +
    "- *Điểm cần cải thiện*: [Liệt kê ngắn gọn điểm chưa sâu sắc, cần làm rõ thêm]\n" +
    "**Lời khuyên**: [Lời khuyên/gợi ý cụ thể để rèn luyện thêm tư duy viết phản tư cho bài này]\n\n" +
    "LƯU Ý THÊM: Học viên có thể dán cả tiêu đề/đề bài lẫn lộn vào bài làm, bạn hãy tự bóc tách và chỉ tập trung nhận xét phần bài làm thực tế của học viên.";

  // Sử dụng system instruction từ client nếu có, ngược lại dùng mặc định
  const systemInstruction = (clientSystemPrompt && clientSystemPrompt.trim() !== "") 
    ? clientSystemPrompt 
    : defaultSystemInstruction;

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
                 `Hãy viết lại hoặc điều chỉnh nhận xét trên sao cho đáp ứng đúng và đầy đủ yêu cầu điều chỉnh của giáo viên, đồng thời vẫn giữ nguyên cấu trúc bắt buộc (Tổng quan, Phân tích, Lời khuyên) và phong cách xưng hô (chị - em).`;
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
      temperature: 0.3,
      maxOutputTokens: 2000
    }
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
  if (!result.candidates || result.candidates.length === 0 || 
      !result.candidates[0].content || !result.candidates[0].content.parts || 
      result.candidates[0].content.parts.length === 0) {
    throw new Error("Không nhận được nội dung nhận xét từ AI model.");
  }

  return result.candidates[0].content.parts[0].text;
}

/**
 * Ghi dữ liệu chấm bài vào dòng cuối cùng của Google Sheet
 */
function saveToSheet(studentId, studentName, day, dayTitle, essay, aiReview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Script chưa được gắn (bind) với Google Sheet. Vui lòng mở Apps Script từ Google Sheet của bạn.");
  }

  let sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    // Tự động tạo tab nếu chưa tồn tại
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(SHEET_HEADERS);
    // Định dạng dòng tiêu đề
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#F3F4F6")
      .setHorizontalAlignment("center");
  }

  // Chuẩn bị dòng dữ liệu
  const timestamp = new Date();
  const rowData = [
    timestamp,
    studentId,
    studentName,
    day,
    dayTitle,
    essay,
    aiReview
  ];

  // Ghi vào sheet
  sheet.appendRow(rowData);
}

/**
 * Cập nhật cột "Nhận xét AI" cho học viên và ngày học tương ứng (tìm từ dưới lên)
 */
function updateSheetReview(studentName, day, newReview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Script chưa được gắn (bind) với Google Sheet.");
  }

  const sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    throw new Error("Không tìm thấy tab '" + TAB_NAME + "' trên Google Sheet.");
  }

  const data = sheet.getDataRange().getValues();
  
  // Tìm từ dưới lên để cập nhật dòng mới nhất khớp Tên học viên (cột C, index 2) và Ngày học (cột D, index 3)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][2] == studentName && data[i][3] == day) {
      // Cột G là Nhận xét AI (index 6, tức là cột số 7 trong Sheets)
      sheet.getRange(i + 1, 7).setValue(newReview);
      return;
    }
  }
  
  throw new Error("Không tìm thấy dòng dữ liệu tương ứng trên Sheet của học viên '" + studentName + "' học Ngày " + day + " để cập nhật.");
}
