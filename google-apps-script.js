// Dán mã nguồn này vào trình biên tập Apps Script của Google Sheets (Extensions > Apps Script)
// Sau đó nhấn Deploy > New Deployment > Chọn Web App > Ai có quyền truy cập: Chọn "Anyone" (Bất kỳ ai).
function doPost(e) {
  try {
    var jsonString = e.postData.contents;
    var data = JSON.parse(jsonString);
    
    // TRƯỜNG HỢP 1: Yêu cầu kích hoạt lưu trữ dữ liệu tuần và gửi email kết quả
    if (data && data.action === "archive_and_email") {
      var archiveResult = archiveAndSendEmails();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", 
        message: "Lưu trữ và gửi email thành công!",
        details: archiveResult
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // TRƯỜNG HỢP 2: Đồng bộ dữ liệu Strava hàng ngày / hàng tuần
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Strava";
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    // Nếu sheet chưa tồn tại thì tạo mới
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    var dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Nhận diện loại hành động (Mặc định nếu gửi danh sách mảng thẳng là sync_weekly)
    var action = "sync_weekly";
    var runnerData = [];
    
    if (data) {
      if (Array.isArray(data)) {
        runnerData = data;
      } else if (data.action === "sync_daily") {
        action = "sync_daily";
        runnerData = data.data || [];
      } else if (data.action === "sync_weekly") {
        action = "sync_weekly";
        runnerData = data.data || [];
      } else {
        // Dự phòng cho phiên bản cũ
        runnerData = data.data || [];
      }
    }
    
    if (!runnerData || runnerData.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Không nhận được dữ liệu runner."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var rows = [];
    for (var i = 0; i < runnerData.length; i++) {
      var runner = runnerData[i];
      rows.push([
        dateString,
        runner.name,
        runner.distance // Quãng đường chạy (km)
      ]);
    }
    
    if (action === "sync_daily") {
      // ── CHẾ ĐỘ HÀNG NGÀY: GHI ĐÈ cột D, E, F ──
      
      // Khởi tạo header D1:F1 nếu trống
      var headerD1 = sheet.getRange("D1").getValue();
      if (!headerD1) {
        sheet.getRange(1, 4, 1, 3).setValues([["Ngày cập nhật", "Tên Runner", "Số km (Tuần này)"]]);
        var headerRangeD = sheet.getRange(1, 4, 1, 3);
        headerRangeD.setFontWeight("bold");
        headerRangeD.setBackgroundColor("#E8F0FE");
        headerRangeD.setHorizontalAlignment("center");
      }
      
      // Xóa sạch dữ liệu cũ cột D, E, F từ dòng 2 trở đi
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        sheet.getRange(2, 4, lastRow - 1, 3).clearContent();
      }
      
      // Ghi đè dữ liệu mới vào cột D, E, F
      sheet.getRange(2, 4, rows.length, 3).setValues(rows);
      
      // Định dạng số km (cột F) thành 1 chữ số thập phân
      sheet.getRange(2, 6, rows.length, 1).setNumberFormat("#,##0.0");
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Đã cập nhật hàng ngày thành công " + rows.length + " dòng vào cột D, E, F."
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else {
      // ── CHẾ ĐỘ HÀNG TUẦN (hoặc mặc định): GHI NỐI TIẾP cột A, B, C ──
      
      // Khởi tạo header A1:C1 nếu trống
      var headerA1 = sheet.getRange("A1").getValue();
      if (!headerA1) {
        sheet.getRange(1, 1, 1, 3).setValues([["Ngày cập nhật", "Tên Runner", "Số km (Tuần trước)"]]);
        var headerRangeA = sheet.getRange(1, 1, 1, 3);
        headerRangeA.setFontWeight("bold");
        headerRangeA.setBackgroundColor("#E8F0FE");
        headerRangeA.setHorizontalAlignment("center");
      }
      
      // Tìm hàng cuối cùng có dữ liệu thực tế tại cột A để tránh ghi đè hoặc bỏ trống hàng do cột D, E, F dài hơn
      var colAValues = sheet.getRange("A1:A").getValues();
      var lastRowABC = 0;
      for (var i = colAValues.length - 1; i >= 0; i--) {
        if (colAValues[i][0] !== "") {
          lastRowABC = i + 1;
          break;
        }
      }
      if (lastRowABC === 0) lastRowABC = 1;
      
      // Ghi nối tiếp vào cột A, B, C
      sheet.getRange(lastRowABC + 1, 1, rows.length, 3).setValues(rows);
      
      // Định dạng số km (cột C) thành 1 chữ số thập phân
      sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).setNumberFormat("#,##0.0");
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Đã cập nhật hàng tuần thành công " + rows.length + " dòng nối tiếp vào cột A, B, C."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
/**
 * Sao chép cột E, F, G từ sheet "tuần trước" và append vào sheet "lịch sử" (3 cột)
 * Đồng thời gửi email kết quả cho các runner (email nằm ở cột H của sheet Strava)
 */
function archiveAndSendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getSheetByName("tuần trước");
  var lastWeekSheet = ss.getSheetByName("lịch sử");
  var stravaSheet = ss.getSheetByName("Strava");
  if (!currentSheet) throw new Error("Không tìm thấy sheet 'tuần trước'");
  if (!lastWeekSheet) throw new Error("Không tìm thấy sheet 'lịch sử'");
  if (!stravaSheet) throw new Error("Không tìm thấy sheet 'Strava'");
  // ── 1. ĐỌC EMAIL MAP từ sheet "email" (cột A = Tên, cột B = Email) ──
  var emailSheet = ss.getSheetByName("email");
  if (!emailSheet) throw new Error("Không tìm thấy sheet 'email'");
  var emailLastRow = emailSheet.getLastRow();
  var emailMap = {};
  if (emailLastRow >= 2) {
    var emailValues = emailSheet.getRange(2, 1, emailLastRow - 1, 2).getValues();
    for (var k = 0; k < emailValues.length; k++) {
      var nameKey = emailValues[k][0] ? emailValues[k][0].toString().trim().toLowerCase() : "";
      var emailVal = emailValues[k][1] ? emailValues[k][1].toString().trim() : "";
      if (nameKey && emailVal) emailMap[nameKey] = emailVal;
    }
  }
  // Đọc link bảng thành tích từ ô F1 sheet "email"
  var groupSheetLink = emailSheet.getRange("F1").getValue().toString().trim();
  // ── 2. ĐỌC DỮ LIỆU từ sheet "tuần trước" (bắt đầu từ DÒNG 4) ──
  var DATA_START_ROW = 4;
  var lastRow = currentSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return "Không có dữ liệu trong sheet 'tuần trước'.";
  var numRows = lastRow - DATA_START_ROW + 1;
  // Đọc cột D(4), E(5), F(6), G(7) => index 0=Tên, 1=ĐăngKý, 2=ThựcTế, 3=GhiChú
  var rawData = currentSheet.getRange(DATA_START_ROW, 4, numRows, 4).getValues();
  // ── 3. TÍNH KHOẢNG THỜI GIAN TUẦN TRƯỚC ──
  var today = new Date();
  var dow = today.getDay(); // 0=CN, 1=T2...
  var daysToLastMon = (dow === 0 ? 6 : dow - 1) + 7;
  var lastMon = new Date(today); lastMon.setDate(today.getDate() - daysToLastMon);
  var lastSun = new Date(lastMon); lastSun.setDate(lastMon.getDate() + 6);
  var tz = Session.getScriptTimeZone();
  var monStr = Utilities.formatDate(lastMon, tz, "dd/MM/yyyy");
  var sunStr = Utilities.formatDate(lastSun, tz, "dd/MM/yyyy");
  var weekRangeStr = monStr + " - " + sunStr;
  var weekHeader = "Tuần từ " + Utilities.formatDate(lastMon, tz, "dd/MM") + " đến " + Utilities.formatDate(lastSun, tz, "dd/MM");
  // ── 4. APPEND 3 CỘT SANG PHẢI trong sheet "lịch sử" ──
  var newStartCol = lastWeekSheet.getLastColumn() + 1;
  // Ghi header tuần (dòng 2)
  var weekHeaderRange = lastWeekSheet.getRange(2, newStartCol, 1, 3);
  weekHeaderRange.setValue(weekHeader);
  weekHeaderRange.merge();
  weekHeaderRange.setHorizontalAlignment("center");
  weekHeaderRange.setFontWeight("bold");
  // Ghi header cột (dòng 3)
  var colHeaderRange = lastWeekSheet.getRange(3, newStartCol, 1, 3);
  colHeaderRange.setValues([["Đăng ký\n(km)", "thực tế\n(km)", "nuôi heo\n(=thiếu *20k)"]]);
  colHeaderRange.setWrap(true);
  // Ghi giá trị (CHỈ VALUE, không copy công thức) từ dòng 4
  var colsToWrite = [];
  for (var r = 0; r < rawData.length; r++) {
    var regKmVal  = rawData[r][1]; // Cột E
    var actKmVal  = rawData[r][2]; // Cột F
    var noteVal   = rawData[r][3]; // Cột G
    colsToWrite.push([regKmVal, actKmVal, noteVal]);
  }
  var destRange = lastWeekSheet.getRange(DATA_START_ROW, newStartCol, colsToWrite.length, 3);
  destRange.setValues(colsToWrite);
  // Thiết lập định dạng số để tránh bị thừa ký tự 'đ' (do thừa hưởng định dạng từ cột sát bên trái)
  var formats = [];
  for (var r = 0; r < colsToWrite.length; r++) {
    formats.push(["#,##0", "#,##0.0", "#,##0\" đ\""]);
  }
  destRange.setNumberFormats(formats);
  // Tô màu nền Cyan (#00ffff) cho tiêu đề tuần (dòng 2) và tiêu đề cột (dòng 3)
  var headersRange = lastWeekSheet.getRange(2, newStartCol, 2, 3);
  headersRange.setBackground("#00ffff");
  // Bôi viền (borders) cho toàn bộ khối tuần mới vừa thêm (từ dòng 2 đến dòng cuối)
  var entireWeekRange = lastWeekSheet.getRange(2, newStartCol, lastRow - 1, 3);
  entireWeekRange.setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    true,  // vertical
    true,  // horizontal
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID
  );
  // Đặt viền dọc ngoài cùng bên trái và bên phải dày hơn để phân tách rõ ràng giữa các tuần
  var leftEdge = lastWeekSheet.getRange(2, newStartCol, lastRow - 1, 1);
  leftEdge.setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  var rightEdge = lastWeekSheet.getRange(2, newStartCol + 2, lastRow - 1, 1);
  rightEdge.setBorder(null, null, null, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // ── 5. GỬI EMAIL cho runner có số km thực tế > 0 ──
  // 30 mẫu câu chào - Chưa đăng ký
  var caseNoReg = [
    "Chúc mừng bạn đã hoàn thành {km} km chạy thực tế trong tuần vừa qua! Tuần trước bạn chưa đăng ký mục tiêu chạy bộ, bạn hãy nhớ đăng ký mục tiêu cho tuần mới này để cùng mọi người tham gia hoạt động nhé!",
    "Tuyệt vời với {km} km tích lũy tuần qua! Tuy nhiên bạn chưa đăng ký mục tiêu tuần trước. Hãy đăng ký ngay cho tuần này để cùng nhóm đua tài nhé!",
    "Thành tích thực tế tuần qua của bạn ghi nhận là {km} km. Bạn chưa đăng ký mục tiêu tuần trước đâu nhé. Nhớ đăng ký cho tuần mới này để thử thách bản thân cùng mọi người!",
    "Chúc mừng bạn đã chạy được {km} km tuần qua! Tiếc là tuần rồi bạn chưa đăng ký số km mục tiêu. Hãy đăng ký cho tuần này để thêm động lực chạy bộ nhé!",
    "Báo cáo ghi nhận tuần trước bạn đạt được {km} km. Hãy nhớ đăng ký mục tiêu tuần này để mọi người cùng theo dõi và cổ vũ bạn nhé!",
    "Tuần qua bạn đã hoàn thành tốt {km} km chạy thực tế. Đừng quên đăng ký mục tiêu cho tuần mới ngay hôm nay để tích cực hoạt động cùng CLB!",
    "Thật tuyệt khi thấy bạn chạy được {km} km tuần trước. Để hành trình chạy bộ thú vị hơn, hãy nhớ đăng ký số km mục tiêu cho tuần này nhé!",
    "Ghi nhận {km} km nỗ lực từ bạn trong tuần vừa qua! Bạn chưa đăng ký mục tiêu tuần cũ. Hãy cập nhật mục tiêu tuần mới này ngay nhé!",
    "Cảm ơn bạn đã đóng góp {km} km vào thành tích chung tuần qua. Nhớ đăng ký mục tiêu tuần này để hoạt động của chúng ta thêm phần sôi nổi!",
    "Thành tích tuần trước của bạn đạt {km} km thực tế. Hãy đặt mục tiêu cho tuần này bằng cách đăng ký số km mong muốn nhé!",
    "Tuyệt vời! Bạn đã tích lũy được {km} km trong tuần vừa qua. Đừng quên đăng ký mục tiêu cho tuần mới để hành trình thêm phần ý nghĩa nhé!",
    "Ghi nhận {km} km chạy thực tế rất tích cực của bạn. Hãy nhớ đặt mục tiêu tuần này để có thêm động lực chinh phục những cột mốc mới nhé!",
    "Tuần qua bạn đã có những bước chạy tuyệt vời với {km} km hoàn thành. Nhớ đăng ký mục tiêu cho tuần này để cùng nhóm thi đua nhé!",
    "Chúc mừng runner đã hoàn thành {km} km thực tế tuần qua. Bạn hãy dành chút thời gian đăng ký mục tiêu tuần này để làm động lực nhé!",
    "Thật đáng quý với {km} km bạn đã đóng góp tuần trước. Nhớ đăng ký mục tiêu cho tuần mới này để thử thách bản thân nhiều hơn nữa nhé!",
    "Thành tích tuần qua của bạn là {km} km. Đừng quên đăng ký số km mục tiêu cho tuần này để mọi người cùng theo dõi và cổ vũ bạn!",
    "Chúc mừng bạn đã hoàn thành {km} km chạy bộ tuần vừa rồi. Hãy nhớ đăng ký mục tiêu cho tuần mới này để không bỏ lỡ niềm vui thi đua!",
    "Ghi nhận {km} km nỗ lực từ bạn trong tuần qua. Nhớ đăng ký mục tiêu tuần này để tiếp tục đồng hành và bứt phá cùng nhóm nhé!",
    "Bạn đã hoàn thành xuất sắc {km} km thực tế tuần qua. Đừng quên đăng ký mục tiêu cho tuần mới để tiếp thêm năng lượng chạy bộ nhé!",
    "Chúc mừng bạn với {km} km chạy thực tế cực kỳ nỗ lực tuần qua. Đăng ký mục tiêu cho tuần mới ngay để bắt đầu hành trình mới nào!",
    "Tuyệt vời, tuần qua bạn đã chạy được {km} km! Nhớ đăng ký mục tiêu cho tuần mới để thử thách giới hạn của bản thân nhé!",
    "Ghi nhận thành tích {km} km chạy thực tế rất đáng khen của bạn. Hãy đăng ký mục tiêu tuần mới ngay để cùng anh em duy trì nhịp chạy!",
    "Chúc mừng bạn đã hoàn thành {km} km chạy bộ. Đừng quên đặt mục tiêu tuần này để có thêm bia bắn và động lực xỏ giày mỗi ngày nhé!",
    "Bạn đã đạt {km} km chạy thực tế tuần qua, thật ấn tượng! Hãy đăng ký mục tiêu tuần này để cùng câu lạc bộ NH9902 bứt tốc nhé!",
    "Chúc mừng bạn đã hoàn thành {km} km. Nhớ đăng ký mục tiêu tuần mới để chúng ta cùng nhau theo dõi và tiến bộ mỗi ngày!",
    "Thành tích tuần trước của bạn là {km} km thực tế. Hãy đặt ngay mục tiêu tuần này để tiếp tục thử thách bản thân và đồng hành cùng nhóm nhé!",
    "Cảm ơn bạn đã đóng góp {km} km tuần qua. Để tuần này thêm phần hào hứng, hãy nhớ đăng ký số km mục tiêu của mình nhé!",
    "Ghi nhận nỗ lực tuyệt vời với {km} km của bạn tuần qua. Hãy đăng ký mục tiêu tuần mới này để tăng thêm phần kịch tính cho cuộc đua nhé!",
    "Chúc mừng bạn đã đạt {km} km thực tế tuần trước. Đừng quên đăng ký mục tiêu tuần mới để cùng mọi người tạo nên những kỷ lục mới!",
    "Bạn đã hoàn thành xuất sắc {km} km tuần qua dù chưa đăng ký. Hãy đặt mục tiêu cho tuần mới ngay hôm nay để hành trình chạy bộ thêm phần thú vị!"
  ];
  // 30 mẫu câu chào - Đạt mục tiêu
  var caseGoalMet = [
    "Chúc mừng bạn đã hoàn thành xuất sắc mục tiêu đã đăng ký trong tuần! Dưới đây là chi tiết thành tích chạy bộ của bạn:",
    "Quá tuyệt vời! Bạn đã vượt qua thử thách và hoàn thành mục tiêu chạy bộ của tuần một cách ngoạn mục. Dưới đây là kết quả của bạn:",
    "Xin chúc mừng! Sự kiên trì của bạn đã được đền đáp bằng việc hoàn thành mục tiêu đăng ký của tuần. Dưới đây là thông số chi tiết:",
    "Xuất sắc! Bạn đã thực hiện đúng cam kết mục tiêu của mình tuần qua. Hãy tiếp tục giữ vững phong độ này trong tuần mới nhé!",
    "Mục tiêu tuần đã hoàn thành! Cảm ơn sự nỗ lực bền bỉ của bạn trên từng cung đường chạy. Thành tích chi tiết của bạn:",
    "Thật tự hào khi bạn đã cán đích mục tiêu đã đăng ký tuần qua! Chúc bạn tiếp tục gặt hái thêm nhiều km trong tuần này. Chi tiết thành tích:",
    "Chúc mừng bạn đã hoàn thành mục tiêu! Bạn đang có một tinh thần kỷ luật thật đáng nể. Xem chi tiết kết quả tuần qua bên dưới:",
    "Đã hoàn thành mục tiêu tuần xuất sắc! Chúc mừng bạn đã vượt qua chính mình của tuần trước. Dưới đây là thành tích ghi nhận:",
    "Tuyệt cú mèo! Bạn đã chinh phục thành công số km đăng ký tuần qua. Cùng xem lại kết quả chạy bộ chi tiết của bạn nào:",
    "Chiến thắng mục tiêu tuần! Nỗ lực tuyệt vời của bạn đã đem lại thành quả xứng đáng. Dưới đây là báo cáo chạy bộ của bạn:",
    "Thật ấn tượng! Bạn đã hoàn thành trọn vẹn mục tiêu đã đề ra tuần qua. Hãy cùng xem lại thành tích đáng tự hào của bạn:",
    "Chúc mừng bạn đã xuất sắc cán mốc mục tiêu tuần! Tinh thần thể thao của bạn là tấm gương cho cả nhóm. Chi tiết thành tích:",
    "Mục tiêu đã được chinh phục hoàn toàn! Sự chăm chỉ của bạn tuần qua thật đáng ngưỡng mộ. Dưới đây là kết quả chi tiết của bạn:",
    "Chúc mừng runner đã hoàn thành chỉ tiêu đăng ký tuần vừa rồi! Hãy tiếp tục duy trì phong độ đỉnh cao này nhé. Dưới đây là kết quả:",
    "Tuyệt vời ông mặt trời! Bạn đã hoàn thành xuất sắc số km đăng ký. Hãy cùng xem lại báo cáo thành tích chạy bộ tuần qua:",
    "Chúc mừng bạn đã vượt qua mục tiêu đã đặt ra! Mỗi km bạn chạy đều mang lại nguồn năng lượng tích cực cho nhóm. Kết quả của bạn:",
    "Bạn đã làm rất tốt khi hoàn thành đúng cam kết tuần qua! Tinh thần kỷ luật của bạn thật tuyệt vời. Dưới đây là chi tiết thành tích:",
    "Xin chúc mừng bạn đã về đích mục tiêu tuần xuất sắc! Hãy cùng giữ vững ngọn lửa đam mê chạy bộ này nhé. Kết quả chi tiết của bạn:",
    "Không gì có thể cản bước bạn! Chúc mừng bạn đã hoàn thành xuất sắc mục tiêu chạy bộ tuần qua. Thành tích cụ thể của bạn:",
    "Thật tự hào với tinh thần kiên cường hoàn thành mục tiêu tuần của bạn! Hãy tiếp tục bứt phá trong tuần mới nhé. Kết quả chi tiết:",
    "Chúc mừng bạn đã chinh phục thành công thử thách tuần qua! Sự nỗ lực của bạn đã mang lại kết quả xứng đáng. Xem chi tiết dưới đây:",
    "Mục tiêu đăng ký tuần qua đã được bạn hoàn thành một cách xuất sắc! Chúc bạn giữ vững phong độ này cho tuần mới. Kết quả của bạn:",
    "Quá xuất sắc! Bạn đã thực hiện đúng cam kết số km đăng ký tuần qua. Cùng xem lại thành tích chạy bộ chi tiết của bạn:",
    "Chúc mừng bạn đã hoàn thành mục tiêu tuần vừa rồi! Tinh thần bền bỉ và kỷ luật của bạn rất đáng biểuương. Thành tích của bạn:",
    "Tuyệt vời! Bạn đã hoàn thành xuất sắc kế hoạch chạy bộ tuần qua. Hãy cùng hướng tới những mục tiêu mới cao hơn nhé! Kết quả chi tiết:",
    "Chúc mừng bạn đã vượt qua thử thách đăng ký tuần trước! Sự chăm chỉ xỏ giày mỗi ngày của bạn đã được đền đáp xứng đáng. Kết quả của bạn:",
    "Bạn đã hoàn thành xuất sắc mục tiêu chạy bộ của tuần! Hãy tiếp tục lan tỏa tinh thần thể thao năng động này nhé. Xem kết quả chi tiết:",
    "Chúc mừng bạn đã cán đích mục tiêu tuần thành công! Từng bước chạy của bạn đều ghi dấu sự nỗ lực vượt bậc. Dưới đây là thành tích:",
    "Thật xuất sắc khi bạn đã đạt được mục tiêu đăng ký tuần qua! Hãy tiếp tục giữ vững nhịp chạy đều đặn này nhé. Chi tiết thành tích:",
    "Mục tiêu tuần đã được chinh phục thành công! Chúc mừng bạn đã vượt qua giới hạn của bản thân tuần vừa rồi. Thành tích chi tiết của bạn:"
  ];
  // 30 mẫu câu chào - Chưa đạt mục tiêu
  var caseGoalNotMet = [
    "Dưới đây là chi tiết thành tích chạy bộ tuần vừa qua của bạn. Rất tiếc bạn chưa hoàn thành mục tiêu đăng ký ban đầu. Hãy nỗ lực hơn trong tuần này nhé!",
    "Tuần qua bạn chưa hoàn thành mục tiêu đăng ký của mình rồi. Đừng nản lòng, hãy tiếp tục cố gắng hơn trong tuần mới này nhé! Chi tiết thành tích:",
    "Mục tiêu chạy bộ tuần trước chưa đạt được rồi bạn ơi. Cùng nỗ lực phục thù và hoàn thành kế hoạch ở tuần mới này nhé! Kết quả tuần qua:",
    "Rất tiếc khi tuần rồi số km thực tế của bạn chưa chạm tới mốc đăng ký. Hãy sạc đầy năng lượng để bứt phá trong tuần này nhé! Kết quả chi tiết:",
    "Mục tiêu tuần chưa hoàn thành hoàn toàn, nhưng mỗi km bạn chạy đều rất đáng quý. Cố lên nhé, tuần mới đang chờ đón bạn! Chi tiết chạy bộ:",
    "Mục tiêu đăng ký chưa đạt được tuần qua. Hãy coi đó là động lực để tuần này chúng ta chạy chăm chỉ và đều đặn hơn nha! Kết quả cụ thể:",
    "Chưa hoàn thành kế hoạch tuần qua rồi runner ơi. Hãy kiên trì rèn luyện và bám sát mục tiêu của tuần này nhé! Dưới đây là thành tích:",
    "Có chút tiếc nuối khi số km thực tế tuần qua chưa đạt mục tiêu đăng ký. Hãy cùng đồng đội bứt tốc trong tuần mới này nhé! Kết quả chi tiết:",
    "Mục tiêu tuần trước chưa vượt qua được. Không sao cả, hãy đặt mục tiêu phù hợp hơn cho tuần này và quyết tâm cán đích nhé! Kết quả cụ thể:",
    "Số km đăng ký chưa được hoàn thành trọn vẹn tuần vừa qua. Chúc bạn sẽ có một tuần mới đầy năng lượng để phục thù mục tiêu chạy bộ!",
    "Đừng nản lòng khi mục tiêu tuần trước chưa đạt được hoàn toàn nhé. Mỗi bước chạy đều giúp bạn khỏe hơn mỗi ngày. Kết quả tuần qua của bạn:",
    "Rất tiếc vì tuần qua bạn chưa đạt số km đăng ký, nhưng nỗ lực của bạn vẫn rất đáng ghi nhận. Hãy cố gắng hơn tuần này nhé! Kết quả chi tiết:",
    "Mục tiêu tuần trước chưa hoàn thành, nhưng tuần mới là cơ hội mới để bạn chinh phục thử thách. Cố gắng lên nhé! Dưới đây là kết quả:",
    "Có vẻ tuần qua bạn bận rộn nên chưa đạt mục tiêu đăng ký. Hãy lên kế hoạch chạy đều đặn hơn cho tuần này nha. Kết quả chi tiết của bạn:",
    "Chưa đạt mục tiêu tuần vừa rồi không sao cả, hãy coi đó là bước đệm để tuần này bứt phá mạnh mẽ hơn. Xem chi tiết thành tích tuần qua:",
    "Dù chưa đạt được số km đăng ký ban đầu, nỗ lực của bạn vẫn rất đáng trân trọng. Chúc bạn tuần mới chạy năng suất hơn! Kết quả chi tiết:",
    "Tuần qua mục tiêu chưa hoàn thành trọn vẹn, hãy điều chỉnh nhịp chạy để cán đích thành công trong tuần mới này nhé. Kết quả chi tiết:",
    "Đừng lo lắng khi chưa hoàn thành mục tiêu chạy bộ tuần trước. Hãy giữ vững tinh thần và phục thù ở tuần mới này nha! Chi tiết thành tích:",
    "Rất tiếc khi số km thực tế chưa chạm mốc đăng ký tuần qua. Hãy sạc lại năng lượng để sẵn sàng cho chặng đường tuần này nhé. Kết quả chi tiết:",
    "Mục tiêu chưa đạt được nhưng hành trình của bạn vẫn đang tiến về phía trước. Hãy nỗ lực hết mình cho tuần này nhé! Kết quả chạy bộ:",
    "Dù chưa hoàn thành mục tiêu đăng ký tuần qua, bạn đã rất cố gắng rồi. Hãy tiếp tục duy trì thói quen xỏ giày trong tuần mới nhé! Kết quả:",
    "Chưa đạt mục tiêu tuần trước chỉ là thử thách nhỏ thôi. Hãy đặt quyết tâm cao hơn và chinh phục tuần này nhé! Chi tiết thành tích của bạn:",
    "Dù chưa hoàn thành xuất sắc số km đăng ký tuần qua, mỗi bước chạy của bạn vẫn tiếp thêm sức mạnh cho nhóm. Kết quả chi tiết của bạn:",
    "Mục tiêu tuần trước chưa đạt được hoàn toàn, hãy cố gắng sắp xếp thời gian chạy đều hơn ở tuần mới này nha. Chi tiết thành tích:",
    "Đừng buồn vì chưa hoàn thành mục tiêu tuần qua. Hãy cùng đồng đội xỏ giày lên đường phục thù trong tuần này nhé! Kết quả chi tiết:",
    "Mục tiêu chưa đạt được tuần qua sẽ là động lực lớn để tuần này bạn chạy chăm chỉ và đều đặn hơn. Dưới đây là kết quả của bạn:",
    "Có chút tiếc nuối khi số km thực tế tuần qua chưa đạt mục tiêu. Hãy cùng nhóm NH9902 bứt tốc mạnh mẽ hơn trong tuần mới này nhé! Kết quả:",
    "Chưa đạt mục tiêu tuần qua không làm giảm đi tinh thần thể thao của bạn. Hãy sẵn sàng cho những mục tiêu mới tuần này nhé! Kết quả chi tiết:",
    "Rất tiếc khi tuần rồi số km thực tế chưa đạt đăng ký. Hãy đặt mục tiêu vừa sức và hoàn thành nó thật tốt trong tuần mới này nha! Kết quả:",
    "Mục tiêu tuần qua chưa hoàn thành trọn vẹn. Hãy cố gắng duy trì nhịp chạy đều đặn mỗi ngày để chinh phục tuần mới nhé! Kết quả chi tiết:"
  ];
  // 30 mẫu câu kết bài
  var footers = [
    "Thói quen chạy bộ không chỉ rèn luyện thể chất mà còn xây dựng tính kỷ luật tuyệt vời. Hãy cùng đồng đội tiếp tục chinh phục những cung đường mới trong tuần này nhé!",
    "Mỗi bước chạy hôm nay là một bước tiến gần hơn tới phiên bản khỏe mạnh hơn của chính bạn. Chúc bạn có một tuần mới ngập tràn năng lượng và duy trì thói quen chạy bộ tuyệt vời này nhé!",
    "Chạy bộ là hành trình kiên trì và bền bỉ. Hãy tiếp tục giữ vững ngọn lửa đam mê, xỏ giày lên đường và cùng đồng đội NH9902 tích lũy thêm nhiều cây số trong tuần mới nha!",
    "Đừng so sánh mình với bất kỳ ai, hãy chạy vì sức khỏe và niềm vui của chính bạn. Chúc bạn tuần mới nhiều sức khỏe, luôn duy trì được tinh thần thể thao năng động cùng nhóm!",
    "Kỷ luật là cầu nối giữa mục tiêu và thành công. Cảm ơn bạn vì đã luôn chăm chỉ rèn luyện thể chất mỗi ngày. Cùng bứt phá hơn nữa trong những ngày sắp tới nhé!",
    "Những cung đường phía trước đang vẫy gọi! Hãy tiếp tục duy trì tinh thần chạy bộ thể thao, cùng kết nối và truyền cảm hứng tập luyện cho các thành viên khác trong nhóm nhé!",
    "Không quan trọng bạn chạy nhanh hay chậm, miễn là bạn không dừng lại. Hãy giữ vững thói quen tốt này và cùng đồng đội NH9902 chinh phục những cột mốc km tiếp theo!",
    "Chạy bộ giúp chúng ta rèn luyện sức bền và giải tỏa căng thẳng sau giờ làm việc. Chúc bạn tuần mới dồi dào sức khỏe và luôn tràn đầy hứng khởi trên mọi hành trình chạy bộ!",
    "Mỗi km chạy đều ghi dấu sự nỗ lực và tính kỷ luật đáng tự hào của bạn. Hãy tiếp tục giữ vững nhịp chạy và cùng mọi người lan tỏa lối sống lành mạnh này nhé!",
    "Sức khỏe là tài sản vô giá nhất. Cảm ơn bạn đã luôn đồng hành và đóng góp những bước chạy ý nghĩa cùng NH9902 Running Club. Chúc bạn tuần mới ngập tràn niềm vui và năng lượng!",
    "Hãy để mỗi bước chạy là một niềm vui giải tỏa mọi lo toan cuộc sống. Chúc bạn tuần mới dồi dào sức khỏe và luôn tràn đầy cảm hứng trên đường chạy!",
    "Sự kiên trì xỏ giày mỗi ngày chính là chìa khóa mở ra sức khỏe và sự dẻo dai. Chúc bạn có một tuần mới nhiều năng lượng và bứt phá mạnh mẽ!",
    "Cảm ơn runner đã luôn đồng hành và truyền lửa chạy bộ cho nhóm. Chúc bạn tuần mới nhiều niềm vui, tràn đầy sức khỏe và gặt hái nhiều km mới nhé!",
    "Chạy bộ không chỉ giúp khỏe mạnh mà còn giúp chúng ta vượt qua giới hạn của chính mình. Chúc bạn một tuần mới ngập tràn hứng khởi và niềm vui!",
    "Mỗi cung đường bạn đi qua đều là một trải nghiệm tuyệt vời của bản thân. Chúc bạn tuần mới tràn đầy năng lượng tích cực và giữ vững nhịp chạy!",
    "Hãy biến chạy bộ thành người bạn đồng hành không thể thiếu mỗi ngày. Chúc bạn có tuần mới dồi dào sức khỏe và chinh phục thêm nhiều km ý nghĩa!",
    "Sự nỗ lực bền bỉ của bạn là mảnh ghép tuyệt vời cho phong trào chạy bộ của nhóm. Chúc bạn tuần mới ngập tràn niềm vui, sức khỏe và năng lượng!",
    "Chạy bộ giúp rèn luyện ý chí kiên cường trước mọi thử thách. Chúc bạn tuần mới gặt hái được nhiều thành công và luôn giữ vững đam mê chạy bộ!",
    "Hãy luôn nhớ rằng mỗi bước chạy của bạn đều hướng tới một phiên bản hoàn hảo hơn. Chúc bạn tuần mới ngập tràn năng lượng và tràn đầy niềm vui!",
    "Cảm ơn bạn vì sự tích cực đóng góp số km cho câu lạc bộ tuần qua. Chúc bạn tuần mới thật nhiều sức khỏe, may mắn và luôn vui tươi xỏ giày!",
    "Duy trì nhịp chạy đều đặn chính là cách tốt nhất để rèn luyện cả thân và tâm. Chúc bạn tuần mới nhiều sức khỏe và tràn đầy hứng khởi trên đường chạy!",
    "Đam mê chạy bộ sẽ dẫn lối bạn đến những giới hạn mới đầy bất ngờ. Chúc bạn tuần mới ngập tràn niềm vui, năng lượng tích cực và luôn khỏe mạnh!",
    "Mỗi km chạy được là một chiến thắng nhỏ đối với bản thân. Chúc bạn có một tuần mới ngập tràn niềm vui, sức khỏe dồi dào và luôn giữ vững phong độ!",
    "Chạy bộ là cách tuyệt vời để kết nối tình đồng đội và lan tỏa lối sống lành mạnh. Chúc bạn tuần mới vui vẻ, nhiều sức khỏe và luôn tràn đầy năng lượng!",
    "Hãy tiếp tục viết nên những câu chuyện đẹp trên đường chạy cùng NH9902 nhé! Chúc bạn tuần mới ngập tràn niềm vui, hạnh phúc và dồi dào sức khỏe!",
    "Mỗi bước chân chạy bộ đều giúp bạn tích lũy thêm sức mạnh cho tương lai. Chúc bạn tuần mới tràn đầy hứng khởi, may mắn và luôn giữ được lửa đam mê!",
    "Cảm ơn runner vì tinh thần kỷ luật và những đóng góp tích cực cho nhóm. Chúc bạn tuần mới dồi dào sức khỏe, luôn tràn đầy năng lượng trên mọi hành trình!",
    "Chạy bộ mang lại cho chúng ta sức khỏe, sự tự tin và tinh thần lạc quan. Chúc bạn tuần mới nhiều sức khỏe, vui tươi và gặt hái thêm nhiều km nhé!",
    "Hãy để tiếng bước chân trên đường chạy là giai điệu năng lượng cho tuần mới của bạn. Chúc bạn tuần mới dồi dào sức khỏe và luôn tràn đầy hạnh phúc!",
    "Hành trình vạn dặm bắt đầu từ một bước chân, và bạn đang làm rất tốt hành trình đó. Chúc bạn tuần mới đầy năng lượng, sức khỏe và luôn yêu đời!"
  ];
  var emailCount = 0;
  var errorMailList = [];
  for (var i = 0; i < rawData.length; i++) {
    var runnerName = rawData[i][0] ? rawData[i][0].toString().trim() : "";
    var regKm  = Number(rawData[i][1]) || 0;
    var actKm  = Number(rawData[i][2]) || 0;
    var note   = rawData[i][3] ? rawData[i][3].toString().trim() : "";
    // Chỉ gửi mail cho người có km thực tế > 0
    if (actKm <= 0) continue;
    var email = runnerName ? (emailMap[runnerName.toLowerCase()] || "") : "";
    if (!email || email.indexOf("@") === -1) continue;
    try {
      var subject = "Kết quả chạy bộ tuần từ " + weekRangeStr;
      var ri = Math.floor(Math.random() * 30);
      var rf = Math.floor(Math.random() * 30);
      var greeting = "";
      if (regKm === 0) {
        greeting = caseNoReg[ri].replace("{km}", actKm);
      } else if (actKm >= regKm) {
        greeting = caseGoalMet[ri];
      } else {
        greeting = caseGoalNotMet[ri];
      }
      var footer = footers[rf];
      // Format ghi chú
      var noteDisplay = "";
      if (note) {
        var ln = note.toLowerCase();
        var hasUnit = ln.endsWith("đ") || ln.endsWith("d") || ln.endsWith("k") || ln.endsWith("đồng") || ln.endsWith("dong");
        noteDisplay = "Nuôi heo " + note + (hasUnit ? "" : "đ");
      }
      // Link bảng thành tích
      var leaderboardHtml = groupSheetLink
        ? `<p style="margin-top:15px;font-size:14px;color:#555;">🏆 Xem bảng thành tích của nhóm tại: <a href="${groupSheetLink}" style="color:#fc4c02;font-weight:bold;">${groupSheetLink}</a></p>`
        : "";
      var htmlBody = `
        <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,0.05);">
          <div style="background:linear-gradient(135deg,#fc4c02,#ff8000);padding:24px;text-align:center;color:white;">
            <h2 style="margin:0;font-size:24px;font-weight:bold;letter-spacing:0.5px;">NH9902 RUNNING CLUB</h2>
            <p style="margin:5px 0 0 0;opacity:0.9;font-size:14px;">Báo cáo kết quả thành tích ${weekHeader}</p>
          </div>
          <div style="padding:24px;background:#fff;color:#333;line-height:1.6;">
            <p style="font-size:16px;margin-top:0;">Chào <strong>${runnerName}</strong>,</p>
            <p>${greeting}</p>
            <div style="background:#f7f7f7;border-left:4px solid #fc4c02;padding:15px;margin:20px 0;border-radius:4px;">
              <table style="width:100%;border-collapse:collapse;">
                <tr>
                  <td style="padding:6px 0;color:#666;font-size:14px;width:45%;">Tên vận động viên:</td>
                  <td style="padding:6px 0;font-weight:bold;color:#111;">${runnerName}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#666;font-size:14px;">Số km đăng ký:</td>
                  <td style="padding:6px 0;font-weight:bold;color:#111;">${regKm} km</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#666;font-size:14px;">Số km thực tế đã chạy:</td>
                  <td style="padding:6px 0;font-weight:bold;color:#fc4c02;font-size:18px;">${actKm} km</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#666;font-size:14px;">Ghi chú:</td>
                  <td style="padding:6px 0;font-weight:bold;color:#111;">${noteDisplay}</td>
                </tr>
              </table>
            </div>
            <p>${footer}</p>
            ${leaderboardHtml}
            <div style="text-align:center;margin:30px 0 10px 0;">
              <a href="https://www.strava.com/clubs/nh9902" style="background:#fc4c02;color:white;padding:12px 25px;text-decoration:none;font-weight:bold;border-radius:4px;display:inline-block;">Xem câu lạc bộ trên Strava</a>
            </div>
          </div>
          <div style="background:#f0f0f0;padding:15px;text-align:center;font-size:12px;color:#888;border-top:1px solid #e0e0e0;">
            <p style="margin:0;">Email này được gửi tự động từ hệ thống quản lý NH9902 Running Club.</p>
            <p style="margin:5px 0 0 0;">© 2026 NH9902 Running Club. All rights reserved.</p>
          </div>
        </div>`;
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
      emailCount++;
    } catch (mailError) {
      errorMailList.push(email + ": " + mailError.toString());
    }
  }
  return {
    weekArchived: weekHeader,
    newColumnStart: newStartCol,
    totalRows: colsToWrite.length,
    sentEmails: emailCount,
    errors: errorMailList
  };
}
/**
 * Cập nhật tiêu đề tuần ở ô E2 (merged E2:G2) trên sheet "tuần trước"
 * Format: "Tuần từ DD/M đến DD/M" (tuần tính từ thứ 2 đến chủ nhật)
 * 
 * Thiết kế để chạy bằng Trigger vào 1h sáng thứ 3 hàng tuần.
 * Khi chạy vào thứ 3, hàm sẽ tính ngày thứ 2 (hôm qua) và chủ nhật (6 ngày sau thứ 2)
 * của tuần hiện tại và ghi vào ô tiêu đề.
 */
function updateWeekHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("tuần trước");
  if (!sheet) throw new Error("Không tìm thấy sheet 'tuần trước'");
  // Tính ngày thứ 2 (Monday) của tuần hiện tại
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0=CN, 1=T2, 2=T3, ..., 6=T7
  // Tính offset để lùi về thứ 2: nếu CN(0) thì lùi 6 ngày, còn lại lùi (dayOfWeek - 1) ngày
  var offsetToMonday = (dayOfWeek === 0) ? -6 : 1 - dayOfWeek;
  
  var monday = new Date(today);
  monday.setDate(today.getDate() + offsetToMonday - 7);
  
  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  // Format DD/M (không có số 0 đứng trước)
  var monStr = monday.getDate() + "/" + (monday.getMonth() + 1);
  var sunStr = sunday.getDate() + "/" + (sunday.getMonth() + 1);
  var weekLabel = "Tuần từ " + monStr + " đến " + sunStr;
  // Ghi vào ô E2 (ô đã merge E2:G2)
  sheet.getRange("E2").setValue(weekLabel);
  Logger.log("Đã cập nhật tiêu đề tuần: " + weekLabel);
}
// ---------------------- Refresh Filters for "tuần trước" & "tuần này" ----------------------
/**
 * Làm mới (refresh) filter cho cả 2 sheet, giữ nguyên các tiêu chí lọc đã đặt.
 * Cơ chế: lưu lại criteria → xóa filter cũ → tạo filter mới → áp dụng lại criteria.
 */
function refreshFilter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheets = ['tuần trước', 'tuần này'];
  for (var s = 0; s < targetSheets.length; s++) {
    var sheet = ss.getSheetByName(targetSheets[s]);
    if (!sheet) continue;
    var existingFilter = sheet.getFilter();
    if (!existingFilter) {
      // Chưa có filter → tạo mới (bắt đầu từ hàng 3 – hàng header)
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow >= 3 && lastCol > 0) {
        sheet.getRange(3, 1, lastRow - 2, lastCol).createFilter();
      }
      continue;
    }
    // ── Lưu lại thông tin filter hiện tại ──
    var filterRange = existingFilter.getRange();
    var startRow = filterRange.getRow();
    var startCol = filterRange.getColumn();
    var numCols  = filterRange.getNumColumns();
    // Lưu criteria từng cột (ví dụ: bạn đã chọn lọc theo giá trị ở cột F)
    var savedCriteria = {};
    for (var col = startCol; col < startCol + numCols; col++) {
      var criteria = existingFilter.getColumnFilterCriteria(col);
      if (criteria) {
        savedCriteria[col] = criteria.copy().build();
      }
    }
    // ── Xóa filter cũ ──
    existingFilter.remove();
    // ── Tạo filter mới (mở rộng đến hàng cuối cùng hiện tại) ──
    var currentLastRow = Math.max(sheet.getLastRow(), startRow + 1);
    var newFilter = sheet.getRange(startRow, startCol, currentLastRow - startRow + 1, numCols).createFilter();
    // ── Áp dụng lại criteria đã lưu ──
    for (var colStr in savedCriteria) {
      try {
        newFilter.setColumnFilterCriteria(parseInt(colStr), savedCriteria[colStr]);
      } catch (err) {
        Logger.log('Không thể khôi phục criteria cột ' + colStr + ': ' + err);
      }
    }
  }
  Logger.log('✅ Filter đã được refresh (giữ nguyên criteria) cho "tuần trước" và "tuần này".');
}
/**
 * Trigger onChange – gọi refreshFilter() khi có thay đổi dữ liệu.
 * Bao gồm thay đổi trên sheet Strava (ảnh hưởng VLOOKUP ở các sheet đích).
 */
function onChange(e) {
  var relevant = [
    'INSERT_ROW', 'INSERT_COLUMN',
    'REMOVE_ROW', 'REMOVE_COLUMN',
    'EDIT'
  ];
  if (relevant.indexOf(e.changeType) > -1) {
    var sheetName = e.source.getActiveSheet().getName();
    // Refresh khi thay đổi trên sheet đích hoặc sheet Strava (ảnh hưởng VLOOKUP)
    if (sheetName === 'tuần trước' || sheetName === 'tuần này' || sheetName === 'Strava') {
      refreshFilter();
    }
  }
}
/**
 * Gọi hàm này một lần để tạo trigger onChange.
 * Sau khi chạy, trigger sẽ tự động cập nhật filter khi dữ liệu thay đổi.
 */
function setupRefreshTrigger() {
  // Xóa trigger cũ nếu đã tồn tại
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onChange') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Tạo trigger mới
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
  Logger.log('Trigger onChange đã được thiết lập.');
}
/**
 * Cập nhật tiêu đề tuần ở ô E2 (merged E2:G2) trên sheet "tuần này"
 * Format: "Tuần từ DD/M đến DD/M" (tuần tính từ thứ 2 đến chủ nhật)
 *
 * Hàm này sẽ được chạy vào 7h sáng thứ Hai hàng tuần.
 */
function updateCurrentWeekHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("tuần này");
  if (!sheet) throw new Error("Không tìm thấy sheet 'tuần này'");
  // Tính ngày thứ 2 (Monday) của tuần hiện tại
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0=CN, 1=T2, 2=T3, ..., 6=T7
  var offsetToMonday = (dayOfWeek === 0) ? -6 : 1 - dayOfWeek;
  var monday = new Date(today);
  monday.setDate(today.getDate() + offsetToMonday);
  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  var monStr = monday.getDate() + "/" + (monday.getMonth() + 1);
  var sunStr = sunday.getDate() + "/" + (sunday.getMonth() + 1);
  var weekLabel = "Tuần từ " + monStr + " đến " + sunStr;
  sheet.getRange("E2").setValue(weekLabel);
  Logger.log("Đã cập nhật tiêu đề tuần trên sheet 'tuần này': " + weekLabel);
}
/**
 * Thiết lập trigger tự động chạy hàm updateCurrentWeekHeader vào 7h sáng thứ Hai hàng tuần.
 * Chạy một lần để tạo trigger.
 */
function setupWeeklyUpdateTrigger() {
  // Xóa trigger cũ nếu tồn tại
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateCurrentWeekHeader") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Tạo trigger mới
  ScriptApp.newTrigger("updateCurrentWeekHeader")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  Logger.log("Trigger cho cập nhật tiêu đề tuần đã được thiết lập mỗi thứ Hai 7h sáng.");
}
