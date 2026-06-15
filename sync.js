const axios = require('axios');
// Get configuration from environment variables
const CLUB_ID = process.env.STRAVA_CLUB_ID || 'nh9902';
const COOKIE = process.env.STRAVA_COOKIE;
const WEBAPP_URL = process.env.GOOGLE_SHEET_WEBAPP_URL;
// Nhận diện chế độ chạy: 'daily' (Tuần này) hoặc 'weekly' (Tuần trước)
// Có thể truyền qua biến môi trường SYNC_MODE hoặc tham số dòng lệnh (vd: node sync.js --daily)
const isDaily = process.env.SYNC_MODE === 'daily' || process.argv.includes('--daily');
const mode = isDaily ? 'daily' : 'weekly';
const weekOffset = isDaily ? 0 : 1; 
async function run() {
  if (!COOKIE) {
    console.error('Lỗi: Thiếu biến môi trường STRAVA_COOKIE. Hãy cấu hình trong Secrets.');
    process.exit(1);
  }
  if (!WEBAPP_URL) {
    console.error('Lỗi: Thiếu biến môi trường GOOGLE_SHEET_WEBAPP_URL. Hãy cấu hình trong Secrets.');
    process.exit(1);
  }
  const url = `https://www.strava.com/clubs/${CLUB_ID}/leaderboard?week_offset=${weekOffset}`;
  console.log(`Bắt đầu lấy dữ liệu bảng xếp hạng từ Strava Club: ${CLUB_ID} (${mode === 'daily' ? 'Tuần này' : 'Tuần trước'})...`);
  try {
    const response = await axios.get(url, {
      headers: {
        'Accept': 'text/javascript, application/javascript, application/ecmascript, application/x-ecmascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Cookie': COOKIE,
        'Referer': `https://www.strava.com/clubs/${CLUB_ID}`
      }
    });
    const data = response.data;
    if (!data || !data.data || !Array.isArray(data.data)) {
      throw new Error('Định dạng dữ liệu trả về từ Strava không đúng hoặc cookie đã hết hạn. Hãy kiểm tra lại.');
    }
    const runners = data.data.map(item => {
      const firstName = item.athlete_firstname || '';
      const lastName = item.athlete_lastname || '';
      const fullName = `${firstName} ${lastName}`.trim() || 'Vận động viên ẩn';
      
      // Distance is returned in meters, convert to km (with 1 decimal place to match Strava web)
      const distanceMeters = item.distance || 0;
      const distanceKm = Math.floor((distanceMeters / 1000) * 10) / 10;
      return {
        name: fullName,
        distance: distanceKm
      };
    });
    console.log(`Đã lấy thành công dữ liệu của ${runners.length} runner.`);
    if (runners.length === 0) {
      console.log(`Không có hoạt động nào được ghi nhận trong ${mode === 'daily' ? 'tuần này' : 'tuần trước'}.`);
      return;
    }
    console.log(`Bắt đầu gửi dữ liệu (${mode}) sang Google Sheet...`);
    
    // Đóng gói payload có kèm action để Google Apps Script phân biệt daily / weekly
    const payload = {
      action: mode === 'daily' ? 'sync_daily' : 'sync_weekly',
      data: runners
    };
    const postResponse = await axios.post(WEBAPP_URL, payload, {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    console.log('Phản hồi từ Google Apps Script:', postResponse.data);
    if (postResponse.data && postResponse.data.status === 'success') {
      console.log('Đã cập nhật Google Sheet thành công!');
    } else {
      console.error('Cập nhật thất bại:', postResponse.data.message || 'Lỗi không xác định');
      process.exit(1);
    }
  } catch (error) {
    console.error('Có lỗi xảy ra trong quá trình xử lý:');
    if (error.response) {
      console.error(`Status code: ${error.response.status}`);
      console.error(error.response.data);
    } else {
      console.error(error.message);
    }
    process.exit(1);
  }
}
run();
