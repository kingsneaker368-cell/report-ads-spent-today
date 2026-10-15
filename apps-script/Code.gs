/**
 * Shared report capture scheduler.
 *
 * Install this script in ONE control Google Spreadsheet only.
 * It dispatches both report repositories when a configured slot matches.
 * Set Script Property GITHUB_TOKEN to a fine-grained PAT with Contents: write
 * access for both report repositories. Never commit tokens here.
 */

const REPORT_CAPTURE_OWNER = "kingsneaker368-cell";
const REPORT_CAPTURE_REPOS = [
  "sheet-export-telegram-bot",
  "report-ads-spent-today"
];
const REPORT_CAPTURE_CONFIG_SHEET = "CAU_HINH_CHUP_NO";
const REPORT_CAPTURE_DEFAULT_TIME_ZONE = "Asia/Ho_Chi_Minh";
const REPORT_CAPTURE_DEFAULT_TIMES = ["10:00", "14:00", "17:00", "00:30"];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CHỤP BÁO CÁO")
    .addItem("CHỤP CẢ 2 BÁO CÁO NGAY", "captureBothReportsNow")
    .addItem("Chụp sheet-export-telegram-bot ngay", "captureSheetExportNow")
    .addItem("Chụp report-ads-spent-today ngay", "captureAdsSpentNow")
    .addSeparator()
    .addItem("Khởi tạo cấu hình (giữ thiết lập hiện có)", "initializeReportCaptureConfig")
    .addItem("Cài kiểm tra lịch tự động", "installReportCapturePollTrigger")
    .addItem("Gỡ kiểm tra lịch tự động", "removeReportCapturePollTrigger")
    .addToUi();
}

function captureBothReportsNow() {
  dispatchReportCapture_("report-capture-manual", REPORT_CAPTURE_REPOS, {});
  SpreadsheetApp.getActive().toast("Đã gửi yêu cầu chụp cho cả 2 báo cáo. Kiểm tra GitHub Actions để xem kết quả.", "Chụp báo cáo", 8);
}

function captureSheetExportNow() {
  dispatchReportCapture_("report-capture-manual", ["sheet-export-telegram-bot"], {});
  SpreadsheetApp.getActive().toast("Đã gửi yêu cầu chụp sheet-export-telegram-bot. Kiểm tra GitHub Actions để xem kết quả.", "Chụp báo cáo", 8);
}

function captureAdsSpentNow() {
  dispatchReportCapture_("report-capture-manual", ["report-ads-spent-today"], {});
  SpreadsheetApp.getActive().toast("Đã gửi yêu cầu chụp report-ads-spent-today. Kiểm tra GitHub Actions để xem kết quả.", "Chụp báo cáo", 8);
}

function dispatchReportCapture_(eventType, repoNames, clientPayload) {
  const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) {
    throw new Error("Thiếu GITHUB_TOKEN trong Apps Script > Project Settings > Script Properties.");
  }

  const errors = [];
  const successes = [];
  for (const repoName of repoNames) {
    const url = `https://api.github.com/repos/${REPORT_CAPTURE_OWNER}/${repoName}/dispatches`;
    try {
      const response = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28"
        },
        payload: JSON.stringify({ event_type: eventType, client_payload: clientPayload }),
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      if (status !== 204) throw new Error(`HTTP ${status}: ${response.getContentText()}`);
      successes.push(repoName);
    } catch (err) {
      errors.push(`${repoName}: ${err.message || err}`);
    }
  }

  updateReportCaptureStatus_(successes, errors, clientPayload);
  if (errors.length) {
    throw new Error("Một số yêu cầu chụp chưa gửi được. " + errors.join(" | "));
  }
}

function pollReportCaptureSchedule() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_CAPTURE_CONFIG_SHEET);
    if (!sheet) throw new Error(`Không tìm thấy tab ${REPORT_CAPTURE_CONFIG_SHEET}`);

    const enabled = String(sheet.getRange("B3").getDisplayValue() || "").trim().toUpperCase();
    if (!["BẬT", "BAT", "ON", "TRUE", "1"].includes(enabled)) return;

    const timeZone = getReportCaptureTimeZone_(sheet);
    const now = new Date();
    const currentTime = Utilities.formatDate(now, timeZone, "HH:mm");
    const date = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const configuredTimes = sheet.getRange("B5:B8").getDisplayValues().flat()
      .map(normalizeReportCaptureTime_)
      .filter(Boolean);

    if (!configuredTimes.includes(currentTime)) return;

    const key = `${date}|${currentTime}|${timeZone}`;
    const props = PropertiesService.getScriptProperties();
    const pendingRepos = REPORT_CAPTURE_REPOS.filter(repo =>
      props.getProperty(`REPORT_CAPTURE_LAST_${repo}`) !== key
    );
    if (!pendingRepos.length) return;

    const payload = { scheduleDate: date, scheduleTime: currentTime, timeZone };
    const token = props.getProperty("GITHUB_TOKEN");
    if (!token) throw new Error("Thiếu GITHUB_TOKEN trong Script Properties.");

    const successes = [];
    const errors = [];
    for (const repoName of pendingRepos) {
      const url = `https://api.github.com/repos/${REPORT_CAPTURE_OWNER}/${repoName}/dispatches`;
      try {
        const response = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28"
          },
          payload: JSON.stringify({ event_type: "report-capture-auto", client_payload: payload }),
          muteHttpExceptions: true
        });
        const status = response.getResponseCode();
        if (status !== 204) throw new Error(`HTTP ${status}: ${response.getContentText()}`);
        props.setProperty(`REPORT_CAPTURE_LAST_${repoName}`, key);
        successes.push(repoName);
      } catch (err) {
        errors.push(`${repoName}: ${err.message || err}`);
      }
    }
    updateReportCaptureStatus_(successes, errors, payload);
  } finally {
    lock.releaseLock();
  }
}

function initializeReportCaptureConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REPORT_CAPTURE_CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(REPORT_CAPTURE_CONFIG_SHEET);

  // Defaults are applied once on this bound control spreadsheet; future user edits are preserved.
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("REPORT_CAPTURE_CONFIG_INITIALIZED") !== "true") {
    sheet.getRange("A3").setValue("Tự động");
    if (!String(sheet.getRange("B3").getDisplayValue()).trim()) sheet.getRange("B3").setValue("BẬT");
    sheet.getRange("A5:A8").setValues([["Lần chụp 1"], ["Lần chụp 2"], ["Lần chụp 3"], ["Lần chụp 4"]]);
    sheet.getRange("B5:B8").setValues(REPORT_CAPTURE_DEFAULT_TIMES.map(t => [t]));
    sheet.getRange("A9").setValue("Múi giờ");
    if (!String(sheet.getRange("B9").getDisplayValue()).trim()) sheet.getRange("B9").setValue(REPORT_CAPTURE_DEFAULT_TIME_ZONE);
    sheet.getRange("A11").setValue("Lần gửi yêu cầu gần nhất");
    sheet.getRange("A12").setValue("Trạng thái gửi yêu cầu");
    sheet.getRange("A13").setValue("Mã lịch gửi cuối");
    props.setProperty("REPORT_CAPTURE_CONFIG_INITIALIZED", "true");
  }
  SpreadsheetApp.getActive().toast("Cấu hình sẵn sàng. Sửa giờ tại B5:B8 và múi giờ IANA tại B9.", "Chụp báo cáo", 8);
}

function installReportCapturePollTrigger() {
  initializeReportCaptureConfig();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "pollReportCaptureSchedule")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("pollReportCaptureSchedule").timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().toast("Đã cài kiểm tra lịch mỗi phút. Giờ chạy được đọc trực tiếp từ tab cấu hình.", "Chụp báo cáo", 8);
}

function removeReportCapturePollTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "pollReportCaptureSchedule")
    .forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getActive().toast("Đã gỡ lịch tự động. Các nút chụp thủ công vẫn dùng được.", "Chụp báo cáo", 8);
}

function getReportCaptureTimeZone_(sheet) {
  const candidate = String(sheet.getRange("B9").getDisplayValue() || "").trim() || REPORT_CAPTURE_DEFAULT_TIME_ZONE;
  try {
    Utilities.formatDate(new Date(), candidate, "yyyy-MM-dd HH:mm");
    return candidate;
  } catch (err) {
    throw new Error(`Múi giờ tại B9 không hợp lệ: ${candidate}. Ví dụ: Asia/Ho_Chi_Minh`);
  }
}

function normalizeReportCaptureTime_(value) {
  const s = String(value || "").trim();
  let match = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (match) {
    const hour = Number(match[1]);
    const minute = Number(match[2]);
    if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
      return String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");
    }
  }
  match = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour < 1 || hour > 12 || minute > 59) return null;
  hour = hour % 12 + (match[3].toUpperCase() === "PM" ? 12 : 0);
  return String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");
}

function updateReportCaptureStatus_(successes, errors, payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_CAPTURE_CONFIG_SHEET);
  if (!sheet) return;
  const timeZone = payload.timeZone || getReportCaptureTimeZone_(sheet);
  const nowText = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
  if (successes.length) sheet.getRange("B11").setValue(nowText);
  const parts = [];
  if (successes.length) parts.push("ĐÃ GỬI: " + successes.join(", "));
  if (errors.length) parts.push("LỖI: " + errors.join(" | "));
  sheet.getRange("B12").setValue(parts.join(" ; ") || "Không có yêu cầu mới");
  if (payload.scheduleDate && payload.scheduleTime) {
    sheet.getRange("B13").setValue(`${payload.scheduleDate}|${payload.scheduleTime}|${timeZone}`);
  } else if (successes.length) {
    sheet.getRange("B13").setValue("THỦ CÔNG|" + nowText);
  }
}
