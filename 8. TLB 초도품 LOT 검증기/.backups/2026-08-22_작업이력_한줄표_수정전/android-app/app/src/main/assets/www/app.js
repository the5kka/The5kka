"use strict";

const LOT_LENGTH = 9;
const HISTORY_KEY = "tlbLotInspectionHistoryV1";
const HISTORY_PASSWORD_KEY = "tlbLotHistoryPasswordV1";
const AUTO_RESET_MINUTES_KEY = "tlbLotAutoResetMinutesV1";
const DEFAULT_HISTORY_PASSWORD = "5300";
const MASTER_PASSWORD = "0160";
const DEFAULT_AUTO_RESET_MINUTES = 10;
const ALLOWED_AUTO_RESET_MINUTES = Object.freeze([0, 1, 5, 10, 30]);
const AUTO_RESET_TICK_MS = 1000;

const elements = {
  lotNo: document.getElementById("lotNo"),
  lotCameraScanButton: document.getElementById("lotCameraScanButton"),
  clearLotButton: document.getElementById("clearLotButton"),
  lotMessage: document.getElementById("lotMessage"),
  resultValue: document.getElementById("resultValue"),
  autoResetStatus: document.getElementById("autoResetStatus"),
  verificationArea: document.getElementById("verificationArea"),
  verificationStatus: document.getElementById("verificationStatus"),
  verifyPanelButton: document.getElementById("verifyPanelButton"),
  appSettingsButton: document.getElementById("appSettingsButton"),
  appSettingsDialog: document.getElementById("appSettingsDialog"),
  appSettingsForm: document.getElementById("appSettingsForm"),
  autoResetMinutesSelect: document.getElementById("autoResetMinutesSelect"),
  cancelAppSettingsButton: document.getElementById("cancelAppSettingsButton"),
  progressButton: document.getElementById("progressButton"),
  historyCountBadge: document.getElementById("historyCountBadge"),
  historyDialog: document.getElementById("historyDialog"),
  closeHistoryButton: document.getElementById("closeHistoryButton"),
  historyTotalCount: document.getElementById("historyTotalCount"),
  historyOkCount: document.getElementById("historyOkCount"),
  historyNgCount: document.getElementById("historyNgCount"),
  historyEmpty: document.getElementById("historyEmpty"),
  historyGroups: document.getElementById("historyGroups"),
  historyLotSearch: document.getElementById("historyLotSearch"),
  historyDateFilter: document.getElementById("historyDateFilter"),
  clearHistoryFiltersButton: document.getElementById("clearHistoryFiltersButton"),
  historyFilterSummary: document.getElementById("historyFilterSummary"),
  changeHistoryPasswordButton: document.getElementById("changeHistoryPasswordButton"),
  historyPasswordDialog: document.getElementById("historyPasswordDialog"),
  historyPasswordForm: document.getElementById("historyPasswordForm"),
  historyPasswordInput: document.getElementById("historyPasswordInput"),
  historyPasswordMessage: document.getElementById("historyPasswordMessage"),
  openPasswordSettingsButton: document.getElementById("openPasswordSettingsButton"),
  cancelHistoryPasswordButton: document.getElementById("cancelHistoryPasswordButton"),
  passwordSettingsDialog: document.getElementById("passwordSettingsDialog"),
  passwordSettingsForm: document.getElementById("passwordSettingsForm"),
  masterPasswordInput: document.getElementById("masterPasswordInput"),
  newHistoryPasswordInput: document.getElementById("newHistoryPasswordInput"),
  confirmHistoryPasswordInput: document.getElementById("confirmHistoryPasswordInput"),
  passwordSettingsMessage: document.getElementById("passwordSettingsMessage"),
  cancelPasswordSettingsButton: document.getElementById("cancelPasswordSettingsButton"),
  judgmentDialog: document.getElementById("judgmentDialog"),
  judgmentMark: document.getElementById("judgmentMark"),
  judgmentTitle: document.getElementById("judgmentTitle"),
  judgmentMessage: document.getElementById("judgmentMessage"),
  judgmentExpected: document.getElementById("judgmentExpected"),
  judgmentRecognized: document.getElementById("judgmentRecognized"),
  retryPhotoButton: document.getElementById("retryPhotoButton"),
  confirmJudgmentButton: document.getElementById("confirmJudgmentButton"),
  toast: document.getElementById("toast"),
};

let currentLot = "";
let scanInProgress = false;
let verificationInProgress = false;
let pendingJudgment = null;
let judgmentFinalizing = false;
let retryCount = 0;
let autoResetTimer = 0;
let autoResetDeadline = 0;
let autoResetLot = "";
let activeAutoResetMinutes = 0;
let workStartedAtIso = "";
let toastTimer = 0;
let passwordSettingsReturnTarget = "none";

function normalizeValue(value) {
  return String(value || "").trim().replace(/[\r\n\t\s]/g, "").toUpperCase();
}

function validateLot(value) {
  const lot = normalizeValue(value);
  if (!lot) return { lot: "", error: "" };
  if (lot.length !== LOT_LENGTH) return { lot, error: `TLB LOT NO는 정확히 ${LOT_LENGTH}자리여야 합니다. (현재 ${lot.length}자리)` };
  if (!/^[A-Z0-9]{9}$/.test(lot)) return { lot, error: "TLB LOT NO에는 영문 A~Z와 숫자만 사용할 수 있습니다." };
  return { lot, error: "" };
}

function getAutoResetMinutes() {
  try {
    const savedValue = localStorage.getItem(AUTO_RESET_MINUTES_KEY);
    if (savedValue === null) return DEFAULT_AUTO_RESET_MINUTES;
    const saved = Number(savedValue);
    return ALLOWED_AUTO_RESET_MINUTES.includes(saved) ? saved : DEFAULT_AUTO_RESET_MINUTES;
  } catch (_error) { return DEFAULT_AUTO_RESET_MINUTES; }
}

function saveAutoResetMinutes(minutes) {
  if (!ALLOWED_AUTO_RESET_MINUTES.includes(minutes)) return false;
  try { localStorage.setItem(AUTO_RESET_MINUTES_KEY, String(minutes)); return true; }
  catch (_error) { return false; }
}

function formatAutoResetMinutes(minutes) { return minutes === 0 ? "사용 안 함" : `${minutes}분`; }

function formatCountdown(milliseconds) {
  const totalSeconds = Math.max(0, Math.ceil(milliseconds / 1000));
  return `${String(Math.floor(totalSeconds / 60)).padStart(2, "0")}:${String(totalSeconds % 60).padStart(2, "0")}`;
}

function clearAutoResetTimer(clearWorkStart = true) {
  window.clearTimeout(autoResetTimer);
  autoResetTimer = 0;
  autoResetDeadline = 0;
  autoResetLot = "";
  activeAutoResetMinutes = 0;
  elements.autoResetStatus.textContent = "";
  elements.autoResetStatus.classList.add("is-hidden");
  if (clearWorkStart) workStartedAtIso = "";
}

function handleAutoResetTick() {
  window.clearTimeout(autoResetTimer);
  autoResetTimer = 0;
  if (!currentLot || currentLot !== autoResetLot || !autoResetDeadline) { clearAutoResetTimer(); return; }
  const remaining = autoResetDeadline - Date.now();
  if (remaining <= 0) {
    if (verificationInProgress || elements.judgmentDialog.open) {
      elements.autoResetStatus.textContent = `${formatAutoResetMinutes(activeAutoResetMinutes)} 경과 · 촬영 판정 완료 후 자동 초기화됩니다.`;
      elements.autoResetStatus.classList.remove("is-hidden");
      autoResetTimer = window.setTimeout(handleAutoResetTick, AUTO_RESET_TICK_MS);
      return;
    }
    const completedMinutes = activeAutoResetMinutes;
    resetCurrentWork();
    showToast(`${formatAutoResetMinutes(completedMinutes)}이 지나 현재 LOT를 자동 초기화했습니다. 작업이력은 유지됩니다.`);
    return;
  }
  elements.autoResetStatus.textContent = `${formatCountdown(remaining)} 후 현재 LOT 자동 초기화 · 작업이력은 유지`;
  elements.autoResetStatus.classList.remove("is-hidden");
  autoResetTimer = window.setTimeout(handleAutoResetTick, AUTO_RESET_TICK_MS);
}

function startAutoResetTimer(lot, forceRestart = false) {
  if (!forceRestart && autoResetLot === lot && autoResetDeadline > Date.now()) { handleAutoResetTick(); return; }
  window.clearTimeout(autoResetTimer);
  autoResetLot = lot;
  workStartedAtIso = new Date().toISOString();
  activeAutoResetMinutes = getAutoResetMinutes();
  if (activeAutoResetMinutes === 0) {
    autoResetDeadline = 0;
    elements.autoResetStatus.textContent = "자동 초기화 사용 안 함 · 설정에서 변경할 수 있습니다.";
    elements.autoResetStatus.classList.remove("is-hidden");
    return;
  }
  autoResetDeadline = Date.now() + (activeAutoResetMinutes * 60 * 1000);
  handleAutoResetTick();
}

function updateLot() {
  const { lot, error } = validateLot(elements.lotNo.value);
  elements.lotNo.value = lot;
  elements.lotNo.setAttribute("aria-invalid", error ? "true" : "false");
  elements.lotMessage.textContent = error;
  document.getElementById("card-lot").classList.toggle("invalid", Boolean(error));
  document.getElementById("card-lot").classList.toggle("valid", Boolean(lot) && !error);
  if (!lot || error) {
    currentLot = "";
    clearAutoResetTimer();
    elements.resultValue.textContent = error ? "LOT NO를 다시 확인하세요" : "카드의 LOT NO를 스캔하세요";
    elements.resultValue.classList.add("empty");
    elements.verificationArea.classList.add("is-hidden");
    elements.verificationArea.setAttribute("aria-hidden", "true");
    elements.verifyPanelButton.disabled = true;
    return;
  }
  const lotChanged = currentLot !== lot;
  currentLot = lot;
  if (lotChanged) retryCount = 0;
  elements.resultValue.textContent = lot;
  elements.resultValue.classList.remove("empty");
  elements.verificationArea.classList.remove("is-hidden");
  elements.verificationArea.setAttribute("aria-hidden", "false");
  elements.verifyPanelButton.disabled = !isPanelVerificationAvailable();
  elements.verificationStatus.textContent = isPanelVerificationAvailable() ? "초도품 LOT 촬영 대기" : "초도품 사진 대조는 Android 앱에서 사용합니다.";
  startAutoResetTimer(lot);
}

function resetCurrentWork() {
  clearAutoResetTimer();
  verificationInProgress = false;
  pendingJudgment = null;
  retryCount = 0;
  elements.lotNo.value = "";
  setWorkControlsDisabled(false);
  updateLot();
  focusLotStep();
}

function showToast(message) {
  window.clearTimeout(toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  toastTimer = window.setTimeout(() => elements.toast.classList.remove("show"), 2400);
}

function isAndroidCameraAvailable() { return Boolean(window.AndroidScanner && typeof window.AndroidScanner.scanLotBarcode === "function"); }
function isPanelVerificationAvailable() { return Boolean(window.AndroidScanner && typeof window.AndroidScanner.captureAndVerifyPanel === "function"); }
function focusLotStep() { if (isAndroidCameraAvailable()) elements.lotCameraScanButton.focus(); else elements.lotNo.focus(); }

function setWorkControlsDisabled(disabled) {
  elements.lotCameraScanButton.disabled = disabled;
  elements.clearLotButton.disabled = disabled;
  elements.verifyPanelButton.disabled = disabled || !currentLot || !isPanelVerificationAvailable();
}

function beginPanelVerification() {
  if (!currentLot) { showToast("먼저 카드의 9자리 LOT NO를 스캔해 주세요."); return; }
  if (!isPanelVerificationAvailable()) { showToast("초도품 사진 대조는 Android 앱에서만 사용할 수 있습니다."); return; }
  if (verificationInProgress) return;
  verificationInProgress = true;
  setWorkControlsDisabled(true);
  elements.verificationStatus.textContent = "카메라 촬영 및 문자 인식 진행 중...";
  try { window.AndroidScanner.captureAndVerifyPanel(currentLot); }
  catch (_error) {
    verificationInProgress = false;
    setWorkControlsDisabled(false);
    elements.verificationStatus.textContent = "카메라를 열지 못했습니다. 다시 시도해 주세요.";
    showToast("초도품 촬영 카메라를 열지 못했습니다.");
  }
}

function openDialog(dialog) {
  if (typeof dialog.showModal === "function") { if (!dialog.open) dialog.showModal(); }
  else dialog.setAttribute("open", "");
}
function closeDialog(dialog) {
  if (typeof dialog.close === "function" && dialog.open) dialog.close();
  else dialog.removeAttribute("open");
}

function showJudgmentDialog(judgment, rawText, normalizedText, reason) {
  pendingJudgment = {
    judgment, reason, expected: currentLot,
    recognizedRaw: String(rawText || "").replace(/\s+/g, " ").trim(),
    recognizedNormalized: String(normalizedText || ""), lot: currentLot,
    startedAtIso: workStartedAtIso, retryCount,
  };
  const isOk = judgment === "OK";
  elements.judgmentDialog.classList.toggle("ok-dialog", isOk);
  elements.judgmentDialog.classList.toggle("ng-dialog", !isOk);
  elements.judgmentMark.textContent = judgment;
  elements.judgmentTitle.textContent = isOk ? "TLB 초도품 LOT 대조 OK" : "조장 확인 필요!!";
  elements.judgmentMessage.textContent = isOk
    ? "카드 LOT와 초도품 사진의 LOT가 일치합니다. 확인을 누르면 이력을 저장하고 현재 작업을 초기화합니다."
    : `${reason} 사진 각도나 빛 반사 문제일 수 있으므로 다시 찍을 수 있습니다.`;
  elements.judgmentExpected.textContent = pendingJudgment.expected || "-";
  elements.judgmentRecognized.textContent = pendingJudgment.recognizedRaw || "인식되지 않음";
  elements.retryPhotoButton.classList.toggle("is-hidden", isOk);
  elements.confirmJudgmentButton.textContent = isOk ? "확인 및 작업 완료" : "NG 확정 · 조장 확인";
  elements.verificationStatus.textContent = isOk ? "사진 대조 OK" : "사진 대조 NG · 재촬영 또는 조장 확인";
  openDialog(elements.judgmentDialog);
}

function loadHistory() {
  try { const parsed = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); return Array.isArray(parsed) ? parsed : []; }
  catch (_error) { return []; }
}
function saveHistoryRecord(record) {
  try { const history = loadHistory(); history.unshift(record); localStorage.setItem(HISTORY_KEY, JSON.stringify(history)); return true; }
  catch (_error) { return false; }
}
function getHistoryPassword() {
  try { const saved = String(localStorage.getItem(HISTORY_PASSWORD_KEY) || ""); return /^\d{4,12}$/.test(saved) ? saved : DEFAULT_HISTORY_PASSWORD; }
  catch (_error) { return DEFAULT_HISTORY_PASSWORD; }
}
function saveHistoryPassword(password) {
  try { localStorage.setItem(HISTORY_PASSWORD_KEY, password); return true; }
  catch (_error) { return false; }
}
function formatDateKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "날짜 미확인";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
function getHistoryRecordDate(item) {
  const date = new Date(String(item.completedAtIso || item.completedAt || ""));
  return Number.isNaN(date.getTime()) ? null : date;
}
function appendHistoryDetail(container, label, value, wide = false) {
  const detail = document.createElement("div");
  detail.className = wide ? "history-detail history-detail-wide" : "history-detail";
  const labelElement = document.createElement("span"); labelElement.textContent = label;
  const valueElement = document.createElement("strong"); valueElement.textContent = String(value ?? "-");
  detail.append(labelElement, valueElement); container.appendChild(detail);
}
function updateHistoryBadge() { elements.historyCountBadge.textContent = String(loadHistory().length); }

function renderHistory() {
  const history = loadHistory();
  const lotQuery = normalizeValue(elements.historyLotSearch.value);
  const dateQuery = elements.historyDateFilter.value;
  const filtered = history.filter((item) => {
    const date = getHistoryRecordDate(item);
    return (!lotQuery || normalizeValue(item.lot).includes(lotQuery)) && (!dateQuery || formatDateKey(date) === dateQuery);
  });
  elements.historyCountBadge.textContent = String(history.length);
  elements.historyTotalCount.textContent = String(history.length);
  elements.historyOkCount.textContent = String(history.filter((item) => item.judgment === "OK").length);
  elements.historyNgCount.textContent = String(history.filter((item) => item.judgment === "NG").length);
  elements.historyFilterSummary.textContent = (lotQuery || dateQuery) ? `검색 결과 ${filtered.length}건 / 전체 ${history.length}건` : `전체 작업이력 ${history.length}건`;
  elements.historyEmpty.textContent = history.length === 0 ? "아직 완료된 TLB 검사 이력이 없습니다." : "검색 조건에 맞는 작업이력이 없습니다.";
  elements.historyEmpty.classList.toggle("is-hidden", filtered.length > 0);
  elements.historyGroups.classList.toggle("is-hidden", filtered.length === 0);
  elements.historyGroups.replaceChildren();
  const groups = new Map();
  filtered.forEach((item) => {
    const dateKey = formatDateKey(getHistoryRecordDate(item));
    if (!groups.has(dateKey)) groups.set(dateKey, []);
    groups.get(dateKey).push(item);
  });
  groups.forEach((records, dateKey) => {
    const group = document.createElement("section"); group.className = "history-date-group";
    const header = document.createElement("div"); header.className = "history-date-header";
    const title = document.createElement("h3");
    title.textContent = dateKey === "날짜 미확인" ? dateKey : new Date(`${dateKey}T00:00:00`).toLocaleDateString("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit", weekday: "short" });
    const summary = document.createElement("span");
    summary.textContent = `${records.length}건 · OK ${records.filter((item) => item.judgment === "OK").length} · NG ${records.filter((item) => item.judgment === "NG").length}`;
    header.append(title, summary);
    const list = document.createElement("div"); list.className = "history-record-list";
    records.forEach((item) => {
      const card = document.createElement("article"); card.className = "history-record-card";
      const top = document.createElement("div"); top.className = "history-record-top";
      const time = document.createElement("span"); time.className = "history-record-time";
      const recordDate = getHistoryRecordDate(item); time.textContent = recordDate ? recordDate.toLocaleTimeString("ko-KR", { hour12: false }) : "시간 미확인";
      const badge = document.createElement("span"); badge.className = `history-judgment-badge ${item.judgment === "OK" ? "history-ok" : "history-ng"}`; badge.textContent = item.judgment || "-";
      top.append(time, badge);
      const details = document.createElement("div"); details.className = "history-detail-grid";
      appendHistoryDetail(details, "카드 LOT NO", item.lot, true);
      appendHistoryDetail(details, "사진 인식값", item.recognizedRaw || "인식되지 않음", true);
      appendHistoryDetail(details, "판정 사유", item.reason || "기록 없음", true);
      appendHistoryDetail(details, "재촬영 횟수", `${Number(item.retryCount || 0)}회`);
      appendHistoryDetail(details, "작업 시작", item.startedAtIso ? new Date(item.startedAtIso).toLocaleString("ko-KR", { hour12: false }) : "기록 없음", true);
      appendHistoryDetail(details, "작업 소요시간", `${Number(item.elapsedSeconds || 0)}초`);
      card.append(top, details); list.appendChild(card);
    });
    group.append(header, list); elements.historyGroups.appendChild(group);
  });
}

function finalizeJudgment() {
  if (!pendingJudgment || judgmentFinalizing) return;
  judgmentFinalizing = true; elements.confirmJudgmentButton.disabled = true;
  const now = new Date(); const started = new Date(pendingJudgment.startedAtIso || now.toISOString());
  const record = {
    id: `${now.toISOString()}-${Math.random().toString(16).slice(2)}`,
    completedAtIso: now.toISOString(), completedAt: now.toLocaleString("ko-KR", { hour12: false }),
    startedAtIso: pendingJudgment.startedAtIso || now.toISOString(),
    elapsedSeconds: Math.max(0, Math.round((now.getTime() - started.getTime()) / 1000)),
    judgment: pendingJudgment.judgment, reason: pendingJudgment.reason, lot: pendingJudgment.lot,
    expected: pendingJudgment.expected, recognizedRaw: pendingJudgment.recognizedRaw,
    recognizedNormalized: pendingJudgment.recognizedNormalized, retryCount: pendingJudgment.retryCount,
  };
  if (!saveHistoryRecord(record)) {
    elements.judgmentMessage.textContent = "이력을 저장하지 못했습니다. 현재 LOT를 삭제하지 않았습니다. 저장공간을 확인한 뒤 다시 눌러주세요.";
    elements.confirmJudgmentButton.disabled = false; judgmentFinalizing = false; return;
  }
  const result = pendingJudgment.judgment;
  closeDialog(elements.judgmentDialog); elements.confirmJudgmentButton.disabled = false; judgmentFinalizing = false;
  updateHistoryBadge(); resetCurrentWork();
  showToast(result === "OK" ? "OK 이력 저장 및 작업 초기화 완료" : "NG 이력 저장 및 작업 초기화 완료");
}

function openHistoryPasswordDialog() {
  elements.historyPasswordInput.value = ""; elements.historyPasswordMessage.textContent = ""; openDialog(elements.historyPasswordDialog);
}
function unlockHistory(event) {
  event.preventDefault();
  if (elements.historyPasswordInput.value !== getHistoryPassword()) { elements.historyPasswordMessage.textContent = "비밀번호가 맞지 않습니다."; return; }
  closeDialog(elements.historyPasswordDialog); elements.historyPasswordInput.value = ""; renderHistory(); openDialog(elements.historyDialog);
}
function openPasswordSettings(returnTarget) {
  passwordSettingsReturnTarget = returnTarget; closeDialog(elements.historyPasswordDialog); closeDialog(elements.historyDialog);
  elements.masterPasswordInput.value = ""; elements.newHistoryPasswordInput.value = ""; elements.confirmHistoryPasswordInput.value = ""; elements.passwordSettingsMessage.textContent = "";
  openDialog(elements.passwordSettingsDialog);
}
function closePasswordSettingsAndReturn() {
  const target = passwordSettingsReturnTarget; passwordSettingsReturnTarget = "none"; closeDialog(elements.passwordSettingsDialog);
  if (target === "history") { renderHistory(); openDialog(elements.historyDialog); }
  else if (target === "unlock") openHistoryPasswordDialog();
}
function changeHistoryPassword(event) {
  event.preventDefault();
  const master = elements.masterPasswordInput.value; const password = elements.newHistoryPasswordInput.value; const confirm = elements.confirmHistoryPasswordInput.value;
  if (master !== MASTER_PASSWORD) { elements.passwordSettingsMessage.textContent = "마스터 비밀번호가 맞지 않습니다."; return; }
  if (!/^\d{4,12}$/.test(password)) { elements.passwordSettingsMessage.textContent = "새 비밀번호는 숫자 4~12자리로 입력하세요."; return; }
  if (password !== confirm) { elements.passwordSettingsMessage.textContent = "새 비밀번호와 확인값이 다릅니다."; return; }
  if (!saveHistoryPassword(password)) { elements.passwordSettingsMessage.textContent = "비밀번호를 저장하지 못했습니다."; return; }
  closePasswordSettingsAndReturn(); showToast("작업이력 비밀번호를 변경했습니다.");
}
function openAppSettings() { elements.autoResetMinutesSelect.value = String(getAutoResetMinutes()); openDialog(elements.appSettingsDialog); }
function saveAppSettings(event) {
  event.preventDefault(); const minutes = Number(elements.autoResetMinutesSelect.value);
  if (!saveAutoResetMinutes(minutes)) { showToast("자동 초기화 설정을 저장하지 못했습니다."); return; }
  closeDialog(elements.appSettingsDialog); if (currentLot) startAutoResetTimer(currentLot, true);
  showToast(`자동 초기화: ${formatAutoResetMinutes(minutes)}`);
}

window.onAndroidLotScanned = function (value) {
  scanInProgress = false; elements.lotNo.value = normalizeValue(value); updateLot();
  if (elements.lotMessage.textContent) { elements.lotCameraScanButton.focus(); showToast("인식한 TLB LOT NO를 확인해 주세요."); }
  else showToast("TLB LOT NO를 카메라로 입력했습니다.");
};
window.onAndroidScanCancelled = function () { scanInProgress = false; showToast("LOT NO 스캔을 취소했습니다."); };
window.onPanelOcrCompleted = function (rawText, normalizedText, matched) {
  verificationInProgress = false; setWorkControlsDisabled(false);
  showJudgmentDialog(matched ? "OK" : "NG", rawText, normalizedText, matched ? "카드 LOT와 사진 LOT가 일치합니다." : "카드 LOT와 사진 인식값이 일치하지 않습니다.");
};
window.onPanelOcrFailed = function (message) {
  verificationInProgress = false; setWorkControlsDisabled(false);
  showJudgmentDialog("NG", "", "", `문자 인식 실패: ${message || "사진을 확인해 주세요."}`);
};
window.onPanelPhotoCancelled = function () {
  verificationInProgress = false; setWorkControlsDisabled(false); elements.verificationStatus.textContent = "촬영이 취소되었습니다. 다시 시도할 수 있습니다.";
};

elements.lotNo.addEventListener("input", updateLot);
elements.lotNo.addEventListener("keydown", (event) => { if (event.key === "Enter") { event.preventDefault(); updateLot(); } });
elements.clearLotButton.addEventListener("click", resetCurrentWork);
if (isAndroidCameraAvailable()) {
  elements.lotNo.readOnly = true; elements.lotNo.classList.add("scan-locked"); elements.lotNo.setAttribute("aria-readonly", "true");
  elements.lotCameraScanButton.classList.remove("is-hidden");
  elements.lotCameraScanButton.addEventListener("click", () => {
    if (scanInProgress) return;
    try { scanInProgress = true; window.AndroidScanner.scanLotBarcode(); }
    catch (_error) { scanInProgress = false; showToast("카메라를 열지 못했습니다. 앱 권한을 확인해 주세요."); }
  });
}
elements.verifyPanelButton.addEventListener("click", beginPanelVerification);
elements.retryPhotoButton.addEventListener("click", () => {
  closeDialog(elements.judgmentDialog); pendingJudgment = null; retryCount += 1;
  if (currentLot) startAutoResetTimer(currentLot, true);
  window.setTimeout(beginPanelVerification, 120);
});
elements.confirmJudgmentButton.addEventListener("click", finalizeJudgment);
elements.judgmentDialog.addEventListener("cancel", (event) => event.preventDefault());
elements.appSettingsButton.addEventListener("click", openAppSettings);
elements.appSettingsForm.addEventListener("submit", saveAppSettings);
elements.cancelAppSettingsButton.addEventListener("click", () => closeDialog(elements.appSettingsDialog));
elements.progressButton.addEventListener("click", openHistoryPasswordDialog);
elements.closeHistoryButton.addEventListener("click", () => closeDialog(elements.historyDialog));
elements.historyPasswordForm.addEventListener("submit", unlockHistory);
elements.cancelHistoryPasswordButton.addEventListener("click", () => closeDialog(elements.historyPasswordDialog));
elements.openPasswordSettingsButton.addEventListener("click", () => openPasswordSettings("unlock"));
elements.changeHistoryPasswordButton.addEventListener("click", () => openPasswordSettings("history"));
elements.passwordSettingsForm.addEventListener("submit", changeHistoryPassword);
elements.cancelPasswordSettingsButton.addEventListener("click", closePasswordSettingsAndReturn);
elements.historyLotSearch.addEventListener("input", renderHistory);
elements.historyDateFilter.addEventListener("change", renderHistory);
elements.clearHistoryFiltersButton.addEventListener("click", () => { elements.historyLotSearch.value = ""; elements.historyDateFilter.value = ""; renderHistory(); });
elements.passwordSettingsDialog.addEventListener("cancel", (event) => { event.preventDefault(); closePasswordSettingsAndReturn(); });

if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
  window.addEventListener("load", () => navigator.serviceWorker.register("service-worker.js").catch(() => {}));
}
elements.historyDateFilter.max = formatDateKey(new Date());
updateHistoryBadge(); updateLot(); focusLotStep();
