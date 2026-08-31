"use strict";

const RULE = Object.freeze({
  managementLength: 10,
  managementPrefixLength: 7,
  managementCharacterIndex: 7,
  lotLength: 14,
  lotMiddleStart: 8,
  lotMiddleLength: 4,
  allowedLotMiddleValues: Object.freeze(["0000", "0108", "0916"]),
  manualLength: 1,
});

const HISTORY_KEY = "kccHdiLaserInspectionHistoryV1";

const elements = {
  managementNo: document.getElementById("managementNo"),
  lotNo: document.getElementById("lotNo"),
  managementCameraScanButton: document.getElementById("managementCameraScanButton"),
  lotCameraScanButton: document.getElementById("lotCameraScanButton"),
  manualValue: document.getElementById("manualValue"),
  managementMessage: document.getElementById("managementMessage"),
  lotMessage: document.getElementById("lotMessage"),
  manualMessage: document.getElementById("manualMessage"),
  resultValue: document.getElementById("resultValue"),
  formulaPreview: document.getElementById("formulaPreview"),
  codeArea: document.getElementById("codeArea"),
  errorBanner: document.getElementById("errorBanner"),
  copyButton: document.getElementById("copyButton"),
  barcodeCanvas: document.getElementById("barcodeCanvas"),
  qrCanvas: document.getElementById("qrCanvas"),
  verificationArea: document.getElementById("verificationArea"),
  verificationStatus: document.getElementById("verificationStatus"),
  verifyPanelButton: document.getElementById("verifyPanelButton"),
  progressButton: document.getElementById("progressButton"),
  historyCountBadge: document.getElementById("historyCountBadge"),
  historyDialog: document.getElementById("historyDialog"),
  closeHistoryButton: document.getElementById("closeHistoryButton"),
  historyTotalCount: document.getElementById("historyTotalCount"),
  historyOkCount: document.getElementById("historyOkCount"),
  historyNgCount: document.getElementById("historyNgCount"),
  historyEmpty: document.getElementById("historyEmpty"),
  historyTableWrap: document.getElementById("historyTableWrap"),
  historyTableBody: document.getElementById("historyTableBody"),
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

const inputs = [elements.managementNo, elements.lotNo, elements.manualValue];
let currentResult = "";
let toastTimer = 0;
let scanInProgress = false;
let verificationInProgress = false;
let pendingVerificationExpected = "";
let pendingJudgment = null;
let judgmentFinalizing = false;

function normalizeScannedValue(value) {
  return String(value || "").trim().replace(/[\r\n\t]/g, "").toUpperCase();
}

function validate() {
  const management = normalizeScannedValue(elements.managementNo.value);
  const lot = normalizeScannedValue(elements.lotNo.value);
  const manual = normalizeScannedValue(elements.manualValue.value);
  const errors = { management: "", lot: "", manual: "" };

  if (management && management.length !== RULE.managementLength) {
    errors.management = `관리번호는 정확히 ${RULE.managementLength}자리여야 합니다. (현재 ${management.length}자리)`;
  } else if (management && !/^[A-Z0-9]+$/.test(management)) {
    errors.management = "관리번호에는 영문과 숫자만 입력할 수 있습니다.";
  } else if (management && !/^[A-Z]$/.test(management.charAt(RULE.managementCharacterIndex))) {
    errors.management = "관리번호 8번째 자리는 반드시 영문 A~Z여야 합니다.";
  }

  if (lot && lot.length !== RULE.lotLength) {
    errors.lot = `LOT NO는 정확히 ${RULE.lotLength}자리여야 합니다. (현재 ${lot.length}자리)`;
  } else if (lot && !/^[A-Z0-9]+$/.test(lot)) {
    errors.lot = "LOT NO에는 영문과 숫자만 입력할 수 있습니다.";
  } else if (lot) {
    const lotMiddle = lot.slice(RULE.lotMiddleStart, RULE.lotMiddleStart + RULE.lotMiddleLength);
    if (!RULE.allowedLotMiddleValues.includes(lotMiddle)) {
      errors.lot = `LOT 중간번호 ${lotMiddle}은 사용할 수 없습니다. 허용값: 0000, 0108, 0916`;
    }
  }

  if (manual && !/^[A-Z0-9]$/.test(manual)) {
    errors.manual = "수기 입력은 영문 또는 숫자 1자리만 가능합니다.";
  }

  return { management, lot, manual, errors };
}

function setFieldState(input, messageElement, cardId, error) {
  const card = document.getElementById(cardId);
  const hasValue = input.value.trim() !== "";
  input.setAttribute("aria-invalid", error ? "true" : "false");
  messageElement.textContent = error;
  card.classList.toggle("invalid", Boolean(error));
  card.classList.toggle("valid", hasValue && !error);
}

function clearCodes() {
  currentResult = "";
  elements.codeArea.classList.add("is-hidden");
  elements.codeArea.setAttribute("aria-hidden", "true");
  elements.verificationArea.classList.add("is-hidden");
  elements.verificationArea.setAttribute("aria-hidden", "true");
  elements.copyButton.disabled = true;
  elements.verifyPanelButton.disabled = true;
  elements.barcodeCanvas.getContext("2d").clearRect(0, 0, elements.barcodeCanvas.width, elements.barcodeCanvas.height);
  elements.qrCanvas.getContext("2d").clearRect(0, 0, elements.qrCanvas.width, elements.qrCanvas.height);
}

function clearAllInputValues() {
  inputs.forEach((input) => {
    input.value = "";
    input.defaultValue = "";
    input.setAttribute("value", "");
  });
}

function showGenerationError(error) {
  clearCodes();
  elements.errorBanner.textContent = `코드 생성 오류: ${error.message || error}`;
  elements.errorBanner.classList.remove("is-hidden");
}

function renderCodes(value) {
  if (typeof window.bwipjs === "undefined") {
    throw new Error("바코드 생성 파일을 불러오지 못했습니다.");
  }

  window.bwipjs.toCanvas(elements.barcodeCanvas, {
    bcid: "code128",
    text: value,
    scale: 4,
    height: 24,
    includetext: true,
    textxalign: "center",
    textsize: 12,
    paddingwidth: 8,
    paddingheight: 8,
    backgroundcolor: "FFFFFF",
  });

  window.bwipjs.toCanvas(elements.qrCanvas, {
    bcid: "qrcode",
    text: value,
    scale: 8,
    eclevel: "M",
    paddingwidth: 4,
    paddingheight: 4,
    backgroundcolor: "FFFFFF",
  });
}

function updateResult() {
  const data = validate();
  setFieldState(elements.managementNo, elements.managementMessage, "card-management", data.errors.management);
  setFieldState(elements.lotNo, elements.lotMessage, "card-lot", data.errors.lot);
  setFieldState(elements.manualValue, elements.manualMessage, "card-manual", data.errors.manual);
  elements.errorBanner.classList.add("is-hidden");

  const hasErrors = Object.values(data.errors).some(Boolean);
  const isComplete = Boolean(data.management && data.lot && data.manual);

  if (!isComplete || hasErrors) {
    clearCodes();
    elements.resultValue.textContent = hasErrors ? "입력값을 다시 확인하세요" : "입력값을 모두 채워주세요";
    elements.resultValue.classList.add("empty");
    elements.formulaPreview.textContent = "관리번호 앞 7자리 + LOT 9~12번째 4자리 + 관리번호 8번째 + 수기 영문/숫자 1자리";
    return;
  }

  const prefix = data.management.slice(0, RULE.managementPrefixLength);
  const lotMiddle = data.lot.slice(RULE.lotMiddleStart, RULE.lotMiddleStart + RULE.lotMiddleLength);
  const managementEighth = data.management.charAt(RULE.managementCharacterIndex);
  const result = `${prefix}${lotMiddle}${managementEighth}${data.manual}`;

  try {
    renderCodes(result);
    currentResult = result;
    elements.resultValue.textContent = result;
    elements.resultValue.classList.remove("empty");
    elements.formulaPreview.textContent = `${prefix} + ${lotMiddle} + ${managementEighth} + ${data.manual}`;
    elements.codeArea.classList.remove("is-hidden");
    elements.codeArea.setAttribute("aria-hidden", "false");
    elements.verificationArea.classList.remove("is-hidden");
    elements.verificationArea.setAttribute("aria-hidden", "false");
    elements.copyButton.disabled = false;
    elements.verifyPanelButton.disabled = !isPanelVerificationAvailable();
    elements.verificationStatus.textContent = isPanelVerificationAvailable()
      ? "초도품 사진 촬영 대기"
      : "초도품 사진 대조는 Android 앱에서 사용합니다.";
  } catch (error) {
    showGenerationError(error);
  }
}

function resetCurrentWork() {
  verificationInProgress = false;
  pendingVerificationExpected = "";
  pendingJudgment = null;
  setWorkControlsDisabled(false);
  clearAllInputValues();
  updateResult();
  focusFirstStep();
}

function showToast(message) {
  window.clearTimeout(toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  toastTimer = window.setTimeout(() => elements.toast.classList.remove("show"), 2200);
}

function isAndroidCameraAvailable() {
  return Boolean(
    window.AndroidScanner
    && typeof window.AndroidScanner.scanManagementBarcode === "function"
    && typeof window.AndroidScanner.scanLotBarcode === "function"
  );
}

function isPanelVerificationAvailable() {
  return Boolean(
    window.AndroidScanner
    && typeof window.AndroidScanner.captureAndVerifyPanel === "function"
  );
}

function focusFirstStep() {
  if (isAndroidCameraAvailable()) {
    elements.managementCameraScanButton.focus();
  } else {
    elements.managementNo.focus();
  }
}

function setWorkControlsDisabled(disabled) {
  document.querySelectorAll(".camera-scan, .clear-field").forEach((button) => {
    button.disabled = disabled;
  });
  elements.manualValue.disabled = disabled;
  elements.verifyPanelButton.disabled = disabled || !currentResult || !isPanelVerificationAvailable();
}

function beginPanelVerification() {
  if (!currentResult) {
    showToast("먼저 최종 Laser 각인값을 생성해 주세요.");
    return;
  }
  if (!isPanelVerificationAvailable()) {
    showToast("초도품 사진 대조는 Android 앱에서만 사용할 수 있습니다.");
    return;
  }
  if (verificationInProgress) return;

  verificationInProgress = true;
  pendingVerificationExpected = currentResult;
  setWorkControlsDisabled(true);
  elements.verificationStatus.textContent = "카메라 촬영 및 문자 인식 진행 중...";

  try {
    window.AndroidScanner.captureAndVerifyPanel(currentResult);
  } catch (_error) {
    verificationInProgress = false;
    setWorkControlsDisabled(false);
    elements.verificationStatus.textContent = "카메라를 열지 못했습니다. 다시 시도해 주세요.";
    showToast("초도품 촬영 카메라를 열지 못했습니다.");
  }
}

function openDialog(dialog) {
  if (typeof dialog.showModal === "function") {
    if (!dialog.open) dialog.showModal();
  } else {
    dialog.setAttribute("open", "");
  }
}

function closeDialog(dialog) {
  if (typeof dialog.close === "function" && dialog.open) {
    dialog.close();
  } else {
    dialog.removeAttribute("open");
  }
}

function createPendingJudgment(judgment, rawText, normalizedText, reason) {
  const data = validate();
  return {
    judgment,
    reason,
    expected: pendingVerificationExpected || currentResult,
    recognizedRaw: String(rawText || "").replace(/\s+/g, " ").trim(),
    recognizedNormalized: String(normalizedText || ""),
    management: data.management,
    lot: data.lot,
    manual: data.manual,
  };
}

function showJudgmentDialog(judgment, rawText, normalizedText, reason) {
  pendingJudgment = createPendingJudgment(judgment, rawText, normalizedText, reason);
  const isOk = judgment === "OK";
  elements.judgmentDialog.classList.toggle("ok-dialog", isOk);
  elements.judgmentDialog.classList.toggle("ng-dialog", !isOk);
  elements.judgmentMark.textContent = isOk ? "OK" : "NG";
  elements.judgmentTitle.textContent = isOk ? "초도품 각인 대조 OK" : "조장 확인 필요!!";
  elements.judgmentMessage.textContent = isOk
    ? "생성 각인값과 사진의 각인 문자가 정확히 일치합니다. 확인을 누르면 이력을 저장하고 현재 작업을 삭제합니다."
    : `${reason} 사진 문제일 수 있으므로 다시 찍어 확인할 수 있습니다.`;
  elements.judgmentExpected.textContent = pendingJudgment.expected || "-";
  elements.judgmentRecognized.textContent = pendingJudgment.recognizedRaw || "인식되지 않음";
  elements.retryPhotoButton.classList.toggle("is-hidden", isOk);
  elements.confirmJudgmentButton.textContent = isOk ? "확인 및 작업 완료" : "NG 확정 · 조장 확인";
  elements.verificationStatus.textContent = isOk ? "사진 대조 OK" : "사진 대조 NG · 재촬영 또는 조장 확인";
  openDialog(elements.judgmentDialog);
}

function loadHistory() {
  try {
    const parsed = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function saveHistoryRecord(record) {
  try {
    const history = loadHistory();
    history.unshift(record);
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
    return true;
  } catch (_error) {
    return false;
  }
}

function appendCell(row, value, className = "") {
  const cell = document.createElement("td");
  cell.textContent = value || "-";
  if (className) cell.className = className;
  row.appendChild(cell);
}

function renderHistory() {
  const history = loadHistory();
  const okCount = history.filter((item) => item.judgment === "OK").length;
  const ngCount = history.filter((item) => item.judgment === "NG").length;
  elements.historyCountBadge.textContent = String(history.length);
  elements.historyTotalCount.textContent = String(history.length);
  elements.historyOkCount.textContent = String(okCount);
  elements.historyNgCount.textContent = String(ngCount);
  elements.historyEmpty.classList.toggle("is-hidden", history.length > 0);
  elements.historyTableWrap.classList.toggle("is-hidden", history.length === 0);
  elements.historyTableBody.replaceChildren();

  history.forEach((item) => {
    const row = document.createElement("tr");
    appendCell(row, item.completedAt);
    appendCell(row, item.judgment, item.judgment === "OK" ? "history-ok" : "history-ng");
    appendCell(row, item.expected, "mono-cell");
    appendCell(row, item.recognizedRaw || item.reason, "mono-cell");
    appendCell(row, item.management, "mono-cell");
    appendCell(row, item.lot, "mono-cell");
    elements.historyTableBody.appendChild(row);
  });
}

function finalizeJudgment() {
  if (!pendingJudgment || judgmentFinalizing) return;
  judgmentFinalizing = true;
  elements.confirmJudgmentButton.disabled = true;

  const now = new Date();
  const record = {
    id: `${now.toISOString()}-${Math.random().toString(16).slice(2)}`,
    completedAt: now.toLocaleString("ko-KR", { hour12: false }),
    judgment: pendingJudgment.judgment,
    reason: pendingJudgment.reason,
    expected: pendingJudgment.expected,
    recognizedRaw: pendingJudgment.recognizedRaw,
    recognizedNormalized: pendingJudgment.recognizedNormalized,
    management: pendingJudgment.management,
    lot: pendingJudgment.lot,
    manual: pendingJudgment.manual,
  };

  if (!saveHistoryRecord(record)) {
    elements.judgmentMessage.textContent = "이력을 저장하지 못했습니다. 작업을 삭제하지 않았습니다. 태블릿 저장공간을 확인한 뒤 다시 눌러주세요.";
    elements.confirmJudgmentButton.disabled = false;
    judgmentFinalizing = false;
    return;
  }

  const completedJudgment = pendingJudgment.judgment;
  closeDialog(elements.judgmentDialog);
  elements.confirmJudgmentButton.disabled = false;
  judgmentFinalizing = false;
  renderHistory();
  resetCurrentWork();
  showToast(completedJudgment === "OK" ? "OK 이력 저장 및 작업 삭제 완료" : "NG 이력 저장 및 작업 삭제 완료");
}

window.onAndroidManagementScanned = function onAndroidManagementScanned(scannedValue) {
  scanInProgress = false;
  const normalized = normalizeScannedValue(scannedValue);
  if (!normalized) {
    showToast("관리번호 바코드 값을 읽지 못했습니다. 다시 촬영해 주세요.");
    return;
  }
  elements.managementNo.value = normalized;
  updateResult();
  if (elements.managementMessage.textContent) {
    elements.managementCameraScanButton.focus();
    showToast("인식한 관리번호를 확인해 주세요.");
  } else {
    elements.lotCameraScanButton.focus();
    showToast("관리번호를 카메라로 입력했습니다.");
  }
};

window.onAndroidLotScanned = function onAndroidLotScanned(scannedValue) {
  scanInProgress = false;
  const normalized = normalizeScannedValue(scannedValue);
  if (!normalized) {
    showToast("LOT NO 바코드 값을 읽지 못했습니다. 다시 촬영해 주세요.");
    return;
  }
  elements.lotNo.value = normalized;
  updateResult();
  if (elements.lotMessage.textContent) {
    elements.lotCameraScanButton.focus();
    showToast("인식한 LOT NO를 확인해 주세요.");
  } else {
    elements.manualValue.focus();
    showToast("LOT NO를 카메라로 입력했습니다.");
  }
};

window.onAndroidScanCancelled = function onAndroidScanCancelled(scanTarget) {
  scanInProgress = false;
  showToast("바코드 스캔을 취소했습니다.");
  if (scanTarget === "management") elements.managementCameraScanButton.focus();
  else elements.lotCameraScanButton.focus();
};

window.onPanelOcrCompleted = function onPanelOcrCompleted(rawText, normalizedText, matched) {
  verificationInProgress = false;
  setWorkControlsDisabled(false);
  const stillSameWork = Boolean(currentResult && currentResult === pendingVerificationExpected);
  const exactMatch = Boolean(matched && stillSameWork && String(normalizedText || "").includes(pendingVerificationExpected));
  showJudgmentDialog(
    exactMatch ? "OK" : "NG",
    rawText,
    normalizedText,
    exactMatch ? "완전 일치" : "생성 각인값과 사진 인식값이 일치하지 않습니다."
  );
};

window.onPanelOcrFailed = function onPanelOcrFailed(message) {
  verificationInProgress = false;
  setWorkControlsDisabled(false);
  showJudgmentDialog("NG", "", "", message || "각인 문자를 인식하지 못했습니다.");
};

window.onPanelPhotoCancelled = function onPanelPhotoCancelled() {
  verificationInProgress = false;
  setWorkControlsDisabled(false);
  elements.verificationStatus.textContent = "촬영이 취소되었습니다. 초도품을 다시 촬영해 주세요.";
  showToast("초도품 사진 촬영을 취소했습니다.");
};

inputs.forEach((input) => {
  input.addEventListener("input", () => {
    const normalized = input === elements.manualValue
      ? normalizeScannedValue(input.value).replace(/[^A-Z0-9]/g, "").slice(0, RULE.manualLength)
      : normalizeScannedValue(input.value);
    if (input.value !== normalized) input.value = normalized;
    updateResult();
  });

  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      updateResult();
    }
  });
});

document.querySelectorAll(".clear-field").forEach((button) => {
  button.addEventListener("click", () => {
    const target = document.getElementById(button.dataset.target);
    target.value = "";
    updateResult();
    target.focus();
  });
});

if (isAndroidCameraAvailable()) {
  [elements.managementNo, elements.lotNo].forEach((input) => {
    input.readOnly = true;
    input.classList.add("scan-locked");
    input.setAttribute("aria-readonly", "true");
  });

  elements.managementCameraScanButton.classList.remove("is-hidden");
  elements.lotCameraScanButton.classList.remove("is-hidden");

  elements.managementCameraScanButton.addEventListener("click", () => {
    if (scanInProgress) return;
    try {
      scanInProgress = true;
      window.AndroidScanner.scanManagementBarcode();
    } catch (_error) {
      scanInProgress = false;
      showToast("카메라를 열지 못했습니다. 앱 권한을 확인해 주세요.");
    }
  });

  elements.lotCameraScanButton.addEventListener("click", () => {
    if (scanInProgress) return;
    try {
      scanInProgress = true;
      window.AndroidScanner.scanLotBarcode();
    } catch (_error) {
      scanInProgress = false;
      showToast("카메라를 열지 못했습니다. 앱 권한을 확인해 주세요.");
    }
  });
}

elements.verifyPanelButton.addEventListener("click", beginPanelVerification);

elements.retryPhotoButton.addEventListener("click", () => {
  closeDialog(elements.judgmentDialog);
  pendingJudgment = null;
  window.setTimeout(beginPanelVerification, 120);
});

elements.confirmJudgmentButton.addEventListener("click", finalizeJudgment);

elements.judgmentDialog.addEventListener("cancel", (event) => {
  event.preventDefault();
});

elements.progressButton.addEventListener("click", () => {
  renderHistory();
  openDialog(elements.historyDialog);
});

elements.closeHistoryButton.addEventListener("click", () => closeDialog(elements.historyDialog));

elements.copyButton.addEventListener("click", async () => {
  if (!currentResult) return;
  try {
    await navigator.clipboard.writeText(currentResult);
    showToast("최종 Laser 각인값을 복사했습니다.");
  } catch (_error) {
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(elements.resultValue);
    selection.removeAllRanges();
    selection.addRange(range);
    showToast("값을 선택했습니다. Ctrl+C를 눌러 복사하세요.");
  }
});

if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
  window.addEventListener("load", () => navigator.serviceWorker.register("service-worker.js").catch(() => {}));
}

renderHistory();
updateResult();
focusFirstStep();
