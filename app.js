(function () {
  "use strict";

  const DATA_PATH = "data/names.xlsx";
  const SHEET_NAME = "判定アプリ用";

  const inputScreen = document.getElementById("inputScreen");
  const resultScreen = document.getElementById("resultScreen");
  const judgeForm = document.getElementById("judgeForm");
  const nameInput = document.getElementById("nameInput");
  const inputError = document.getElementById("inputError");
  const judgeButton = document.getElementById("judgeButton");
  const loadMessage = document.getElementById("loadMessage");
  const resultReading = document.getElementById("resultReading");
  const resultStatus = document.getElementById("resultStatus");
  const resultNames = document.getElementById("resultNames");
  const backButton = document.getElementById("backButton");

  const readings = new Map();
  let dataReady = false;

  function normalizeReading(value) {
    return String(value || "")
      .trim()
      .replace(/[ 　]/g, "")
      .replace(/[\u30a1-\u30f6]/g, function (char) {
        return String.fromCharCode(char.charCodeAt(0) - 0x60);
      });
  }

  function setLoadingState(message, isError) {
    loadMessage.textContent = message;
    loadMessage.style.color = isError ? "var(--fail)" : "";
    judgeButton.disabled = !dataReady;
  }

  async function loadWorkbook() {
    try {
      const cacheBuster = new Date().getTime();
      const response = await fetch(`${DATA_PATH}?v=${cacheBuster}`, {
        cache: "no-store",
      });

      if (!response.ok) {
        throw new Error("Excelファイルを読み込めませんでした。");
      }

      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const worksheet = workbook.Sheets[SHEET_NAME];

      if (!worksheet) {
        throw new Error(`「${SHEET_NAME}」シートが見つかりません。`);
      }

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: ["よみ", "表示人物名"],
        range: 1,
        blankrows: false,
        defval: "",
      });

      rows.forEach(function (row) {
        const reading = normalizeReading(row["よみ"]);
        const displayName = String(row["表示人物名"]);

        if (!reading || !displayName) {
          return;
        }

        if (!readings.has(reading)) {
          readings.set(reading, []);
        }

        readings.get(reading).push(displayName);
      });

      dataReady = true;
      setLoadingState("", false);
    } catch (error) {
      dataReady = false;
      setLoadingState(error.message || "判定データの読み込みに失敗しました。", true);
    }
  }

  function showInputScreen() {
    resultScreen.classList.add("hidden");
    inputScreen.classList.remove("hidden");
    inputError.textContent = "";
    nameInput.focus();
  }

  function showResultScreen(rawInput, matchedNames) {
    inputScreen.classList.add("hidden");
    resultScreen.classList.remove("hidden");

    resultReading.textContent = normalizeReading(rawInput);
    resultStatus.classList.remove("success", "fail");

    if (matchedNames && matchedNames.length > 0) {
      resultStatus.textContent = "成功！";
      resultStatus.classList.add("success");
      resultNames.textContent = matchedNames.join("／");
    } else {
      resultStatus.textContent = "失敗…";
      resultStatus.classList.add("fail");
      resultNames.textContent = "";
    }
  }

  judgeForm.addEventListener("submit", function (event) {
    event.preventDefault();

    inputError.textContent = "";

    if (!dataReady) {
      inputError.textContent = "判定データを読み込み中です";
      return;
    }

    const rawInput = nameInput.value;
    const normalizedInput = normalizeReading(rawInput);

    if (!normalizedInput) {
      inputError.textContent = "人物名を入力してください";
      return;
    }

    showResultScreen(rawInput, readings.get(normalizedInput));
  });

  backButton.addEventListener("click", function () {
    showInputScreen();
  });

  window.addEventListener("DOMContentLoaded", function () {
    judgeButton.disabled = true;
    loadWorkbook();
  });
})();
