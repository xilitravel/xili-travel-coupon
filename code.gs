/*****************************************************************
 *
 * 🌸 囍莉旅遊｜電子優惠券系統
 * 完整版
 *
 * 功能：
 * ✔ 工作人員後台
 * ✔ 客人優惠券
 * ✔ QR Code
 * ✔ 相機掃描核銷
 * ✔ 手動輸入核銷
 * ✔ Enter 快速核銷
 * ✔ PIN 登入
 * ✔ 發券後 +180 天
 * ✔ 新增部署後自動使用最新網址
 * ✔ 一次建立取消 100 張限制
 * ✔ 永久記錄歷史券號
 * ✔ 刪除已使用優惠券後不會產生重複券號
 *
 *****************************************************************/


// ============================================================
// ⚙️ 基本設定
// ============================================================


// 工作人員 PIN
const STAFF_PIN = "1234";


// 優惠券主要工作表名稱
const SHEET_NAME = "優惠券";


// 優惠券序號永久歷史紀錄工作表
const HISTORY_SHEET_NAME = "優惠券序號歷史";


// 優惠券有效天數
const VALID_DAYS = 180;






// ============================================================
// 🌐 網頁入口
// ============================================================

function doGet(e) {


  // ==========================================================
  // 取得網址參數
  // ==========================================================

  const page =

    (
      e &&
      e.parameter &&
      e.parameter.page
    )

      ? String(
          e.parameter.page
        )
        .trim()
        .toLowerCase()

      : "index";



  // ==========================================================
  // 🎫 客人優惠券頁面
  //
  // 網址：
  // ?page=coupon&coupon=XL-XXXX-XXXX
  // ==========================================================

  if (
    page === "coupon"
  ) {


    const couponCode =

      (
        e &&
        e.parameter &&
        e.parameter.coupon
      )

        ? String(
            e.parameter.coupon
          )
          .trim()

        : "";



    const template =

      HtmlService
        .createTemplateFromFile(
          "Coupon"
        );



    // 把優惠券序號傳給 Coupon.html
    template.couponCode =
      couponCode;



    return template

      .evaluate()

      .setTitle(
        "囍莉旅遊｜電子優惠券"
      )

      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }



  // ==========================================================
  // 👩‍💼 工作人員後台
  //
  // 支援：
  // ?page=backend
  // ?page=index
  //
  // 或直接開 Web App 網址
  // ==========================================================

  const template =

    HtmlService
      .createTemplateFromFile(
        "Index"
      );



  return template

    .evaluate()

    .setTitle(
      "囍莉旅遊｜工作人員後台"
    )

    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}






// ============================================================
// 🌐 取得目前 Web App 網址
// ============================================================

function getWebAppUrl() {


  const url =

    ScriptApp
      .getService()
      .getUrl();



  if (!url) {

    throw new Error(
      "無法取得 Web App 網址，請確認 Apps Script 已完成網頁應用程式部署"
    );

  }



  return url;

}






// ============================================================
// 📊 取得優惠券主要工作表
// ============================================================

function getSheet() {


  const spreadsheet =

    SpreadsheetApp
      .getActiveSpreadsheet();



  let sheet =

    spreadsheet
      .getSheetByName(
        SHEET_NAME
      );



  // ==========================================================
  // 如果沒有工作表，自動建立
  // ==========================================================

  if (!sheet) {

    sheet =

      spreadsheet
        .insertSheet(
          SHEET_NAME
        );

  }



  // ==========================================================
  // 如果沒有標題，自動建立
  // ==========================================================

  if (
    sheet.getLastRow() === 0
  ) {


    sheet.appendRow([

      "券號",

      "折扣金額",

      "狀態",

      "發行日期",

      "有效期限",

      "使用日期",

      "備註",

      "客人連結",

      "QR內容"

    ]);

  }



  return sheet;

}






// ============================================================
// 🔒 取得優惠券永久歷史紀錄工作表
//
// 這個工作表專門記錄所有曾經發過的券號
//
// 即使主優惠券工作表刪除資料
// 這裡仍然保留歷史券號
//
// 目的：
// 永遠避免產生相同優惠券序號
// ============================================================

function getHistorySheet() {


  const spreadsheet =

    SpreadsheetApp
      .getActiveSpreadsheet();



  let sheet =

    spreadsheet
      .getSheetByName(
        HISTORY_SHEET_NAME
      );



  // ==========================================================
  // 如果沒有歷史工作表
  // 自動建立
  // ==========================================================

  if (!sheet) {


    sheet =

      spreadsheet
        .insertSheet(
          HISTORY_SHEET_NAME
        );



    // 建立標題
    sheet.appendRow([

      "券號",

      "首次建立日期"

    ]);



    // 隱藏系統工作表
    sheet.hideSheet();

  }



  return sheet;

}






// ============================================================
// 🔄 將目前主優惠券工作表
// 自動同步到永久歷史紀錄
//
// 用途：
// 讓以前已經建立的優惠券
// 也自動加入永久歷史紀錄
// ============================================================

function syncExistingCouponsToHistory() {


  const couponSheet =

    getSheet();



  const historySheet =

    getHistorySheet();



  const couponLastRow =

    couponSheet.getLastRow();



  // ==========================================================
  // 主工作表沒有優惠券
  // ==========================================================

  if (
    couponLastRow < 2
  ) {

    return;

  }



  // ==========================================================
  // 取得目前所有優惠券
  // ==========================================================

  const couponCodes =

    couponSheet

      .getRange(

        2,

        1,

        couponLastRow - 1,

        1

      )

      .getValues()

      .flat()

      .map(

        value =>

          String(value)
            .trim()
            .toUpperCase()

      )

      .filter(

        value => value !== ""

      );



  // ==========================================================
  // 取得歷史券號
  // ==========================================================

  const historyLastRow =

    historySheet.getLastRow();



  let historyCodes = [];



  if (
    historyLastRow >= 2
  ) {


    historyCodes =

      historySheet

        .getRange(

          2,

          1,

          historyLastRow - 1,

          1

        )

        .getValues()

        .flat()

        .map(

          value =>

            String(value)
              .trim()
              .toUpperCase()

        );

  }



  // ==========================================================
  // 建立 Set
  // ==========================================================

  const historySet =

    new Set(
      historyCodes
    );



  const newHistoryRows = [];



  // ==========================================================
  // 找出還沒有加入歷史紀錄的券號
  // ==========================================================

  couponCodes.forEach(

    code => {


      if (
        !historySet.has(code)
      ) {


        newHistoryRows.push([

          code,

          new Date()

        ]);



        historySet.add(
          code
        );

      }

    }

  );



  // ==========================================================
  // 批次寫入歷史紀錄
  // ==========================================================

  if (
    newHistoryRows.length > 0
  ) {


    historySheet

      .getRange(

        historySheet.getLastRow() + 1,

        1,

        newHistoryRows.length,

        2

      )

      .setValues(
        newHistoryRows
      );



    historySheet

      .getRange(

        historySheet.getLastRow() -

        newHistoryRows.length +

        1,

        2,

        newHistoryRows.length,

        1

      )

      .setNumberFormat(
        "yyyy/MM/dd HH:mm:ss"
      );

  }

}






// ============================================================
// 🔐 工作人員 PIN 登入
// ============================================================

function loginStaff(pin) {


  if (

    String(pin)
      .trim()

    ===

    STAFF_PIN

  ) {


    return {

      success: true

    };

  }



  return {

    success: false,

    message:
      "PIN 錯誤"

  };

}






// ============================================================
// 🎟️ 建立優惠券
//
// ✔ 取消一次最多 100 張限制
// ✔ 自動檢查永久歷史券號
// ✔ 永遠避免重複券號
// ============================================================

function createCoupons(
  quantity,
  discount,
  note
) {


  // ==========================================================
  // 🔒 鎖定
  //
  // 避免多人同時建立優惠券
  // 造成重複券號
  // ==========================================================

  const lock =

    LockService
      .getScriptLock();



  lock.waitLock(
    30000
  );



  try {


    const sheet =

      getSheet();



    const historySheet =

      getHistorySheet();



    // ========================================================
    // 先同步以前已存在的優惠券
    // 到永久歷史紀錄
    // ========================================================

    syncExistingCouponsToHistory();



    quantity =

      Number(
        quantity
      );



    discount =

      Number(
        discount
      );



    note =

      String(
        note || ""
      )
      .trim();



    // ========================================================
    // 檢查數量
    // ========================================================

    if (

      !Number.isInteger(
        quantity
      )

      ||

      quantity < 1

    ) {

      throw new Error(
        "建立數量必須是大於 0 的整數"
      );

    }



    // ========================================================
    // 檢查折扣金額
    // ========================================================

    if (

      !discount ||

      discount < 1

    ) {

      throw new Error(
        "折扣金額必須大於 0"
      );

    }



    // ========================================================
    // 📅 發行日期
    // ========================================================

    const issueDate =

      new Date();



    // ========================================================
    // 📅 有效期限
    //
    // 發券日期 + 180 天
    // ========================================================

    const validUntil =

      new Date(
        issueDate
      );



    validUntil.setDate(

      validUntil.getDate() +

      VALID_DAYS

    );



    // ========================================================
    // 🌐 取得 Web App 網址
    // ========================================================

    const WEB_APP_URL =

      getWebAppUrl();



    // ========================================================
    // 取得永久歷史券號
    // ========================================================

    const historyLastRow =

      historySheet.getLastRow();



    let historyCodes = [];



    if (
      historyLastRow >= 2
    ) {


      historyCodes =

        historySheet

          .getRange(

            2,

            1,

            historyLastRow - 1,

            1

          )

          .getValues()

          .flat()

          .map(

            value =>

              String(value)
                .trim()
                .toUpperCase()

          )

          .filter(

            value => value !== ""

          );

    }



    // ========================================================
    // 取得目前主工作表券號
    //
    // 雙重檢查
    // ========================================================

    const sheetLastRow =

      sheet.getLastRow();



    let currentCodes = [];



    if (
      sheetLastRow >= 2
    ) {


      currentCodes =

        sheet

          .getRange(

            2,

            1,

            sheetLastRow - 1,

            1

          )

          .getValues()

          .flat()

          .map(

            value =>

              String(value)
                .trim()
                .toUpperCase()

          )

          .filter(

            value => value !== ""

          );

    }



    // ========================================================
    // 建立所有已使用券號 Set
    // ========================================================

    const usedCodes =

      new Set([

        ...historyCodes,

        ...currentCodes

      ]);



    // ========================================================
    // 準備批次資料
    // ========================================================

    const couponRows = [];



    const historyRows = [];



    const result = [];



    // ========================================================
    // 🎫 開始建立優惠券
    //
    // 不限制 100 張
    // ========================================================

    for (

      let i = 0;

      i < quantity;

      i++

    ) {


      // ======================================================
      // 產生唯一券號
      // ======================================================

      const code =

        generateUniqueCouponCode(
          usedCodes
        );



      // ======================================================
      // 立即加入 Set
      //
      // 避免這一批新券彼此重複
      // ======================================================

      usedCodes.add(
        code
      );



      // ======================================================
      // 客人優惠券網址
      // ======================================================

      const customerUrl =

        WEB_APP_URL +

        "?page=coupon&coupon=" +

        encodeURIComponent(
          code
        );



      // ======================================================
      // QR Code 內容
      // ======================================================

      const qrContent =

        customerUrl;



      // ======================================================
      // 主優惠券資料
      // ======================================================

      couponRows.push([

        code,

        discount,

        "未使用",

        issueDate,

        validUntil,

        "",

        note,

        customerUrl,

        qrContent

      ]);



      // ======================================================
      // 永久歷史紀錄
      // ======================================================

      historyRows.push([

        code,

        issueDate

      ]);



      // ======================================================
      // 回傳建立結果
      // ======================================================

      result.push({

        code:
          code,

        discount:
          discount,

        validUntil:
          formatDate(
            validUntil
          ),

        customerUrl:
          customerUrl

      });

    }



    // ========================================================
    // 批次寫入主優惠券工作表
    // ========================================================

    if (
      couponRows.length > 0
    ) {


      const startRow =

        sheet.getLastRow() + 1;



      sheet

        .getRange(

          startRow,

          1,

          couponRows.length,

          9

        )

        .setValues(
          couponRows
        );



      // 發行日期格式
      sheet

        .getRange(

          startRow,

          4,

          couponRows.length,

          1

        )

        .setNumberFormat(
          "yyyy/MM/dd"
        );



      // 有效期限格式
      sheet

        .getRange(

          startRow,

          5,

          couponRows.length,

          1

        )

        .setNumberFormat(
          "yyyy/MM/dd"
        );

    }



    // ========================================================
    // 批次寫入永久歷史紀錄
    // ========================================================

    if (
      historyRows.length > 0
    ) {


      const historyStartRow =

        historySheet.getLastRow() + 1;



      historySheet

        .getRange(

          historyStartRow,

          1,

          historyRows.length,

          2

        )

        .setValues(
          historyRows
        );



      historySheet

        .getRange(

          historyStartRow,

          2,

          historyRows.length,

          1

        )

        .setNumberFormat(
          "yyyy/MM/dd HH:mm:ss"
        );

    }



    // ========================================================
    // 回傳結果
    // ========================================================

    return {

      success: true,

      coupons:
        result

    };


  } finally {


    // ========================================================
    // 🔓 解除鎖定
    // ========================================================

    lock.releaseLock();

  }

}






// ============================================================
// 🎲 產生唯一優惠券序號
//
// 傳入所有已使用券號
// 確保不會重複
// ============================================================

function generateUniqueCouponCode(
  usedCodes
) {


  let code =

    generateRandomCouponCode();



  // ==========================================================
  // 如果重複
  // 一直重新產生
  // ==========================================================

  while (

    usedCodes.has(
      code
    )

  ) {


    code =

      generateRandomCouponCode();

  }



  return code;

}






// ============================================================
// 🎲 產生隨機優惠券序號
// ============================================================

function generateRandomCouponCode() {


  const chars =

    "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";



  let part1 = "";



  let part2 = "";



  for (

    let i = 0;

    i < 4;

    i++

  ) {


    part1 +=

      chars.charAt(

        Math.floor(

          Math.random() *

          chars.length

        )

      );



    part2 +=

      chars.charAt(

        Math.floor(

          Math.random() *

          chars.length

        )

      );

  }



  return (

    "XL-" +

    part1 +

    "-" +

    part2

  );

}






// ============================================================
// 🎲 舊版產生優惠券函式
//
// 保留給其他可能使用的功能
// ============================================================

function generateCouponCode() {


  const historySheet =

    getHistorySheet();



  const historyLastRow =

    historySheet.getLastRow();



  let historyCodes = [];



  if (
    historyLastRow >= 2
  ) {


    historyCodes =

      historySheet

        .getRange(

          2,

          1,

          historyLastRow - 1,

          1

        )

        .getValues()

        .flat()

        .map(

          value =>

            String(value)
              .trim()
              .toUpperCase()

        );

  }



  const usedCodes =

    new Set(
      historyCodes
    );



  return generateUniqueCouponCode(
    usedCodes
  );

}






// ============================================================
// 📅 日期格式
// ============================================================

function formatDate(date) {


  if (!date) {

    return "";

  }



  return Utilities.formatDate(

    new Date(
      date
    ),

    Session.getScriptTimeZone(),

    "yyyy/MM/dd"

  );

}






// ============================================================
// 🕒 日期時間格式
// ============================================================

function formatDateTime(date) {


  if (!date) {

    return "";

  }



  return Utilities.formatDate(

    new Date(
      date
    ),

    Session.getScriptTimeZone(),

    "yyyy/MM/dd HH:mm:ss"

  );

}






// ============================================================
// 📅 取得日期物件
// ============================================================

function parseDate(value) {


  if (!value) {

    return null;

  }



  // ==========================================================
  // Google Sheet 本身就是 Date
  // ==========================================================

  if (

    value instanceof Date

  ) {

    return value;

  }



  // ==========================================================
  // yyyy/MM/dd
  // ==========================================================

  const text =

    String(
      value
    )
    .trim();



  const parts =

    text.split(
      "/"
    );



  if (

    parts.length === 3

  ) {


    return new Date(

      Number(
        parts[0]
      ),

      Number(
        parts[1]
      ) - 1,

      Number(
        parts[2]
      )

    );

  }



  return new Date(
    value
  );

}






// ============================================================
// 🎫 取得優惠券資料
// ============================================================

function getCoupon(code) {


  if (!code) {

    return null;

  }



  code =

    String(
      code
    )
    .trim();



  const sheet =

    getSheet();



  const lastRow =

    sheet.getLastRow();



  if (

    lastRow < 2

  ) {

    return null;

  }



  const data =

    sheet

      .getRange(

        2,

        1,

        lastRow - 1,

        9

      )

      .getValues();



  // ==========================================================
  // 尋找優惠券
  // ==========================================================

  for (

    let i = 0;

    i < data.length;

    i++

  ) {


    const row =

      data[i];



    if (

      String(
        row[0]
      )

      .trim()

      .toUpperCase()

      ===

      code.toUpperCase()

    ) {


      let status =

        String(
          row[2]
        )
        .trim();



      const today =

        new Date();



      const todayStart =

        new Date(

          today.getFullYear(),

          today.getMonth(),

          today.getDate()

        );



      const validUntil =

        parseDate(
          row[4]
        );



      // ======================================================
      // 檢查是否過期
      // ======================================================

      if (

        status === "未使用" &&

        validUntil &&

        validUntil < todayStart

      ) {


        status =

          "已過期";



        sheet

          .getRange(

            i + 2,

            3

          )

          .setValue(

            "已過期"

          );

      }



      // ======================================================
      // QR Code 圖片
      // ======================================================

      const qrContent =

        row[8] ||

        row[7];



      const qrUrl =

        "https://quickchart.io/qr?size=300&text=" +

        encodeURIComponent(
          qrContent
        );



      // ======================================================
      // 回傳資料
      // ======================================================

      return {

        code:

          String(
            row[0]
          ),

        discount:

          row[1],

        status:

          status,

        issueDate:

          formatDate(
            row[3]
          ),

        validUntil:

          formatDate(
            row[4]
          ),

        usedDate:

          row[5]

            ? formatDateTime(
                row[5]
              )

            : "",

        note:

          row[6] || "",

        customerUrl:

          row[7] || "",

        qrUrl:

          qrUrl

      };

    }

  }



  return null;

}






// ============================================================
// 📷 QR Code / 手動輸入核銷
// ============================================================

function redeemCoupon(input) {


  if (!input) {

    return {

      success: false,

      message:
        "請輸入或掃描優惠券"

    };

  }



  const couponCode =

    extractCouponCode(
      input
    );



  if (!couponCode) {

    return {

      success: false,

      message:
        "無法辨識優惠券"

    };

  }



  const lock =

    LockService
      .getScriptLock();



  lock.waitLock(
    30000
  );



  try {


    const sheet =

      getSheet();



    const lastRow =

      sheet.getLastRow();



    if (

      lastRow < 2

    ) {

      return {

        success: false,

        message:
          "目前沒有優惠券資料"

      };

    }



    const data =

      sheet

        .getRange(

          2,

          1,

          lastRow - 1,

          9

        )

        .getValues();



    // ========================================================
    // 尋找優惠券
    // ========================================================

    for (

      let i = 0;

      i < data.length;

      i++

    ) {


      const row =

        data[i];



      const currentCode =

        String(
          row[0]
        )

        .trim()

        .toUpperCase();



      if (

        currentCode ===

        couponCode

      ) {


        const status =

          String(
            row[2]
          )
          .trim();



        // ====================================================
        // 已使用
        // ====================================================

        if (

          status === "已使用"

        ) {

          return {

            success: false,

            message:
              "❌ 此優惠券已經使用過",

            code:
              couponCode,

            usedDate:

              row[5]

                ? formatDateTime(
                    row[5]
                  )

                : ""

          };

        }



        // ====================================================
        // 已過期
        // ====================================================

        if (

          status === "已過期"

        ) {

          return {

            success: false,

            message:
              "❌ 此優惠券已過期",

            code:
              couponCode

          };

        }



        // ====================================================
        // 檢查有效期限
        // ====================================================

        const today =

          new Date();



        const todayStart =

          new Date(

            today.getFullYear(),

            today.getMonth(),

            today.getDate()

          );



        const validUntil =

          parseDate(
            row[4]
          );



        if (

          validUntil &&

          validUntil < todayStart

        ) {


          sheet

            .getRange(

              i + 2,

              3

            )

            .setValue(

              "已過期"

            );



          return {

            success: false,

            message:
              "❌ 此優惠券已過期",

            code:
              couponCode

          };

        }



        // ====================================================
        // 🎉 執行核銷
        // ====================================================

        const usedDate =

          new Date();



        sheet

          .getRange(

            i + 2,

            3

          )

          .setValue(

            "已使用"

          );



        sheet

          .getRange(

            i + 2,

            6

          )

          .setValue(

            usedDate

          );



        sheet

          .getRange(

            i + 2,

            6

          )

          .setNumberFormat(

            "yyyy/MM/dd HH:mm:ss"

          );



        return {

          success: true,

          message:
            "🎉 核銷成功",

          code:
            couponCode,

          discount:
            row[1],

          note:
            row[6] || "",

          usedDate:

            formatDateTime(
              usedDate
            )

        };

      }

    }



    // ========================================================
    // 找不到優惠券
    // ========================================================

    return {

      success: false,

      message:
        "❌ 找不到此優惠券"

    };


  } finally {


    lock.releaseLock();

  }

}






// ============================================================
// 🔍 從 QR Code / 網址 / 券號
// 取得優惠券序號
// ============================================================

function extractCouponCode(input) {


  input =

    String(
      input
    )
    .trim();



  // ==========================================================
  // 如果掃描到的是網址
  // ==========================================================

  if (

    input.indexOf(
      "coupon="
    )

    !==

    -1

  ) {


    const match =

      input.match(

        /coupon=([^&]+)/i

      );



    if (

      match &&

      match[1]

    ) {


      return decodeURIComponent(

        match[1]

      )

      .trim()

      .toUpperCase();

    }

  }



  // ==========================================================
  // 如果直接掃到券號
  // ==========================================================

  const match =

    input.match(

      /XL-[A-Z0-9]{4}-[A-Z0-9]{4}/i

    );



  if (match) {


    return match[0]

      .toUpperCase();

  }



  return null;

}
