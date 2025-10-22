// ===== Config.js =====

/**
 * 設定・定数管理
 * @version 1.1.0
 */

/**
 * アプリケーション設定
 */
const CONFIG = {
  // アプリケーション情報
  APP: {
    NAME: "文化祭スマート販売システム",
    VERSION: "1.1.0",
    TIMEZONE: "Asia/Tokyo",
  },

  // シート名
  SHEETS: {
    PRODUCTS: "商品マスタ",
    INVENTORY: "在庫管理",
    SALES: "販売履歴",
    SETTINGS: "設定",
    DASHBOARD: "ダッシュボード",
    ERROR_LOG: "エラーログ",
  },

  // 商品マスタのカラム定義
  PRODUCT_COLUMNS: {
    ID: "商品ID",
    NAME: "商品名",
    PRICE: "価格",
    CATEGORY: "カテゴリ",
    INITIAL_STOCK: "初期在庫",
    IMAGE_URL: "画像URL",
    DESCRIPTION: "説明",
    IS_ACTIVE: "有効フラグ",
  },

  // 在庫管理のカラム定義
  INVENTORY_COLUMNS: {
    ID: "商品ID",
    NAME: "商品名",
    CURRENT_STOCK: "現在庫",
  },

  // 販売履歴のカラム定義
  SALES_COLUMNS: {
    TRANSACTION_ID: "取引ID",
    TIMESTAMP: "販売日時",
    PRODUCT_ID: "商品ID",
    PRODUCT_NAME: "商品名",
    PRICE: "価格",
    QUANTITY: "数量",
    SUBTOTAL: "小計",
    PAYMENT_METHOD: "決済方法",
    STAFF_ID: "スタッフID",
  },

  // 設定のキー
  SETTINGS_KEYS: {
    SHOP_NAME: "shop_name",
    QR_PAYMENT_NAME: "qr_payment_name",
    QR_CODE_IMAGE_URL: "qr_code_image_url",
    SALES_TARGET: "sales_target",
    OPENING_TIME: "opening_time",
    CLOSING_TIME: "closing_time",
    STAFF_COUNT: "staff_count",
    CURRENCY_SYMBOL: "currency_symbol",
    LOCATION: "location",
    WEATHER_API_KEY: "weather_api_key",
  },

  // API設定
  API: {
    GEMINI: {
      BASE_URL: "https://generativelanguage.googleapis.com/v1beta/models",
      MODEL: "gemini-2.5-flash-lite",
      ENDPOINT: "generateContent",
      TEMPERATURE: 0.1,
      MAX_OUTPUT_TOKENS: 1024,
    },
  },

  // デフォルト値
  DEFAULTS: {
    SHOP_NAME: "3年A組 文化祭模擬店",
    PAYMENT_METHODS: ["現金", "QR"],
    CURRENCY_SYMBOL: "¥",
    STAFF_ID: "STAFF001",
  },

  // 制限値
  LIMITS: {
    MAX_CONCURRENT_USERS: 5,
    ERROR_LOG_RETENTION_DAYS: 30,
    DASHBOARD_UPDATE_INTERVAL: 5, // 分
    MAX_CART_ITEMS: 50,
  },

  // 在庫アラート設定
  INVENTORY_ALERT: {
    LOW_STOCK_THRESHOLD: 10, // 在庫僅少の閾値
    CRITICAL_STOCK_THRESHOLD: 5, // 危険在庫の閾値
    REORDER_POINT_MULTIPLIER: 1.5, // 発注点の係数
    SALES_HISTORY_HOURS: 2, // 売上予測に使用する過去時間
    PREDICTION_INTERVAL_MINUTES: 30, // 予測更新間隔
    ALERT_LEVELS: {
      INFO: "info",
      WARNING: "warning",
      CRITICAL: "critical",
      SOLD_OUT: "sold_out",
    },
  },
};

/**
 * 設定をオブジェクトにフリーズして変更を防ぐ
 */
Object.freeze(CONFIG);
Object.freeze(CONFIG.APP);
Object.freeze(CONFIG.SHEETS);
Object.freeze(CONFIG.PRODUCT_COLUMNS);
Object.freeze(CONFIG.INVENTORY_COLUMNS);
Object.freeze(CONFIG.SALES_COLUMNS);
Object.freeze(CONFIG.SETTINGS_KEYS);
Object.freeze(CONFIG.API);
Object.freeze(CONFIG.API.GEMINI);
Object.freeze(CONFIG.DEFAULTS);
Object.freeze(CONFIG.LIMITS);
Object.freeze(CONFIG.INVENTORY_ALERT);
Object.freeze(CONFIG.INVENTORY_ALERT.ALERT_LEVELS);

// ===== Namespace bootstrap (non-breaking refactor) =====
// 提供: シンプルな名前空間でクラス／サービス／エントリポイントをまとめ、
// 既存のグローバルはそのまま維持します（GASトリガー互換）。
(function bootstrapNamespace(global) {
  try {
    var ns = global.FesPos || {};
    ns.META = {
      NAME:
        (typeof CONFIG !== "undefined" && CONFIG.APP && CONFIG.APP.NAME) ||
        "Fes POS",
      VERSION:
        (typeof CONFIG !== "undefined" && CONFIG.APP && CONFIG.APP.VERSION) ||
        "1.x",
    };

    // 型/クラス参照
    ns.Classes = ns.Classes || {};
    if (typeof Database !== "undefined") ns.Classes.Database = Database;
    if (typeof ProductService !== "undefined")
      ns.Classes.ProductService = ProductService;
    if (typeof InventoryService !== "undefined")
      ns.Classes.InventoryService = InventoryService;
    if (typeof SalesService !== "undefined")
      ns.Classes.SalesService = SalesService;
    if (typeof InventoryAlertService !== "undefined")
      ns.Classes.InventoryAlertService = InventoryAlertService;
    if (typeof GeminiService !== "undefined")
      ns.Classes.GeminiService = GeminiService;

    // 設定参照
    if (typeof CONFIG !== "undefined") ns.CONFIG = CONFIG;

    // 既存のシングルトン／サービスインスタンス参照
    ns.Services = ns.Services || {};
    if (typeof db !== "undefined") ns.Services.db = db;
    if (typeof productService !== "undefined")
      ns.Services.productService = productService;
    if (typeof inventoryService !== "undefined")
      ns.Services.inventoryService = inventoryService;
    if (typeof salesService !== "undefined")
      ns.Services.salesService = salesService;
    if (typeof geminiService !== "undefined")
      ns.Services.geminiService = geminiService;
    if (typeof inventoryAlertService !== "undefined")
      ns.Services.inventoryAlertService = inventoryAlertService;

    // エントリポイント（GAS メニューやトリガーの関数）
    ns.Entry = ns.Entry || {};
    if (typeof onOpen === "function") ns.Entry.onOpen = onOpen;
    if (typeof initializeSpreadsheet === "function")
      ns.Entry.initializeSpreadsheet = initializeSpreadsheet;
    if (typeof generateDummyData === "function")
      ns.Entry.generateDummyData = generateDummyData;
    if (typeof runInventoryAlertCheck === "function")
      ns.Entry.runInventoryAlertCheck = runInventoryAlertCheck;
    if (typeof showReorderListInSheet === "function")
      ns.Entry.showReorderListInSheet = showReorderListInSheet;
    if (typeof showApiKeyDialog === "function")
      ns.Entry.showApiKeyDialog = showApiKeyDialog;
    if (typeof setWebAppUrl === "function")
      ns.Entry.setWebAppUrl = setWebAppUrl;
    if (typeof openWebApp === "function") ns.Entry.openWebApp = openWebApp;
    if (typeof openManual === "function") ns.Entry.openManual = openManual;
    if (typeof openCurriculum === "function")
      ns.Entry.openCurriculum = openCurriculum;

    global.FesPos = ns;
  } catch (e) {
    // Namespacing 自体は副作用なしが前提。失敗しても本体動作を阻害しない。
    if (typeof console !== "undefined" && console.error) {
      console.error("FesPos namespace bootstrap failed:", e);
    }
  }
})(this);

// ===== Utils.js =====

/**
 * 共通ユーティリティ関数
 * @version 1.1.0
 */

/**
 * 在庫データから商品IDをキーとしたマップを作成
 * @param {Array} inventoryData - 在庫データの配列
 * @returns {Object} 商品IDをキー、在庫数を値とするマップ
 */
function createStockMap(inventoryData) {
  const stockMap = {};
  inventoryData.forEach((row) => {
    const productId = row[CONFIG.INVENTORY_COLUMNS.ID];
    const currentStock = row[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK];
    if (productId) {
      stockMap[productId] = parseInt(currentStock) || 0;
    }
  });
  return stockMap;
}

/**
 * シートデータを作成または初期化するヘルパー関数
 * @param {string} sheetName - シート名
 * @param {Array} headers - ヘッダー配列
 * @param {Array} [initialData] - 初期データ（オプション）
 * @returns {Sheet} 作成されたシート
 */
function createSheetWithData(sheetName, headers, initialData = null) {
  const sheet = db.createSheet(sheetName, headers);
  if (initialData && initialData.length > 0) {
    sheet
      .getRange(2, 1, initialData.length, initialData[0].length)
      .setValues(initialData);
  }
  return sheet;
}

/**
 * エラーハンドリングラッパー
 * @param {Function} fn - ラップする関数
 * @param {string} context - エラーコンテキスト
 * @returns {Function} ラップされた関数
 */
function withErrorHandler(fn, context) {
  return function (...args) {
    try {
      return fn.apply(this, args);
    } catch (error) {
      console.error(`Error in ${context}:`, error);

      // エラーログシートが存在する場合は記録
      if (db && db.hasSheet(CONFIG.SHEETS.ERROR_LOG)) {
        db.addData(CONFIG.SHEETS.ERROR_LOG, [
          {
            発生日時: new Date(),
            コンテキスト: context,
            エラーメッセージ: error.toString(),
            スタックトレース: error.stack || "",
          },
        ]);
      }

      // ユーザー向けエラーレスポンス
      return {
        success: false,
        message: "システムエラーが発生しました。管理者に連絡してください。",
        error: error.toString(),
      };
    }
  };
}

/**
 * エラーログの古いエントリを削除
 */
function cleanupErrorLogs() {
  if (!db || !db.hasSheet(CONFIG.SHEETS.ERROR_LOG)) {
    return;
  }

  const errorLogs = db.getData(CONFIG.SHEETS.ERROR_LOG);
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(
    thirtyDaysAgo.getDate() - CONFIG.LIMITS.ERROR_LOG_RETENTION_DAYS,
  );

  const recentLogs = errorLogs.filter(
    (log) => new Date(log["発生日時"]) > thirtyDaysAgo,
  );

  if (recentLogs.length < errorLogs.length) {
    db.clearData(CONFIG.SHEETS.ERROR_LOG);
    db.addData(CONFIG.SHEETS.ERROR_LOG, recentLogs);
  }
}

// ===== Database.js =====

/**
 * データベース操作クラス
 * @version 1.1.0
 */

/**
 * スプレッドシートをデータベースとして操作するクラス
 */
class Database {
  constructor() {
    this.spreadsheet = null;
    this.sheets = {};
  }

  /**
   * スプレッドシートの取得（遅延初期化）
   */
  getSpreadsheet() {
    if (!this.spreadsheet) {
      this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this.spreadsheet;
  }

  /**
   * シートの取得（キャッシュ付き）
   */
  getSheet(sheetName) {
    if (!this.sheets[sheetName]) {
      this.sheets[sheetName] = this.getSpreadsheet().getSheetByName(sheetName);
    }
    return this.sheets[sheetName];
  }

  /**
   * シートの存在確認
   */
  hasSheet(sheetName) {
    return this.getSheet(sheetName) !== null;
  }

  /**
   * データを取得（ヘッダー付き）
   */
  getData(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = data[i][j];
      }
      records.push(record);
    }

    return records;
  }

  /**
   * データを追加
   */
  addData(sheetName, records) {
    const sheet = this.getSheet(sheetName);
    if (!sheet || !records || records.length === 0) return false;

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const rows = records.map((record) =>
      headers.map((header) => record[header] || ""),
    );

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    return true;
  }

  /**
   * データを更新
   */
  updateData(sheetName, keyColumn, keyValue, updates) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const keyIndex = headers.indexOf(keyColumn);

    if (keyIndex === -1) return false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][keyIndex] === keyValue) {
        Object.entries(updates).forEach(([column, value]) => {
          const columnIndex = headers.indexOf(column);
          if (columnIndex !== -1) {
            sheet.getRange(i + 1, columnIndex + 1).setValue(value);
          }
        });
        return true;
      }
    }

    return false;
  }

  /**
   * バッチ更新（高速化）
   */
  batchUpdate(sheetName, updates) {
    const sheet = this.getSheet(sheetName);
    if (!sheet || !updates || updates.length === 0) return false;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const keyIndex = headers.indexOf(updates[0].keyColumn);

    if (keyIndex === -1) return false;

    // 更新対象を収集
    const updateMap = new Map();
    updates.forEach((update) => {
      updateMap.set(update.keyValue, update.values);
    });

    // 一括更新
    const newData = data.map((row, index) => {
      if (index === 0) return row; // ヘッダーはスキップ

      const keyValue = row[keyIndex];
      const updateValues = updateMap.get(keyValue);

      if (updateValues) {
        const newRow = [...row];
        Object.entries(updateValues).forEach(([column, value]) => {
          const columnIndex = headers.indexOf(column);
          if (columnIndex !== -1) {
            newRow[columnIndex] = value;
          }
        });
        return newRow;
      }

      return row;
    });

    sheet.getDataRange().setValues(newData);
    return true;
  }

  /**
   * データをクリア（ヘッダーは残す）
   */
  clearData(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return false;

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return true;
  }

  /**
   * シートを作成
   */
  createSheet(sheetName, headers) {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (headers && headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    // キャッシュに追加
    this.sheets[sheetName] = sheet;
    return sheet;
  }
}

// ===== ProductService.js =====

/**
 * 商品管理サービス
 * @version 1.1.0
 */
class ProductService {
  constructor(database) {
    this.db = database;
  }

  getActiveProducts() {
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    return products.filter(
      (product) => product[CONFIG.PRODUCT_COLUMNS.IS_ACTIVE] === true,
    );
  }

  getProductsWithStock() {
    const products = this.getActiveProducts();
    const inventory = this.db.getData(CONFIG.SHEETS.INVENTORY);

    const stockMap = createStockMap(inventory);

    return products.map((product) => ({
      ...product,
      [CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]:
        stockMap[product[CONFIG.PRODUCT_COLUMNS.ID]] || 0,
    }));
  }

  getProductById(productId) {
    const products = this.getActiveProducts();
    return products.find(
      (product) => product[CONFIG.PRODUCT_COLUMNS.ID] === productId,
    );
  }

  searchProductsByName(searchTerm) {
    const products = this.getProductsWithStock();
    const lowerSearchTerm = searchTerm.toLowerCase();

    return products.filter((product) => {
      const name = product[CONFIG.PRODUCT_COLUMNS.NAME].toLowerCase();
      const description = (
        product[CONFIG.PRODUCT_COLUMNS.DESCRIPTION] || ""
      ).toLowerCase();

      return (
        name.includes(lowerSearchTerm) || description.includes(lowerSearchTerm)
      );
    });
  }

  getProductsByCategory(category) {
    const products = this.getProductsWithStock();

    if (!category) return products;

    return products.filter(
      (product) => product[CONFIG.PRODUCT_COLUMNS.CATEGORY] === category,
    );
  }

  getCategories() {
    const products = this.getActiveProducts();
    const categories = new Set();

    products.forEach((product) => {
      if (product[CONFIG.PRODUCT_COLUMNS.CATEGORY]) {
        categories.add(product[CONFIG.PRODUCT_COLUMNS.CATEGORY]);
      }
    });

    return Array.from(categories).sort();
  }

  addProduct(productData) {
    const productId = this.generateProductId();

    const newProduct = {
      [CONFIG.PRODUCT_COLUMNS.ID]: productId,
      [CONFIG.PRODUCT_COLUMNS.NAME]: productData.name,
      [CONFIG.PRODUCT_COLUMNS.PRICE]: productData.price,
      [CONFIG.PRODUCT_COLUMNS.CATEGORY]: productData.category || "",
      [CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK]: productData.initialStock || 0,
      [CONFIG.PRODUCT_COLUMNS.IMAGE_URL]: productData.imageUrl || "",
      [CONFIG.PRODUCT_COLUMNS.DESCRIPTION]: productData.description || "",
      [CONFIG.PRODUCT_COLUMNS.IS_ACTIVE]: true,
    };

    // 商品マスタに追加
    this.db.addData(CONFIG.SHEETS.PRODUCTS, [newProduct]);

    // 在庫管理に追加
    const inventoryData = {
      [CONFIG.INVENTORY_COLUMNS.ID]: productId,
      [CONFIG.INVENTORY_COLUMNS.NAME]: "", // VLOOKUPで自動入力
      [CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]: productData.initialStock || 0,
    };

    this.db.addData(CONFIG.SHEETS.INVENTORY, [inventoryData]);

    return productId;
  }

  generateProductId() {
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const maxId = products.reduce((max, product) => {
      const id = product[CONFIG.PRODUCT_COLUMNS.ID];
      const num = parseInt(id.replace("P", ""), 10);
      return Math.max(max, num || 0);
    }, 0);

    return `P${String(maxId + 1).padStart(3, "0")}`;
  }

  updateProduct(productId, updates) {
    return this.db.updateData(
      CONFIG.SHEETS.PRODUCTS,
      CONFIG.PRODUCT_COLUMNS.ID,
      productId,
      updates,
    );
  }

  deactivateProduct(productId) {
    return this.updateProduct(productId, {
      [CONFIG.PRODUCT_COLUMNS.IS_ACTIVE]: false,
    });
  }
}

// ===== InventoryService.js =====

/**
 * 在庫管理サービス
 * @version 1.1.0
 */
class InventoryService {
  constructor(database) {
    this.db = database;
  }

  /**
   * 在庫シートに商品IDの行が存在することを保証（不足分を追加）
   */
  syncInventoryWithProducts() {
    // シート存在を保証
    if (!this.db.hasSheet(CONFIG.SHEETS.INVENTORY)) {
      this.db.createSheet(
        CONFIG.SHEETS.INVENTORY,
        Object.values(CONFIG.INVENTORY_COLUMNS),
      );
    }
    const sheet = this.db.getSheet(CONFIG.SHEETS.INVENTORY);
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const inventory = this.db.getData(CONFIG.SHEETS.INVENTORY);
    const existingIds = new Set(
      inventory.map((r) => r[CONFIG.INVENTORY_COLUMNS.ID]),
    );

    const rowsToAppend = [];
    const formulas = [];
    products.forEach((p) => {
      const pid = p[CONFIG.PRODUCT_COLUMNS.ID];
      if (!existingIds.has(pid)) {
        rowsToAppend.push([pid, "", p[CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK]]);
      }
    });

    if (rowsToAppend.length > 0) {
      const startRow = sheet.getLastRow() + 1; // 追加開始行
      sheet
        .getRange(startRow, 1, rowsToAppend.length, 3)
        .setValues(rowsToAppend);
      // 商品名列にVLOOKUPを設定
      for (let i = 0; i < rowsToAppend.length; i++) {
        const r = startRow + i;
        sheet
          .getRange(r, 2)
          .setFormula(
            `=IFERROR(VLOOKUP(A${r},'${CONFIG.SHEETS.PRODUCTS}'!A:B,2,FALSE),"")`,
          );
      }
    }
    return true;
  }

  /**
   * すべての在庫データを取得
   */
  getInventory() {
    return this.db.getData(CONFIG.SHEETS.INVENTORY);
  }

  getStock(productId) {
    const inventory = this.db.getData(CONFIG.SHEETS.INVENTORY);
    const item = inventory.find(
      (row) => row[CONFIG.INVENTORY_COLUMNS.ID] === productId,
    );

    return item ? item[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK] : 0;
  }

  updateStock(productId, newStock) {
    return this.db.updateData(
      CONFIG.SHEETS.INVENTORY,
      CONFIG.INVENTORY_COLUMNS.ID,
      productId,
      { [CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]: Math.max(0, newStock) },
    );
  }

  decreaseStock(productId, quantity) {
    const currentStock = this.getStock(productId);
    const newStock = currentStock - quantity;

    if (newStock < 0) {
      throw new Error(`在庫不足: 商品ID ${productId}（残り${currentStock}個）`);
    }

    return this.updateStock(productId, newStock);
  }

  increaseStock(productId, quantity) {
    const currentStock = this.getStock(productId);
    return this.updateStock(productId, currentStock + quantity);
  }

  batchUpdateStock(updates) {
    const batchUpdates = updates.map((update) => ({
      keyColumn: CONFIG.INVENTORY_COLUMNS.ID,
      keyValue: update.productId,
      values: {
        [CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]: update.newStock,
      },
    }));

    return this.db.batchUpdate(CONFIG.SHEETS.INVENTORY, batchUpdates);
  }

  processSaleInventory(soldItems) {
    const inventory = this.db.getData(CONFIG.SHEETS.INVENTORY);
    const stockMap = createStockMap(inventory);

    // 在庫チェックと更新データの準備
    const updates = [];
    const errors = [];

    soldItems.forEach((item) => {
      const currentStock = stockMap[item[CONFIG.PRODUCT_COLUMNS.ID]] || 0;
      const newStock = currentStock - item["数量"];

      if (newStock < 0) {
        errors.push({
          productId: item[CONFIG.PRODUCT_COLUMNS.ID],
          productName: item[CONFIG.PRODUCT_COLUMNS.NAME],
          requested: item["数量"],
          available: currentStock,
        });
      } else {
        updates.push({
          productId: item[CONFIG.PRODUCT_COLUMNS.ID],
          newStock: newStock,
        });
      }
    });

    // エラーがある場合は例外を投げる
    if (errors.length > 0) {
      const errorMessages = errors.map(
        (err) => `${err.productName}: 在庫不足（残り${err.available}個）`,
      );
      throw new Error("在庫エラー:\n" + errorMessages.join("\n"));
    }

    // 一括更新
    return this.batchUpdateStock(updates);
  }

  getInventoryReport() {
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const inventory = this.db.getData(CONFIG.SHEETS.INVENTORY);
    const stockMap = createStockMap(inventory);

    const productsWithStock = products
      .filter((product) => product[CONFIG.PRODUCT_COLUMNS.IS_ACTIVE] === true)
      .map((product) => ({
        ...product,
        [CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]:
          stockMap[product[CONFIG.PRODUCT_COLUMNS.ID]] || 0,
      }));

    const report = {
      totalProducts: productsWithStock.length,
      totalStock: 0,
      lowStock: [],
      outOfStock: [],
      categories: {},
    };

    productsWithStock.forEach((product) => {
      const stock = product[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK];
      const category = product[CONFIG.PRODUCT_COLUMNS.CATEGORY];

      report.totalStock += stock;

      if (!report.categories[category]) {
        report.categories[category] = {
          count: 0,
          totalStock: 0,
        };
      }
      report.categories[category].count++;
      report.categories[category].totalStock += stock;

      if (stock === 0) {
        report.outOfStock.push({
          id: product[CONFIG.PRODUCT_COLUMNS.ID],
          name: product[CONFIG.PRODUCT_COLUMNS.NAME],
        });
      } else if (stock <= 10) {
        report.lowStock.push({
          id: product[CONFIG.PRODUCT_COLUMNS.ID],
          name: product[CONFIG.PRODUCT_COLUMNS.NAME],
          stock: stock,
        });
      }
    });

    return report;
  }

  resetInventory() {
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const updates = [];

    products.forEach((product) => {
      updates.push({
        productId: product[CONFIG.PRODUCT_COLUMNS.ID],
        newStock: product[CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK],
      });
    });

    return this.batchUpdateStock(updates);
  }
}

// ===== SalesService.js =====

/**
 * 販売管理サービス
 * @version 1.1.0
 */

/**
 * 販売に関する操作を管理するクラス
 */
class SalesService {
  constructor(database, inventoryService) {
    this.db = database;
    this.inventoryService = inventoryService;
  }

  /**
   * 取引を処理
   */
  processTransaction(
    cartItems,
    paymentMethod,
    staffId = CONFIG.DEFAULTS.STAFF_ID,
  ) {
    try {
      // トランザクションIDを生成
      const transactionId = this.generateTransactionId();

      // 在庫チェックと更新
      this.inventoryService.processSaleInventory(cartItems);

      // 販売履歴に記録
      this.recordSales(transactionId, cartItems, paymentMethod, staffId);

      // ダッシュボードを更新
      this.updateDashboard();

      return {
        success: true,
        transactionId: transactionId,
        total: this.calculateTotal(cartItems),
      };
    } catch (error) {
      console.error("Transaction error:", error);
      throw error;
    }
  }

  /**
   * 取引IDを生成
   */
  generateTransactionId() {
    const now = new Date();
    const timestamp = Utilities.formatDate(
      now,
      CONFIG.APP.TIMEZONE,
      "yyyyMMddHHmmss",
    );
    const random = Math.floor(Math.random() * 1000)
      .toString()
      .padStart(3, "0");
    return `T${timestamp}${random}`;
  }

  /**
   * 販売履歴に記録
   */
  recordSales(transactionId, cartItems, paymentMethod, staffId) {
    const timestamp = new Date();
    const salesRecords = [];

    cartItems.forEach((item) => {
      salesRecords.push({
        [CONFIG.SALES_COLUMNS.TRANSACTION_ID]: transactionId,
        [CONFIG.SALES_COLUMNS.TIMESTAMP]: timestamp,
        [CONFIG.SALES_COLUMNS.PRODUCT_ID]: item[CONFIG.PRODUCT_COLUMNS.ID],
        [CONFIG.SALES_COLUMNS.PRODUCT_NAME]: item[CONFIG.PRODUCT_COLUMNS.NAME],
        [CONFIG.SALES_COLUMNS.PRICE]: item[CONFIG.PRODUCT_COLUMNS.PRICE],
        [CONFIG.SALES_COLUMNS.QUANTITY]: item["数量"],
        [CONFIG.SALES_COLUMNS.SUBTOTAL]:
          item[CONFIG.PRODUCT_COLUMNS.PRICE] * item["数量"],
        [CONFIG.SALES_COLUMNS.PAYMENT_METHOD]: paymentMethod,
        [CONFIG.SALES_COLUMNS.STAFF_ID]: staffId,
      });
    });

    return this.db.addData(CONFIG.SHEETS.SALES, salesRecords);
  }

  /**
   * 合計金額を計算
   */
  calculateTotal(cartItems) {
    return cartItems.reduce(
      (total, item) =>
        total + item[CONFIG.PRODUCT_COLUMNS.PRICE] * item["数量"],
      0,
    );
  }

  /**
   * 販売履歴を取得
   */
  getSalesHistory(options = {}) {
    const sales = this.db.getData(CONFIG.SHEETS.SALES);

    // フィルタリング
    let filtered = sales;

    if (options.startDate) {
      filtered = filtered.filter(
        (sale) =>
          new Date(sale[CONFIG.SALES_COLUMNS.TIMESTAMP]) >= options.startDate,
      );
    }

    if (options.endDate) {
      filtered = filtered.filter(
        (sale) =>
          new Date(sale[CONFIG.SALES_COLUMNS.TIMESTAMP]) <= options.endDate,
      );
    }

    if (options.productId) {
      filtered = filtered.filter(
        (sale) => sale[CONFIG.SALES_COLUMNS.PRODUCT_ID] === options.productId,
      );
    }

    if (options.paymentMethod) {
      filtered = filtered.filter(
        (sale) =>
          sale[CONFIG.SALES_COLUMNS.PAYMENT_METHOD] === options.paymentMethod,
      );
    }

    // ソート（新しい順）
    filtered.sort(
      (a, b) =>
        new Date(b[CONFIG.SALES_COLUMNS.TIMESTAMP]) -
        new Date(a[CONFIG.SALES_COLUMNS.TIMESTAMP]),
    );

    return filtered;
  }

  /**
   * 売上統計を取得
   */
  getSalesStatistics(startDate, endDate) {
    const sales = this.getSalesHistory({ startDate, endDate });

    const stats = {
      totalSales: 0,
      totalTransactions: new Set(),
      productSales: {},
      paymentMethods: {},
      hourlySales: {},
      topProducts: [],
    };

    sales.forEach((sale) => {
      const subtotal = sale[CONFIG.SALES_COLUMNS.SUBTOTAL];
      const productName = sale[CONFIG.SALES_COLUMNS.PRODUCT_NAME];
      const paymentMethod = sale[CONFIG.SALES_COLUMNS.PAYMENT_METHOD];
      const transactionId = sale[CONFIG.SALES_COLUMNS.TRANSACTION_ID];
      const hour = new Date(sale[CONFIG.SALES_COLUMNS.TIMESTAMP]).getHours();

      // 総売上
      stats.totalSales += subtotal;

      // 取引数
      stats.totalTransactions.add(transactionId);

      // 商品別売上
      if (!stats.productSales[productName]) {
        stats.productSales[productName] = {
          quantity: 0,
          sales: 0,
        };
      }
      stats.productSales[productName].quantity +=
        sale[CONFIG.SALES_COLUMNS.QUANTITY];
      stats.productSales[productName].sales += subtotal;

      // 決済方法別
      if (!stats.paymentMethods[paymentMethod]) {
        stats.paymentMethods[paymentMethod] = {
          count: 0,
          sales: 0,
        };
      }
      stats.paymentMethods[paymentMethod].count++;
      stats.paymentMethods[paymentMethod].sales += subtotal;

      // 時間帯別
      if (!stats.hourlySales[hour]) {
        stats.hourlySales[hour] = 0;
      }
      stats.hourlySales[hour] += subtotal;
    });

    // 取引数を数値に変換
    stats.totalTransactions = stats.totalTransactions.size;

    // 売れ筋商品TOP5
    stats.topProducts = Object.entries(stats.productSales)
      .sort((a, b) => b[1].sales - a[1].sales)
      .slice(0, 5)
      .map(([name, data]) => ({
        name: name,
        quantity: data.quantity,
        sales: data.sales,
      }));

    return stats;
  }

  /**
   * ダッシュボードを更新
   */
  updateDashboard() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const stats = this.getSalesStatistics(today, new Date());
    const dashboardSheet = this.db.getSheet(CONFIG.SHEETS.DASHBOARD);

    if (!dashboardSheet) return;

    // シートをクリア
    dashboardSheet.clear();

    // ヘッダー
    const headers = [["項目", "値", "", "最終更新"]];
    dashboardSheet.getRange(1, 1, 1, 4).setValues(headers);
    dashboardSheet.getRange(1, 4).setValue(new Date());

    // 基本統計
    const basicStats = [
      ["売上総額", stats.totalSales],
      ["取引数", stats.totalTransactions],
      [
        "平均単価",
        stats.totalTransactions > 0
          ? Math.round(stats.totalSales / stats.totalTransactions)
          : 0,
      ],
    ];

    dashboardSheet.getRange(2, 1, basicStats.length, 2).setValues(basicStats);

    // 数値フォーマット
    dashboardSheet.getRange(2, 2, 3, 1).setNumberFormat("¥#,##0");

    // 商品別売上（見出し）
    dashboardSheet.getRange(6, 1).setValue("商品別売上TOP5");
    dashboardSheet
      .getRange(7, 1, 1, 3)
      .setValues([["商品名", "数量", "売上金額"]]);

    // 商品別売上（データ）
    if (stats.topProducts.length > 0) {
      const productData = stats.topProducts.map((product) => [
        product.name,
        product.quantity,
        product.sales,
      ]);
      dashboardSheet
        .getRange(8, 1, productData.length, 3)
        .setValues(productData);
      dashboardSheet
        .getRange(8, 3, productData.length, 1)
        .setNumberFormat("¥#,##0");
    }

    // 決済方法別
    const paymentStartRow = 8 + stats.topProducts.length + 2;
    dashboardSheet.getRange(paymentStartRow, 1).setValue("決済方法別");
    dashboardSheet
      .getRange(paymentStartRow + 1, 1, 1, 3)
      .setValues([["決済方法", "件数", "売上金額"]]);

    const paymentData = Object.entries(stats.paymentMethods).map(
      ([method, data]) => [method, data.count, data.sales],
    );

    if (paymentData.length > 0) {
      dashboardSheet
        .getRange(paymentStartRow + 2, 1, paymentData.length, 3)
        .setValues(paymentData);
      dashboardSheet
        .getRange(paymentStartRow + 2, 3, paymentData.length, 1)
        .setNumberFormat("¥#,##0");
    }
  }
}

// ===== DailyOpsService.js =====

/**
 * 日次作業サービス
 * @description 一日の売上集計と翌日の在庫準備（初期在庫への補充/リセット）
 */
class DailyOpsService {
  constructor(database, salesService, inventoryService) {
    this.db = database;
    this.salesService = salesService;
    this.inventoryService = inventoryService;
  }

  /**
   * 指定日の売上データを取得
   * @param {Date} date - 対象日（ローカル日付）
   * @returns {Array<Object>} 売上行
   */
  getSalesOfDate(date) {
    const tz = CONFIG.APP.TIMEZONE;
    const dayStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
    const sales = this.db.getData(CONFIG.SHEETS.SALES);
    return sales.filter((row) => {
      const ts = row[CONFIG.SALES_COLUMNS.TIMESTAMP];
      if (!ts) return false;
      const d = ts instanceof Date ? ts : new Date(ts);
      const s = Utilities.formatDate(d, tz, "yyyy-MM-dd");
      return s === dayStr;
    });
  }

  /**
   * 指定日の売上集計
   */
  summarizeSales(date) {
    const rows = this.getSalesOfDate(date);
    const summary = {
      date: Utilities.formatDate(date, CONFIG.APP.TIMEZONE, "yyyy-MM-dd"),
      totalAmount: 0,
      totalTransactions: 0,
      totalItems: 0,
      byPayment: {},
      byProduct: {},
    };

    rows.forEach((r) => {
      const amount = Number(r[CONFIG.SALES_COLUMNS.SUBTOTAL]) || 0;
      const qty = Number(r[CONFIG.SALES_COLUMNS.QUANTITY]) || 0;
      const pid = r[CONFIG.SALES_COLUMNS.PRODUCT_ID];
      const pname = r[CONFIG.SALES_COLUMNS.PRODUCT_NAME];
      const pay = r[CONFIG.SALES_COLUMNS.PAYMENT_METHOD] || "不明";

      summary.totalAmount += amount;
      summary.totalItems += qty;
      summary.totalTransactions += 1;
      summary.byPayment[pay] = (summary.byPayment[pay] || 0) + amount;
      if (!summary.byProduct[pid]) {
        summary.byProduct[pid] = { id: pid, name: pname, qty: 0, amount: 0 };
      }
      summary.byProduct[pid].qty += qty;
      summary.byProduct[pid].amount += amount;
    });

    return summary;
  }

  /**
   * 日次レポートシートを作成
   */
  generateDailyReport(date = new Date()) {
    const ss = this.db.getSpreadsheet();
    const sdate = Utilities.formatDate(date, CONFIG.APP.TIMEZONE, "yyyyMMdd");
    const baseName = `日次集計_${sdate}`;
    let sheetName = baseName;
    let i = 1;
    while (ss.getSheetByName(sheetName)) {
      sheetName = `${baseName}_${String(i++)}`;
    }

    const sheet = ss.insertSheet(sheetName);

    // ヘッダー
    sheet.getRange(1, 1, 1, 1).setValue("日次集計");
    sheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);

    const summary = this.summarizeSales(date);

    // 概要
    const overview = [
      ["日付", summary.date],
      ["総売上", summary.totalAmount],
      ["取引件数", summary.totalTransactions],
      ["販売個数", summary.totalItems],
    ];
    sheet.getRange(3, 1, overview.length, 2).setValues(overview);

    // 決済内訳
    const paymentRows = Object.keys(summary.byPayment).map((k) => [
      k,
      summary.byPayment[k],
    ]);
    sheet.getRange(3, 4).setValue("決済内訳").setFontWeight("bold");
    if (paymentRows.length > 0) {
      sheet
        .getRange(4, 4, 1, 2)
        .setValues([["決済方法", "金額"]])
        .setFontWeight("bold")
        .setBackground("#f0f0f0");
      sheet.getRange(5, 4, paymentRows.length, 2).setValues(paymentRows);
    }

    // 商品別売上（数量降順）
    const products = Object.values(summary.byProduct).sort(
      (a, b) => b.qty - a.qty,
    );
    let startRow = 6 + Math.max(paymentRows.length, 1);
    sheet.getRange(startRow, 1).setValue("商品別売上").setFontWeight("bold");
    sheet
      .getRange(startRow + 1, 1, 1, 4)
      .setValues([["商品ID", "商品名", "数量", "売上"]])
      .setFontWeight("bold")
      .setBackground("#f0f0f0");
    if (products.length > 0) {
      const prodRows = products.map((p) => [p.id, p.name, p.qty, p.amount]);
      sheet.getRange(startRow + 2, 1, prodRows.length, 4).setValues(prodRows);
    }

    // 翌日の準備在庫（初期在庫との差分）
    const prodData = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const invData = this.db.getData(CONFIG.SHEETS.INVENTORY);
    const stockMap = createStockMap(invData);
    const restock = prodData
      .map((p) => {
        const pid = p[CONFIG.PRODUCT_COLUMNS.ID];
        const name = p[CONFIG.PRODUCT_COLUMNS.NAME];
        const init = Number(p[CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK]) || 0;
        const cur = stockMap[pid] || 0;
        const need = Math.max(init - cur, 0);
        return [pid, name, cur, init, need];
      })
      .filter((row) => row[4] > 0);

    startRow = startRow + 4 + (products.length || 1);
    sheet
      .getRange(startRow, 1)
      .setValue("翌日の準備在庫（差分）")
      .setFontWeight("bold");
    sheet
      .getRange(startRow + 1, 1, 1, 5)
      .setValues([["商品ID", "商品名", "現在庫", "初期在庫", "補充必要数"]])
      .setFontWeight("bold")
      .setBackground("#f0f0f0");
    if (restock.length > 0) {
      sheet.getRange(startRow + 2, 1, restock.length, 5).setValues(restock);
    } else {
      sheet.getRange(startRow + 2, 1).setValue("補充が必要な商品はありません");
    }

    sheet.autoResizeColumns(1, 8);

    // 日次サマリーを更新
    this.upsertDailySummary(summary);
    return { sheetName };
  }

  /**
   * 在庫を初期在庫にリセット（補充済としてカウント開始）
   */
  restockToInitial() {
    return this.inventoryService.resetInventory();
  }

  /**
   * 日次サマリーシートの用意
   */
  ensureDailySummarySheet() {
    const name = "日次サマリー";
    const headers = [
      "日付",
      "総売上",
      "取引件数",
      "販売個数",
      "現金売上",
      "QR売上",
      "前日総売上",
      "前日比(額)",
      "前日比(%)",
      "累計売上",
    ];
    if (!this.db.hasSheet(name)) {
      return this.db.createSheet(name, headers);
    }
    const sheet = this.db.getSheet(name);
    // ヘッダーが未設定/ずれている場合に整える
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (firstRow[0] !== "日付") {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    return sheet;
  }

  /**
   * 日次サマリーに今日の結果を反映し、累計を更新
   */
  upsertDailySummary(summary) {
    const sheet = this.ensureDailySummarySheet();
    const tz = CONFIG.APP.TIMEZONE;
    const dateStr = summary.date; // yyyy-MM-dd

    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const rows = values.slice(1);

    const indexByDate = new Map();
    rows.forEach((r, i) => indexByDate.set(r[0], i + 2)); // 2-based row index

    const cash = summary.byPayment["現金"] || 0;
    const qr = summary.byPayment["QR"] || 0;

    // 前日データ
    const d = new Date(dateStr + "T00:00:00");
    const prev = new Date(d.getTime() - 24 * 60 * 60 * 1000);
    const prevKey = Utilities.formatDate(prev, tz, "yyyy-MM-dd");
    const prevRowIdx = indexByDate.get(prevKey);
    let prevTotal = 0;
    if (prevRowIdx) {
      prevTotal = Number(sheet.getRange(prevRowIdx, 2).getValue()) || 0;
    }

    const diffAmt = summary.totalAmount - prevTotal;
    const diffPct =
      prevTotal === 0 ? "" : Math.round((diffAmt / prevTotal) * 1000) / 10; // 0.1%刻み

    // 既存更新 or 追加
    const rowValues = [
      dateStr,
      summary.totalAmount,
      summary.totalTransactions,
      summary.totalItems,
      cash,
      qr,
      prevTotal,
      diffAmt,
      diffPct,
      "", // 累計は後で計算
    ];

    let writeRow = indexByDate.get(dateStr);
    if (writeRow) {
      sheet.getRange(writeRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      const last = sheet.getLastRow();
      sheet.getRange(last + 1, 1, 1, rowValues.length).setValues([rowValues]);
    }

    // 累計を再計算（日付昇順）
    const dataRange = sheet.getDataRange().getValues();
    const data = dataRange.slice(1).filter((r) => r[0]);
    data.sort((a, b) => (a[0] < b[0] ? -1 : 1));
    let running = 0;
    for (let i = 0; i < data.length; i++) {
      running += Number(data[i][1]) || 0;
      data[i][9 - 1] = running; // 9th index (0-based 8) but we have 10 columns; cum at index 9 (1-based 10)
    }
    // 上書き（2行目から）
    if (data.length > 0) {
      // 並び替え後のデータを反映（既存の順序は維持せず、単純更新に留める）
      // 既存の行数をクリアしてから書き直し
      const sheetTotalRows = sheet.getLastRow();
      if (sheetTotalRows > 1) sheet.deleteRows(2, sheetTotalRows - 1);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }

  /**
   * 販売履歴を日別シートへアーカイブ（指定日）。指定日の売上行を移動し、元シートから削除。
   */
  archiveSales(date = new Date()) {
    const tz = CONFIG.APP.TIMEZONE;
    const dayStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
    const dayKey = Utilities.formatDate(date, tz, "yyyyMMdd");
    const salesSheet = this.db.getSheet(CONFIG.SHEETS.SALES);
    if (!salesSheet)
      return { success: false, message: "販売履歴シートが見つかりません" };

    const values = salesSheet.getDataRange().getValues();
    if (values.length <= 1)
      return { success: false, message: "販売データがありません" };
    const headers = values[0];
    const tsIndex = headers.indexOf(CONFIG.SALES_COLUMNS.TIMESTAMP);
    if (tsIndex === -1)
      return { success: false, message: "タイムスタンプ列が見つかりません" };

    const data = values.slice(1);
    const todays = [];
    const remains = [];
    data.forEach((row) => {
      const ts = row[tsIndex];
      const d = ts instanceof Date ? ts : new Date(ts);
      const s = Utilities.formatDate(d, tz, "yyyy-MM-dd");
      if (s === dayStr) todays.push(row);
      else remains.push(row);
    });

    if (todays.length === 0)
      return { success: false, message: `${dayStr} の販売データはありません` };

    // アーカイブシート作成
    const ss = this.db.getSpreadsheet();
    let name = `販売履歴_${dayKey}`;
    let idx = 1;
    while (ss.getSheetByName(name)) name = `販売履歴_${dayKey}_${idx++}`;
    const archive = ss.insertSheet(name);
    archive.getRange(1, 1, 1, headers.length).setValues([headers]);
    archive.getRange(2, 1, todays.length, headers.length).setValues(todays);
    archive.autoResizeColumns(1, headers.length);

    // 元シートを更新（ヘッダー + remains）
    salesSheet.clearContents();
    salesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (remains.length > 0) {
      salesSheet
        .getRange(2, 1, remains.length, headers.length)
        .setValues(remains);
    }

    return { success: true, sheetName: name, archived: todays.length };
  }

  /**
   * 月次サマリーを作成/更新（年月: yyyy-MM）
   */
  generateMonthlySummary(monthDate = new Date()) {
    const tz = CONFIG.APP.TIMEZONE;
    const ym = Utilities.formatDate(monthDate, tz, "yyyy-MM");
    const ymKey = Utilities.formatDate(monthDate, tz, "yyyyMM");
    const sales = this.db.getData(CONFIG.SHEETS.SALES);

    // 対象月の行抽出
    const monthRows = sales.filter((r) => {
      const ts = r[CONFIG.SALES_COLUMNS.TIMESTAMP];
      if (!ts) return false;
      const d = ts instanceof Date ? ts : new Date(ts);
      return Utilities.formatDate(d, tz, "yyyy-MM") === ym;
    });

    const summary = {
      ym,
      totalAmount: 0,
      totalTransactions: 0,
      totalItems: 0,
      byPayment: {},
    };

    monthRows.forEach((r) => {
      const amt = Number(r[CONFIG.SALES_COLUMNS.SUBTOTAL]) || 0;
      const qty = Number(r[CONFIG.SALES_COLUMNS.QUANTITY]) || 0;
      const pay = r[CONFIG.SALES_COLUMNS.PAYMENT_METHOD] || "不明";
      summary.totalAmount += amt;
      summary.totalTransactions += 1;
      summary.totalItems += qty;
      summary.byPayment[pay] = (summary.byPayment[pay] || 0) + amt;
    });

    const sheetName = "月次サマリー";
    const headers = [
      "年月",
      "総売上",
      "取引件数",
      "販売個数",
      "現金売上",
      "QR売上",
      "前月総売上",
      "前月比(額)",
      "前月比(%)",
      "累計売上",
    ];
    if (!this.db.hasSheet(sheetName)) this.db.createSheet(sheetName, headers);
    const sheet = this.db.getSheet(sheetName);
    // ヘッダー整合
    if (sheet.getLastRow() === 0)
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1);
    const indexByYm = new Map();
    rows.forEach((r, i) => indexByYm.set(r[0], i + 2));

    const cash = summary.byPayment["現金"] || 0;
    const qr = summary.byPayment["QR"] || 0;

    // 前月値
    const y = Number(ym.substring(0, 4));
    const m = Number(ym.substring(5, 7));
    const prev = new Date(y, m - 2, 1); // JS月-1, 前月= -2
    const prevYm = Utilities.formatDate(prev, tz, "yyyy-MM");
    const prevRow = indexByYm.get(prevYm);
    let prevTotal = 0;
    if (prevRow) prevTotal = Number(sheet.getRange(prevRow, 2).getValue()) || 0;

    const diffAmt = summary.totalAmount - prevTotal;
    const diffPct =
      prevTotal === 0 ? "" : Math.round((diffAmt / prevTotal) * 1000) / 10;

    const rowValues = [
      summary.ym,
      summary.totalAmount,
      summary.totalTransactions,
      summary.totalItems,
      cash,
      qr,
      prevTotal,
      diffAmt,
      diffPct,
      "",
    ];
    const writeRow = indexByYm.get(summary.ym);
    if (writeRow) {
      sheet.getRange(writeRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      const last = sheet.getLastRow();
      sheet.getRange(last + 1, 1, 1, rowValues.length).setValues([rowValues]);
    }

    // 累計再計算（年月昇順）
    const data = sheet
      .getDataRange()
      .getValues()
      .slice(1)
      .filter((r) => r[0]);
    data.sort((a, b) => (a[0] < b[0] ? -1 : 1));
    let running = 0;
    for (let i = 0; i < data.length; i++) {
      running += Number(data[i][1]) || 0;
      data[i][9 - 1] = running;
    }
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    if (data.length > 0)
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

    sheet.autoResizeColumns(1, headers.length);
    return { sheetName };
  }

  /**
   * 月次サマリーを全期間で再構築
   */
  rebuildMonthlySummaryAll() {
    const tz = CONFIG.APP.TIMEZONE;
    const sales = this.db.getData(CONFIG.SHEETS.SALES);
    const months = new Set();
    sales.forEach((r) => {
      const ts = r[CONFIG.SALES_COLUMNS.TIMESTAMP];
      if (!ts) return;
      const d = ts instanceof Date ? ts : new Date(ts);
      months.add(Utilities.formatDate(d, tz, "yyyy-MM"));
    });
    const list = [...months].sort();
    list.forEach((ym) => {
      const parts = ym.split("-");
      const d = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
      this.generateMonthlySummary(d);
    });
    return { months: list.length };
  }

  /**
   * カテゴリ別月次（全期間）
   */
  generateCategoryMonthlySummary() {
    const tz = CONFIG.APP.TIMEZONE;
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const sales = this.db.getData(CONFIG.SHEETS.SALES);
    const pid2cat = new Map(
      products.map((p) => [
        p[CONFIG.PRODUCT_COLUMNS.ID],
        p[CONFIG.PRODUCT_COLUMNS.CATEGORY] || "未分類",
      ]),
    );
    const table = new Map(); // key: ym|cat -> {qty, amount}
    sales.forEach((r) => {
      const ts = r[CONFIG.SALES_COLUMNS.TIMESTAMP];
      if (!ts) return;
      const d = ts instanceof Date ? ts : new Date(ts);
      const ym = Utilities.formatDate(d, tz, "yyyy-MM");
      const pid = r[CONFIG.SALES_COLUMNS.PRODUCT_ID];
      const cat = pid2cat.get(pid) || "未分類";
      const key = ym + "|" + cat;
      const qty = Number(r[CONFIG.SALES_COLUMNS.QUANTITY]) || 0;
      const amt = Number(r[CONFIG.SALES_COLUMNS.SUBTOTAL]) || 0;
      const cur = table.get(key) || { qty: 0, amount: 0 };
      cur.qty += qty;
      cur.amount += amt;
      table.set(key, cur);
    });

    const rows = [];
    table.forEach((v, k) => {
      const [ym, cat] = k.split("|");
      rows.push([ym, cat, v.qty, v.amount]);
    });
    rows.sort((a, b) =>
      a[0] === b[0] ? (a[1] < b[1] ? -1 : 1) : a[0] < b[0] ? -1 : 1,
    );

    const name = "カテゴリ別月次";
    const headers = ["年月", "カテゴリ", "数量", "売上"];
    if (!this.db.hasSheet(name)) this.db.createSheet(name, headers);
    else {
      const sheet = this.db.getSheet(name);
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    const sheet = this.db.getSheet(name);
    if (rows.length > 0)
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
    return { sheetName: name, rows: rows.length };
  }

  /**
   * 日次レポート（指定日）のPDFを書き出し
   */
  exportDailyReportPdf(date = new Date()) {
    const tz = CONFIG.APP.TIMEZONE;
    const sdate = Utilities.formatDate(date, tz, "yyyyMMdd");
    const name = `日次集計_${sdate}`;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      this.generateDailyReport(date);
      sheet = ss.getSheetByName(name);
    }
    if (!sheet) throw new Error("日次集計シートが見つかりません");
    const gid = sheet.getSheetId();
    const exportUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&portrait=true&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenum=RIGHT&gridlines=false&fzr=false&gid=${gid}`;
    const options = {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    };
    const blob = UrlFetchApp.fetch(exportUrl, options)
      .getBlob()
      .setName(name + ".pdf");
    const folderName = "日次レポートPDF";
    const folder = (function () {
      const it = DriveApp.getFoldersByName(folderName);
      return it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    })();
    const file = folder.createFile(blob);
    return {
      fileId: file.getId(),
      url: file.getUrl(),
      folder: folder.getName(),
    };
  }
}

// ===== InventoryAlertService.js =====

/**
 * 在庫アラートサービス
 * @version 1.0.0
 * @description 在庫僅少通知、自動発注リスト生成、売り切れ予測機能
 */
class InventoryAlertService {
  constructor(database, inventoryService, salesService) {
    this.db = database;
    this.inventoryService = inventoryService;
    this.salesService = salesService;
    this.alertCache = new Map();
  }

  /**
   * 在庫アラートをチェック
   * @returns {Array} アラート情報の配列
   */
  checkInventoryAlerts() {
    try {
      const alerts = [];

      // 在庫データ取得
      const inventory = this.inventoryService.getInventory();
      if (!inventory || inventory.length === 0) {
        console.log("No inventory data found");
        return alerts;
      }

      // 販売履歴取得
      const salesHistory = this.getSalesHistory();

      inventory.forEach((item) => {
        const stock = item[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK];
        const productId = item[CONFIG.INVENTORY_COLUMNS.ID];
        const productName = item[CONFIG.INVENTORY_COLUMNS.NAME];

        // null/undefined チェック
        if (stock === null || stock === undefined) {
          console.log(`Stock is null for product ${productId}`);
          return;
        }

        // 在庫レベルチェック
        const stockLevel = this.getStockLevel(stock);

        // 売上速度の計算
        const salesRate = this.calculateSalesRate(productId, salesHistory);

        // 売り切れ予測時間
        const timeToSoldOut = this.predictTimeToSoldOut(stock, salesRate);

        // アラート生成
        if (stockLevel !== CONFIG.INVENTORY_ALERT.ALERT_LEVELS.INFO) {
          alerts.push({
            productId,
            productName,
            currentStock: stock,
            level: stockLevel,
            salesRate: salesRate.toFixed(2),
            timeToSoldOut,
            reorderQuantity: this.calculateReorderQuantity(
              productId,
              salesRate,
            ),
            timestamp: new Date(),
          });
        }
      });

      // キャッシュに保存
      this.alertCache.set("lastCheck", alerts);
      this.alertCache.set("lastCheckTime", new Date());

      console.log(`Alert check completed. Found ${alerts.length} alerts`);
      return alerts;
    } catch (error) {
      console.error("Error in checkInventoryAlerts:", error);
      throw error;
    }
  }

  /**
   * 在庫レベルを判定
   */
  getStockLevel(stock) {
    if (stock <= 0) {
      return CONFIG.INVENTORY_ALERT.ALERT_LEVELS.SOLD_OUT;
    } else if (stock <= CONFIG.INVENTORY_ALERT.CRITICAL_STOCK_THRESHOLD) {
      return CONFIG.INVENTORY_ALERT.ALERT_LEVELS.CRITICAL;
    } else if (stock <= CONFIG.INVENTORY_ALERT.LOW_STOCK_THRESHOLD) {
      return CONFIG.INVENTORY_ALERT.ALERT_LEVELS.WARNING;
    }
    return CONFIG.INVENTORY_ALERT.ALERT_LEVELS.INFO;
  }

  /**
   * 売上速度を計算（個/時間）
   */
  calculateSalesRate(productId, salesHistory) {
    const recentSales = salesHistory.filter(
      (sale) => sale[CONFIG.SALES_COLUMNS.PRODUCT_ID] === productId,
    );

    if (recentSales.length === 0) return 0;

    const totalQuantity = recentSales.reduce(
      (sum, sale) => sum + sale[CONFIG.SALES_COLUMNS.QUANTITY],
      0,
    );

    return totalQuantity / CONFIG.INVENTORY_ALERT.SALES_HISTORY_HOURS;
  }

  /**
   * 売り切れまでの予測時間（分）
   */
  predictTimeToSoldOut(stock, salesRate) {
    if (salesRate === 0) return null;
    if (stock <= 0) return 0;

    const hours = stock / salesRate;
    const minutes = Math.round(hours * 60);

    if (minutes < 60) {
      return `${minutes}分`;
    } else if (minutes < 1440) {
      const h = Math.floor(minutes / 60);
      const m = minutes % 60;
      return m > 0 ? `${h}時間${m}分` : `${h}時間`;
    } else {
      return "24時間以上";
    }
  }

  /**
   * 推奨発注数量を計算
   */
  calculateReorderQuantity(productId, salesRate) {
    // 過去の初期在庫を取得
    const products = this.db.getData(CONFIG.SHEETS.PRODUCTS);
    const product = products.find(
      (p) => p[CONFIG.PRODUCT_COLUMNS.ID] === productId,
    );

    if (!product) return 0;

    const initialStock = product[CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK] || 50;
    const estimatedDailySales = salesRate * 8; // 8時間営業想定

    // 発注点 = 1日の予測売上 × 係数
    const reorderPoint = Math.ceil(
      estimatedDailySales * CONFIG.INVENTORY_ALERT.REORDER_POINT_MULTIPLIER,
    );

    // 推奨発注数 = 初期在庫 - 現在庫
    return Math.max(0, Math.min(initialStock, reorderPoint));
  }

  /**
   * 最近の販売履歴を取得
   */
  getSalesHistory() {
    const sales = this.db.getData(CONFIG.SHEETS.SALES);
    const cutoffTime = new Date();
    cutoffTime.setHours(
      cutoffTime.getHours() - CONFIG.INVENTORY_ALERT.SALES_HISTORY_HOURS,
    );

    return sales.filter((sale) => {
      const saleTime = new Date(sale[CONFIG.SALES_COLUMNS.TIMESTAMP]);
      return saleTime >= cutoffTime;
    });
  }

  /**
   * 自動発注リストを生成
   */
  generateReorderList() {
    const alerts = this.checkInventoryAlerts();
    const reorderList = [];

    alerts.forEach((alert) => {
      if (
        alert.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.CRITICAL ||
        alert.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.SOLD_OUT
      ) {
        reorderList.push({
          productId: alert.productId,
          productName: alert.productName,
          currentStock: alert.currentStock,
          recommendedQuantity: alert.reorderQuantity,
          urgency:
            alert.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.SOLD_OUT
              ? "緊急"
              : "高",
          estimatedSoldOut: alert.timeToSoldOut,
        });
      }
    });

    return reorderList;
  }

  /**
   * アラート履歴を保存
   */
  saveAlertHistory(alerts) {
    const sheetName = "アラート履歴";

    // シートが存在しない場合は作成
    if (!this.db.hasSheet(sheetName)) {
      const headers = [
        "タイムスタンプ",
        "商品ID",
        "商品名",
        "現在庫",
        "アラートレベル",
        "売上速度",
        "売り切れ予測",
        "推奨発注数",
      ];
      this.db.createSheet(sheetName, headers);
    }

    // アラートデータを整形
    const alertRecords = alerts.map((alert) => ({
      タイムスタンプ: alert.timestamp,
      商品ID: alert.productId,
      商品名: alert.productName,
      現在庫: alert.currentStock,
      アラートレベル: alert.level,
      売上速度: alert.salesRate,
      売り切れ予測: alert.timeToSoldOut || "-",
      推奨発注数: alert.reorderQuantity,
    }));

    if (alertRecords.length > 0) {
      this.db.addData(sheetName, alertRecords);
    }
  }

  /**
   * ダッシュボードにアラート情報を追加
   */
  updateDashboardWithAlerts(alerts) {
    const sheet = this.db.getSheet(CONFIG.SHEETS.DASHBOARD);
    if (!sheet) return;

    // アラートサマリーを作成
    const criticalCount = alerts.filter(
      (a) => a.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.CRITICAL,
    ).length;

    const warningCount = alerts.filter(
      (a) => a.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.WARNING,
    ).length;

    const soldOutCount = alerts.filter(
      (a) => a.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.SOLD_OUT,
    ).length;

    // ダッシュボードの特定位置に書き込み
    const startRow = 10; // アラート情報開始行
    sheet.getRange(startRow, 1).setValue("在庫アラート");
    sheet.getRange(startRow + 1, 1).setValue("売り切れ");
    sheet.getRange(startRow + 1, 2).setValue(soldOutCount);
    sheet.getRange(startRow + 2, 1).setValue("危険在庫");
    sheet.getRange(startRow + 2, 2).setValue(criticalCount);
    sheet.getRange(startRow + 3, 1).setValue("在庫僅少");
    sheet.getRange(startRow + 3, 2).setValue(warningCount);
    sheet.getRange(startRow + 4, 1).setValue("最終チェック");
    sheet.getRange(startRow + 4, 2).setValue(new Date());
  }

  /**
   * 定期実行用のトリガー処理
   */
  runScheduledCheck() {
    try {
      const alerts = this.checkInventoryAlerts();

      // アラート履歴を保存
      if (alerts.length > 0) {
        this.saveAlertHistory(alerts);
      }

      // ダッシュボードを更新
      this.updateDashboardWithAlerts(alerts);

      // 危険レベルのアラートがある場合は通知（将来の拡張用）
      const criticalAlerts = alerts.filter(
        (a) =>
          a.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.CRITICAL ||
          a.level === CONFIG.INVENTORY_ALERT.ALERT_LEVELS.SOLD_OUT,
      );

      if (criticalAlerts.length > 0) {
        // ここにメール送信やSlack通知などを実装可能
        console.log(`Critical alerts: ${criticalAlerts.length} items`);
      }

      return {
        success: true,
        alertCount: alerts.length,
        criticalCount: criticalAlerts.length,
      };
    } catch (error) {
      console.error("Inventory alert check error:", error);
      return {
        success: false,
        error: error.message,
      };
    }
  }
}

// ===== GeminiService.js =====

/**
 * Gemini API連携サービス
 * @version 1.1.0
 */
class GeminiService {
  constructor() {
    this.apiKey = null;
  }

  getApiKey() {
    if (!this.apiKey) {
      this.apiKey =
        PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    }
    return this.apiKey;
  }

  setApiKey(apiKey) {
    PropertiesService.getScriptProperties().setProperty(
      "GEMINI_API_KEY",
      apiKey,
    );
    this.apiKey = apiKey;
  }

  processVoiceOrder(voiceText, products) {
    const apiKey = this.getApiKey();

    // APIキーがない場合は簡易パーサーを使用
    if (!apiKey) {
      return this.processWithSimpleParser(voiceText, products);
    }

    try {
      const functionCalls = this.callGeminiAPI(voiceText, products);
      return this.processFunctionCalls(functionCalls, products);
    } catch (error) {
      console.error("Gemini API error:", error);
      // エラー時は簡易パーサーにフォールバック
      return this.processWithSimpleParser(voiceText, products);
    }
  }

  callGeminiAPI(userInput, products) {
    const apiKey = this.getApiKey();
    if (!apiKey) {
      throw new Error("APIキーが設定されていません");
    }

    // 商品カタログを準備
    const productCatalog = products.map((p) => ({
      商品名: p[CONFIG.PRODUCT_COLUMNS.NAME],
      説明: p[CONFIG.PRODUCT_COLUMNS.DESCRIPTION] || "",
      価格: p[CONFIG.PRODUCT_COLUMNS.PRICE],
      在庫: p[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK],
    }));

    // APIリクエストを構築
    const payload = this.buildApiPayload(userInput, productCatalog);

    // API呼び出し
    const url = `${CONFIG.API.GEMINI.BASE_URL}/${CONFIG.API.GEMINI.MODEL}:${CONFIG.API.GEMINI.ENDPOINT}?key=${apiKey}`;

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      console.error("Gemini API error response:", result.error);
      throw new Error(`API error: ${result.error.message}`);
    }

    return this.extractFunctionCalls(result);
  }

  buildApiPayload(userInput, productCatalog) {
    return {
      contents: [
        {
          parts: [
            {
              text: `あなたは文化祭の販売システムのアシスタントです。以下の商品カタログを参考に、ユーザーの注文を解析してください。

商品カタログ:
${JSON.stringify(productCatalog, null, 2)}

ユーザーの注文: "${userInput}"

注意事項:
- 商品名は正確に一致させてください
- 数量が明示されていない場合は1つとしてください
- 在庫を超える注文はしないでください
- 複数の商品が含まれる場合は、それぞれに対してaddItemToCartを呼び出してください`,
            },
          ],
        },
      ],
      tools: [
        {
          function_declarations: [
            {
              name: "addItemToCart",
              description:
                "注文カートに商品を追加する。商品名と数量を特定できた場合に呼び出す",
              parameters: {
                type: "object",
                properties: {
                  productName: {
                    type: "string",
                    description: "商品マスタに存在する商品名",
                  },
                  quantity: {
                    type: "integer",
                    description: "注文数量",
                  },
                },
                required: ["productName", "quantity"],
              },
            },
          ],
        },
      ],
      generationConfig: {
        temperature: CONFIG.API.GEMINI.TEMPERATURE,
        maxOutputTokens: CONFIG.API.GEMINI.MAX_OUTPUT_TOKENS,
      },
    };
  }

  extractFunctionCalls(result) {
    if (
      !result.candidates ||
      !result.candidates[0] ||
      !result.candidates[0].content
    ) {
      return [];
    }

    const parts = result.candidates[0].content.parts || [];
    return parts
      .filter((part) => part.functionCall)
      .map((part) => part.functionCall);
  }

  processFunctionCalls(functionCalls, products) {
    const orderItems = [];

    functionCalls.forEach((call) => {
      if (call.name === "addItemToCart" && call.args) {
        const productName = call.args.productName;
        const quantity = call.args.quantity || 1;

        const product = products.find(
          (p) => p[CONFIG.PRODUCT_COLUMNS.NAME] === productName,
        );

        if (
          product &&
          quantity > 0 &&
          quantity <= product[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]
        ) {
          orderItems.push({
            ...product,
            数量: quantity,
          });
        }
      }
    });

    if (orderItems.length > 0) {
      const itemDescriptions = orderItems
        .map((item) => `${item[CONFIG.PRODUCT_COLUMNS.NAME]}×${item["数量"]}`)
        .join("、");

      return {
        success: true,
        items: orderItems,
        message: `${itemDescriptions}を追加しました`,
      };
    } else {
      return {
        success: false,
        message: "在庫不足または商品が見つかりませんでした",
      };
    }
  }

  processWithSimpleParser(voiceText, products) {
    const orderItems = [];
    const quantityPattern = /(\d+)\s*(?:つ|個|本|杯|枚)/g;

    products.forEach((product) => {
      const productName = product[CONFIG.PRODUCT_COLUMNS.NAME];
      const description = product[CONFIG.PRODUCT_COLUMNS.DESCRIPTION] || "";
      const searchTerms = [productName, ...description.split("、")];

      for (const term of searchTerms) {
        if (term && voiceText.includes(term)) {
          let quantity = 1;

          // 数量を探す
          const matches = voiceText.matchAll(quantityPattern);
          for (const match of matches) {
            const matchIndex = match.index;
            const termIndex = voiceText.indexOf(term);

            // 数量と商品名が近い場合（前後50文字以内）
            if (Math.abs(matchIndex - termIndex) < 50) {
              quantity = parseInt(match[1]);
              break;
            }
          }

          if (
            quantity > 0 &&
            quantity <= product[CONFIG.INVENTORY_COLUMNS.CURRENT_STOCK]
          ) {
            orderItems.push({
              ...product,
              数量: quantity,
            });
            break; // 同じ商品を重複して追加しない
          }
        }
      }
    });

    if (orderItems.length > 0) {
      const itemDescriptions = orderItems
        .map((item) => `${item[CONFIG.PRODUCT_COLUMNS.NAME]}×${item["数量"]}`)
        .join("、");

      return {
        success: true,
        items: orderItems,
        message: `${itemDescriptions}を追加しました`,
      };
    } else {
      return {
        success: false,
        message:
          "商品が見つかりませんでした。商品名をはっきりと話してください。",
      };
    }
  }
}

// ===== ServiceInitializer.js =====

/**
 * サービス初期化
 * @version 1.1.0
 * @description すべてのサービスインスタンスを正しい順序で初期化
 */

// データベースインスタンス
const db = new Database();

// サービスインスタンス
const productService = new ProductService(db);
const inventoryService = new InventoryService(db);
const salesService = new SalesService(db, inventoryService);
const dailyOpsService = new DailyOpsService(db, salesService, inventoryService);
const geminiService = new GeminiService();
const inventoryAlertService = new InventoryAlertService(
  db,
  inventoryService,
  salesService,
);

// ===== Main.js =====

/**
 * メインアプリケーション
 * @version 1.1.0
 * @description サービスクラスを統合し、GAS固有の関数を定義
 */

const SAMPLE_PRODUCTS = [
  [
    "P001",
    "焼きそば",
    350,
    "麺類",
    100,
    "https://example.com/yakisoba.jpg",
    "やきそば、ヤキソバ、ソース焼きそば",
    true,
  ],
  [
    "P002",
    "たこ焼き",
    400,
    "粉もの",
    80,
    "https://example.com/takoyaki.jpg",
    "たこやき、タコヤキ、たこ焼",
    true,
  ],
  [
    "P003",
    "フランクフルト",
    300,
    "肉類",
    120,
    "",
    "フランク、ソーセージ",
    true,
  ],
  [
    "P004",
    "からあげ",
    400,
    "肉類",
    100,
    "",
    "唐揚げ、から揚げ、カラアゲ、鶏肉",
    true,
  ],
  [
    "P005",
    "ポテトフライ",
    250,
    "揚げ物",
    150,
    "",
    "ポテト、フライドポテト、フライ",
    true,
  ],
  [
    "P006",
    "チョコバナナ",
    200,
    "デザート",
    80,
    "",
    "チョコ、バナナ、デザート",
    true,
  ],
  ["P007", "かき氷", 300, "デザート", 100, "", "カキ氷、氷、シロップ", true],
  [
    "P008",
    "わたあめ",
    200,
    "デザート",
    60,
    "",
    "綿あめ、わたがし、綿菓子",
    true,
  ],
  ["P009", "ラムネ", 150, "飲み物", 200, "", "らむね、炭酸", true],
  [
    "P010",
    "オレンジジュース",
    150,
    "飲み物",
    150,
    "",
    "オレンジ、ジュース、果汁",
    true,
  ],
  ["P011", "お茶", 150, "飲み物", 200, "", "緑茶、日本茶、ペットボトル", true],
  ["P012", "コーラ", 150, "飲み物", 100, "", "コカコーラ、炭酸、cola", true],
  [
    "P013",
    "クレープ",
    500,
    "デザート",
    50,
    "",
    "クレープ、生クリーム、フルーツ",
    true,
  ],
  [
    "P014",
    "ホットドッグ",
    350,
    "パン類",
    80,
    "",
    "ホットドック、パン、ソーセージ",
    true,
  ],
  [
    "P015",
    "チュロス",
    300,
    "デザート",
    60,
    "",
    "チュロス、シナモン、砂糖",
    true,
  ],
];
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🎪 POSシステム設定")
    .addItem("📋 初期設定（シート作成）", "initializeSpreadsheet")
    .addItem("🎯 ダミーデータ生成（テスト用）", "generateDummyData")
    .addSeparator()
    .addItem("🔑 Gemini APIキー設定", "showApiKeyDialog")
    .addItem("🌐 WebアプリURL設定", "setWebAppUrl")
    .addItem("🚀 Webアプリを開く", "openWebApp")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("🗓 日次作業")
        .addItem("📊 日次集計（本日）", "runDailySummaryToday")
        .addItem("📊 日次集計（任意日）", "runDailySummaryPickDate")
        .addItem("🗃 販売履歴アーカイブ（任意日）", "archiveSalesPickDate")
        .addItem("🗃 販売履歴アーカイブ（本日）", "archiveSalesToday")
        .addItem("🔁 在庫準備：初期在庫へ補充", "restockInventoryToInitial")
        .addItem("📄 日次レポートPDF（本日）", "exportDailyReportPdfToday")
        .addItem(
          "📊+🗃+🔁 締め処理（集計+アーカイブ+在庫準備）",
          "runDailyClosingWithArchive",
        ),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("📅 月次・分析")
        .addItem("📅 月次サマリー（今月）", "runMonthlySummaryThisMonth")
        .addItem("📅 月次サマリー（全期間再構築）", "rebuildMonthlySummaryAll")
        .addItem("📊 カテゴリ別月次（全期間）", "buildCategoryMonthlySummary"),
    )
    .addSubMenu(
      ui
        .createMenu("📦 在庫管理")
        .addItem("⚠️ 在庫アラートチェック", "runInventoryAlertCheck")
        .addItem("📋 発注リスト生成", "showReorderListInSheet"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("📚 ドキュメント")
        .addItem("📖 開発マニュアル", "openManual")
        .addItem("🎓 カリキュラム", "openCurriculum"),
    )
    .addSeparator()
    .addItem("📊 ダッシュボード更新", "updateDashboard")
    .addItem("⚙️ トリガー設定", "setupTriggers")
    .addItem("🗑️ データ全削除（注意）", "clearAllData")
    .addToUi();
}

function initializeSpreadsheet(showAlert = true) {
  const ss = db.getSpreadsheet();

  // 商品マスタシートの作成
  if (!db.hasSheet(CONFIG.SHEETS.PRODUCTS)) {
    const headers = Object.values(CONFIG.PRODUCT_COLUMNS);
    const sheet = db.createSheet(CONFIG.SHEETS.PRODUCTS, headers);

    // サンプルデータの初期5件を追加
    const initialData = SAMPLE_PRODUCTS.slice(0, 5);
    sheet
      .getRange(2, 1, initialData.length, initialData[0].length)
      .setValues(initialData);
  }

  // 在庫管理シートの作成
  if (!db.hasSheet(CONFIG.SHEETS.INVENTORY)) {
    const headers = Object.values(CONFIG.INVENTORY_COLUMNS);
    const sheet = db.createSheet(CONFIG.SHEETS.INVENTORY, headers);

    // 商品マスタからデータを生成
    const products = db.getData(CONFIG.SHEETS.PRODUCTS);
    const inventoryData = [];
    const formulas = [];

    products.forEach((product, index) => {
      inventoryData.push([
        product[CONFIG.PRODUCT_COLUMNS.ID],
        "",
        product[CONFIG.PRODUCT_COLUMNS.INITIAL_STOCK],
      ]);
      formulas.push([
        `=IFERROR(VLOOKUP(A${index + 2},'${CONFIG.SHEETS.PRODUCTS}'!A:B,2,FALSE),"")`,
      ]);
    });

    if (inventoryData.length > 0) {
      sheet.getRange(2, 1, inventoryData.length, 3).setValues(inventoryData);
      sheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    }
  } else {
    // 既存シートの場合でも不足行があれば補充
    try {
      inventoryService.syncInventoryWithProducts();
    } catch (e) {
      console.log("在庫シート同期をスキップ", e);
    }
  }

  // 販売履歴シートの作成
  if (!db.hasSheet(CONFIG.SHEETS.SALES)) {
    const headers = Object.values(CONFIG.SALES_COLUMNS);
    db.createSheet(CONFIG.SHEETS.SALES, headers);
  }

  // 設定シートの作成
  if (!db.hasSheet(CONFIG.SHEETS.SETTINGS)) {
    const sheet = db.createSheet(CONFIG.SHEETS.SETTINGS, ["キー", "値"]);

    // デフォルト設定
    const defaultSettings = [
      [CONFIG.SETTINGS_KEYS.SHOP_NAME, CONFIG.DEFAULTS.SHOP_NAME],
      [CONFIG.SETTINGS_KEYS.QR_PAYMENT_NAME, "PayPay"],
      [CONFIG.SETTINGS_KEYS.QR_CODE_IMAGE_URL, ""],
      [CONFIG.SETTINGS_KEYS.LOCATION, "岡崎市"],
      [CONFIG.SETTINGS_KEYS.WEATHER_API_KEY, ""],
    ];
    sheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }

  // ダッシュボードシートの作成
  if (!db.hasSheet(CONFIG.SHEETS.DASHBOARD)) {
    const sheet = db.createSheet(CONFIG.SHEETS.DASHBOARD, []);
    sheet.getRange(1, 1).setValue("売上総額");
    sheet.getRange(1, 2).setValue(0);
    sheet.getRange(3, 1).setValue("商品別売上");
  }

  console.log("スプレッドシートの初期化が完了しました");

  // UIに完了メッセージを表示（メニューから呼ばれた場合のみ）
  if (showAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        "✅ 初期設定完了",
        "スプレッドシートの初期化が完了しました。\n\n次は「POSシステム設定」メニューから「ダミーデータ生成」を実行してください。",
        SpreadsheetApp.getUi().ButtonSet.OK,
      );
    } catch (e) {
      // Webコンテキストからの呼び出しの場合はエラーを無視
      console.log("UIアラート表示をスキップ（Webコンテキスト）");
    }
  }
}

function generateDummyData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "ダミーデータ生成",
    "テスト用のダミーデータを生成しますか？\n\n※既存のデータは上書きされます",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // 商品マスタのダミーデータ
  db.clearData(CONFIG.SHEETS.PRODUCTS);

  // 商品データを追加
  const productRecords = SAMPLE_PRODUCTS.map((data) => {
    const record = {};
    Object.values(CONFIG.PRODUCT_COLUMNS).forEach((col, index) => {
      record[col] = data[index];
    });
    return record;
  });

  db.addData(CONFIG.SHEETS.PRODUCTS, productRecords);

  // 在庫管理データを再構築・初期化
  db.clearData(CONFIG.SHEETS.INVENTORY);
  inventoryService.syncInventoryWithProducts();
  inventoryService.resetInventory();

  // 設定データを拡張
  db.clearData(CONFIG.SHEETS.SETTINGS);
  const settingsData = [
    [CONFIG.SETTINGS_KEYS.SHOP_NAME, "3年A組 文化祭模擬店"],
    [CONFIG.SETTINGS_KEYS.QR_PAYMENT_NAME, "PayPay"],
    [CONFIG.SETTINGS_KEYS.QR_CODE_IMAGE_URL, ""],
    [CONFIG.SETTINGS_KEYS.SALES_TARGET, "50000"],
    [CONFIG.SETTINGS_KEYS.OPENING_TIME, "10:00"],
    [CONFIG.SETTINGS_KEYS.CLOSING_TIME, "15:00"],
    [CONFIG.SETTINGS_KEYS.STAFF_COUNT, "8"],
    [CONFIG.SETTINGS_KEYS.CURRENCY_SYMBOL, "¥"],
    [CONFIG.SETTINGS_KEYS.LOCATION, "岡崎市"],
    [CONFIG.SETTINGS_KEYS.WEATHER_API_KEY, ""],
  ];

  const settingsRecords = settingsData.map(([key, value]) => ({
    キー: key,
    値: value,
  }));

  db.addData(CONFIG.SHEETS.SETTINGS, settingsRecords);

  // サンプル販売履歴
  const now = new Date();
  const sampleSales = [];

  for (let i = 0; i < 20; i++) {
    const timestamp = new Date(now.getTime() - (60 - i * 3) * 60 * 1000);
    const productIndex = Math.floor(Math.random() * 10);
    const product = SAMPLE_PRODUCTS[productIndex];
    const quantity = Math.floor(Math.random() * 3) + 1;

    sampleSales.push({
      [CONFIG.SALES_COLUMNS.TRANSACTION_ID]:
        `T${Utilities.formatDate(timestamp, CONFIG.APP.TIMEZONE, "yyyyMMddHHmmss")}${String(i).padStart(3, "0")}`,
      [CONFIG.SALES_COLUMNS.TIMESTAMP]: timestamp,
      [CONFIG.SALES_COLUMNS.PRODUCT_ID]: product[0],
      [CONFIG.SALES_COLUMNS.PRODUCT_NAME]: product[1],
      [CONFIG.SALES_COLUMNS.PRICE]: product[2],
      [CONFIG.SALES_COLUMNS.QUANTITY]: quantity,
      [CONFIG.SALES_COLUMNS.SUBTOTAL]: product[2] * quantity,
      [CONFIG.SALES_COLUMNS.PAYMENT_METHOD]:
        Math.random() > 0.5 ? "現金" : "QR",
      [CONFIG.SALES_COLUMNS.STAFF_ID]: `STAFF00${(i % 3) + 1}`,
    });
  }

  if (sampleSales.length > 0) {
    db.addData(CONFIG.SHEETS.SALES, sampleSales);
  }

  // ダッシュボードを更新
  salesService.updateDashboard();

  ui.alert(
    "✅ 完了",
    "ダミーデータの生成が完了しました。\n\nWebアプリを開いて動作を確認してください。",
    ui.ButtonSet.OK,
  );
}

function doGet(e) {
  // 現在のURLを保存（Webアプリとして実行されている場合）
  saveCurrentUrl();

  // URLパラメータから表示するページを決定
  const page =
    e && e.parameter && e.parameter.page ? e.parameter.page : "index";

  let htmlFile = "index";
  let title = "文化祭スマートPOS";

  switch (page) {
    case "manual":
      htmlFile = "manual";
      title = "文化祭POSシステム開発ガイド";
      break;
    case "curriculum":
      htmlFile = "curriculum";
      title = "プログラミング学習カリキュラム";
      break;
    default:
      // 初回起動時にシートが存在しない場合は初期化（UIアラートなし）
      if (!db.hasSheet(CONFIG.SHEETS.PRODUCTS)) {
        initializeSpreadsheet(false);
      }
      break;
  }

  const template = HtmlService.createTemplateFromFile(htmlFile);
  return template
    .evaluate()
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * 店舗設定を取得
 */
function getShopSettings() {
  const settings = db.getData(CONFIG.SHEETS.SETTINGS);
  const result = {};

  settings.forEach((row) => {
    if (row["キー"]) {
      result[row["キー"]] = row["値"];
    }
  });

  return result;
}

/**
 * 在庫情報付き商品リストを取得（Webアプリ用）
 */
function getProductsWithStock() {
  return productService.getProductsWithStock();
}

function processTransaction(cartItems, paymentMethod) {
  try {
    return salesService.processTransaction(cartItems, paymentMethod);
  } catch (error) {
    console.error("Transaction error:", error);
    return handleError(error, "processTransaction");
  }
}

function processVoiceOrder(voiceText) {
  try {
    const products = productService.getProductsWithStock();
    return geminiService.processVoiceOrder(voiceText, products);
  } catch (error) {
    console.error("Voice order error:", error);
    return handleError(error, "processVoiceOrder");
  }
}

/**
 * 在庫アラートをチェック（Web UI用）
 */
function getInventoryAlerts() {
  try {
    return inventoryAlertService.checkInventoryAlerts();
  } catch (error) {
    console.error("Inventory alert error:", error);
    return handleError(error, "getInventoryAlerts");
  }
}

/**
 * 自動発注リストを取得
 */
function getReorderList() {
  try {
    return inventoryAlertService.generateReorderList();
  } catch (error) {
    console.error("Reorder list error:", error);
    return handleError(error, "getReorderList");
  }
}

/**
 * 在庫アラートの定期チェック（トリガー用）
 */
function runInventoryAlertCheck() {
  try {
    const result = inventoryAlertService.runScheduledCheck();

    // UIアラート表示（メニューから実行時）
    try {
      const ui = SpreadsheetApp.getUi();
      if (result.success) {
        ui.alert(
          "✅ 在庫アラートチェック完了",
          `チェック完了しました。\n\nアラート総数: ${result.alertCount}件\n危険レベル: ${result.criticalCount}件`,
          ui.ButtonSet.OK,
        );
      }
    } catch (e) {
      // トリガーから実行された場合はUIが使えないので無視
    }

    return result;
  } catch (error) {
    console.error("Scheduled alert check error:", error);
    return handleError(error, "runInventoryAlertCheck");
  }
}

/**
 * 日次集計（本日）
 */
function runDailySummaryToday() {
  const today = new Date();
  const result = dailyOpsService.generateDailyReport(today);
  SpreadsheetApp.getUi().alert(
    "✅ 日次集計",
    `シート「${result.sheetName}」を作成しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * 日次集計（任意日を入力: yyyy-MM-dd）
 */
function runDailySummaryPickDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "日付を入力",
    "yyyy-MM-dd 形式で入力してください（例: 2025-10-14）",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const text = resp.getResponseText().trim();
  const parts = text.split("-");
  if (parts.length !== 3) {
    ui.alert(
      "形式エラー",
      "yyyy-MM-dd 形式で入力してください。",
      ui.ButtonSet.OK,
    );
    return;
  }
  const date = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
  );
  const result = dailyOpsService.generateDailyReport(date);
  ui.alert(
    "✅ 日次集計",
    `シート「${result.sheetName}」を作成しました。`,
    ui.ButtonSet.OK,
  );
}

/** 任意日アーカイブ */
function archiveSalesPickDate() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "日付を入力",
    "yyyy-MM-dd 形式で入力してください（例: 2025-10-14）",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const text = resp.getResponseText().trim();
  const parts = text.split("-");
  if (parts.length !== 3) {
    ui.alert(
      "形式エラー",
      "yyyy-MM-dd 形式で入力してください。",
      ui.ButtonSet.OK,
    );
    return;
  }
  const date = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
  );
  const res = dailyOpsService.archiveSales(date);
  if (res.success) {
    ui.alert(
      "✅ アーカイブ完了",
      `シート「${res.sheetName}」に ${res.archived} 件を移動しました。`,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "情報",
      res.message || "アーカイブ対象がありません。",
      ui.ButtonSet.OK,
    );
  }
}

/** 日次レポートPDF（本日） */
function exportDailyReportPdfToday() {
  try {
    const res = dailyOpsService.exportDailyReportPdf(new Date());
    SpreadsheetApp.getUi().alert(
      "✅ PDF出力",
      `フォルダ「${res.folder}」に保存しました。\nURL: ${res.url}`,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "PDF出力に失敗しました: " + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/** 月次サマリー（今月） */
function runMonthlySummaryThisMonth() {
  const res = dailyOpsService.generateMonthlySummary(new Date());
  SpreadsheetApp.getUi().alert(
    "✅ 月次サマリー",
    `シート「${res.sheetName}」を更新しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/** 月次サマリー（全期間再構築） */
function rebuildMonthlySummaryAll() {
  const res = dailyOpsService.rebuildMonthlySummaryAll();
  SpreadsheetApp.getUi().alert(
    "✅ 月次サマリー",
    `全期間 ${res.months} ヶ月分を再構築しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/** カテゴリ別月次（全期間） */
function buildCategoryMonthlySummary() {
  const res = dailyOpsService.generateCategoryMonthlySummary();
  SpreadsheetApp.getUi().alert(
    "✅ カテゴリ別月次",
    `シート「${res.sheetName}」に ${res.rows} 行を書き出しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * 在庫を初期在庫へ補充（リセット）
 */
function restockInventoryToInitial() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "在庫準備",
    "在庫を商品ごとの初期在庫にリセットします。よろしいですか？",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;
  dailyOpsService.restockToInitial();
  ui.alert("✅ 完了", "在庫を初期在庫にリセットしました。", ui.ButtonSet.OK);
}

/**
 * 日次集計 + 在庫準備（本日）
 */
function runDailyClosing() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const result = dailyOpsService.generateDailyReport(today);
  const confirm = ui.alert(
    "在庫準備",
    "日次集計を出力しました。続けて在庫を初期在庫にリセットしますか？",
    ui.ButtonSet.YES_NO,
  );
  if (confirm === ui.Button.YES) {
    dailyOpsService.restockToInitial();
    ui.alert(
      "✅ 完了",
      `「${result.sheetName}」を作成し、在庫を初期在庫にリセットしました。`,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "✅ 完了",
      `「${result.sheetName}」を作成しました（在庫は変更していません）。`,
      ui.ButtonSet.OK,
    );
  }
}

/**
 * 販売履歴アーカイブ（本日）
 */
function archiveSalesToday() {
  const res = dailyOpsService.archiveSales(new Date());
  const ui = SpreadsheetApp.getUi();
  if (res.success) {
    ui.alert(
      "✅ アーカイブ完了",
      `シート「${res.sheetName}」に ${res.archived} 件を移動しました。`,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(
      "情報",
      res.message || "アーカイブ対象がありません。",
      ui.ButtonSet.OK,
    );
  }
}

/**
 * 締め処理（集計+アーカイブ+在庫準備）
 */
function runDailyClosingWithArchive() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const report = dailyOpsService.generateDailyReport(today);
  const arch = dailyOpsService.archiveSales(today);
  let msg = `「${report.sheetName}」を作成しました。\n`;
  msg += arch.success
    ? `販売履歴を「${arch.sheetName}」に ${arch.archived} 件移動しました。`
    : arch.message || "販売履歴の移動はありません。";
  const confirm = ui.alert(
    "在庫準備",
    msg + "\n\n続けて在庫を初期在庫にリセットしますか？",
    ui.ButtonSet.YES_NO,
  );
  if (confirm === ui.Button.YES) {
    dailyOpsService.restockToInitial();
    ui.alert(
      "✅ 締め処理完了",
      "在庫を初期在庫にリセットしました。",
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert("✅ 締め処理完了", "在庫は変更していません。", ui.ButtonSet.OK);
  }
}

/**
 * 発注リストをシートに表示
 */
function showReorderListInSheet() {
  try {
    const ui = SpreadsheetApp.getUi();
    const reorderList = inventoryAlertService.generateReorderList();

    if (reorderList.length === 0) {
      ui.alert(
        "📋 発注リスト",
        "現在、発注が必要な商品はありません。",
        ui.ButtonSet.OK,
      );
      return;
    }

    // 新しいシートを作成または既存のシートをクリア
    const sheetName =
      "発注リスト_" +
      Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet(sheetName);

    // ヘッダー設定
    const headers = [
      "商品ID",
      "商品名",
      "現在庫",
      "推奨発注数",
      "緊急度",
      "売り切れ予測",
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#f0f0f0");

    // データ設定
    const data = reorderList.map((item) => [
      item.productId,
      item.productName,
      item.currentStock,
      item.recommendedQuantity,
      item.urgency,
      item.estimatedSoldOut || "-",
    ]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);

      // 緊急度によって色分け
      for (let i = 0; i < data.length; i++) {
        const urgency = data[i][4];
        const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
        if (urgency === "緊急") {
          rowRange.setBackground("#ffe6e6");
        } else if (urgency === "高") {
          rowRange.setBackground("#fff3cd");
        }
      }
    }

    // 列幅調整
    sheet.autoResizeColumns(1, headers.length);

    ui.alert(
      "✅ 発注リスト作成完了",
      `発注リストを「${sheetName}」シートに作成しました。\n\n対象商品: ${reorderList.length}件`,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    console.error("Show reorder list error:", error);
    SpreadsheetApp.getUi().alert(
      "エラー",
      "発注リストの作成中にエラーが発生しました。",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentKey =
    PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  let message = "Gemini APIキーを設定します。\n\n";
  if (currentKey) {
    message += "現在のAPIキー: " + currentKey.substring(0, 10) + "...\n\n";
  } else {
    message += "現在APIキーは設定されていません。\n\n";
  }
  message +=
    "Google AI Studioでキーを取得してください:\nhttps://makersuite.google.com/app/apikey";

  const response = ui.prompt("APIキー設定", message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const newKey = response.getResponseText().trim();
    if (newKey) {
      geminiService.setApiKey(newKey);
      ui.alert("✅ 設定完了", "APIキーを設定しました。", ui.ButtonSet.OK);
    }
  }
}

/**
 * 実際のWebアプリURLを自動取得
 */
function getWebAppUrl() {
  try {
    // まず保存されたURLを確認
    const savedUrl =
      PropertiesService.getScriptProperties().getProperty("WEB_APP_URL");
    if (savedUrl) {
      return savedUrl;
    }

    // ScriptApp.getService().getUrl()を試す
    const serviceUrl = ScriptApp.getService().getUrl();
    if (serviceUrl && serviceUrl.includes("/exec")) {
      // このURLが実際のデプロイURLの可能性が高い
      PropertiesService.getScriptProperties().setProperty(
        "WEB_APP_URL",
        serviceUrl,
      );
      return serviceUrl;
    }

    // もしURLが開発用URLの場合、デプロイIDを使って実際のURLを構築
    const scriptId = ScriptApp.getScriptId();

    // デプロイされたWebアプリのURLパターン
    // 通常は /macros/s/[DEPLOYMENT_ID]/exec の形式

    // 現在実行中のコンテキストからURLを取得する別の方法
    try {
      // この関数を実行しているユーザーのセッションから情報を取得
      const activeUrl = HtmlService.createHtmlOutput("").getUrl();
      if (activeUrl && activeUrl.includes("/exec")) {
        PropertiesService.getScriptProperties().setProperty(
          "WEB_APP_URL",
          activeUrl,
        );
        return activeUrl;
      }
    } catch (e) {
      // この方法が使えない場合は次へ
    }

    // 最終手段：nullを返して手動設定を促す
    return null;
  } catch (e) {
    console.error("URL取得エラー:", e);
    return null;
  }
}

/**
 * doGet実行時にURLを保存
 */
function saveCurrentUrl() {
  // この関数はdoGetから呼び出される
  const url = ScriptApp.getService().getUrl();
  if (url && url.includes("/exec")) {
    PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", url);
  }
}

/**
 * WebアプリURLを設定（デプロイ後に一度実行）
 */
function setWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "WebアプリURL設定",
    "デプロイされたWebアプリのURLを入力してください：",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const url = response.getResponseText().trim();
    if (url) {
      PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", url);
      ui.alert("✅ 設定完了", "WebアプリURLを設定しました。", ui.ButtonSet.OK);
    }
  }
}

/**
 * Webアプリを開く
 */
function openWebApp() {
  // URLを自動取得
  let url = getWebAppUrl();

  // 自動取得に失敗した場合は手動設定を促す
  if (!url) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "⚠️ URL自動取得失敗",
      "WebアプリのURLを自動取得できませんでした。\n\n手動でURLを設定しますか？",
      ui.ButtonSet.YES_NO,
    );

    if (response === ui.Button.YES) {
      setWebAppUrl();
      url = getWebAppUrl();
    }
  }

  if (url) {
    // 自動的に新しいタブで開く
    const html = `
      <div style="font-family: sans-serif; padding: 20px;">
        <h3>Webアプリを開く</h3>
        <p>Webアプリを新しいタブで開いています...</p>
        <p><a href="${url}" target="_blank" style="color: #1a73e8; font-size: 14px;">${url}</a></p>
        <br>
        <p style="color: #666; font-size: 12px;">
          ※自動的に開かない場合は、上記のリンクをクリックしてください<br>
          ※初回アクセス時は承認が必要です
        </p>
      </div>
      <script>
        window.open('${url}', '_blank');
      </script>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(450)
      .setHeight(200);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Webアプリ URL");
  } else {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "Webアプリがまだデプロイされていません。\n\n「デプロイ」→「新しいデプロイ」から公開してください。",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * 開発マニュアルを開く
 */
function openManual() {
  // URLを自動取得
  let url = getWebAppUrl();

  // 自動取得に失敗した場合は手動設定を促す
  if (!url) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "⚠️ URL自動取得失敗",
      "WebアプリのURLを自動取得できませんでした。\n\n手動でURLを設定しますか？",
      ui.ButtonSet.YES_NO,
    );

    if (response === ui.Button.YES) {
      setWebAppUrl();
      url = getWebAppUrl();
    }
  }

  if (url) {
    // WebアプリのURLからmanual.htmlのURLを生成
    const manualUrl = url.replace(/\/exec$/, "/exec?page=manual");

    // 自動的に新しいタブで開く
    const html = `
      <div style="font-family: sans-serif; padding: 20px;">
        <h3>📖 開発マニュアル</h3>
        <p>開発マニュアルを新しいタブで開いています...</p>
        <p><a href="${manualUrl}" target="_blank" style="color: #1a73e8; font-size: 14px;">${manualUrl}</a></p>
        <br>
        <p style="color: #666; font-size: 12px;">
          ※自動的に開かない場合は、上記のリンクをクリックしてください<br>
          ※PDFダウンロード機能付き
        </p>
      </div>
      <script>
        window.open('${manualUrl}', '_blank');
      </script>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(450)
      .setHeight(200);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "開発マニュアル");
  } else {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "Webアプリがまだデプロイされていません。",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * カリキュラムを開く
 */
function openCurriculum() {
  // URLを自動取得
  let url = getWebAppUrl();

  // 自動取得に失敗した場合は手動設定を促す
  if (!url) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "⚠️ URL自動取得失敗",
      "WebアプリのURLを自動取得できませんでした。\n\n手動でURLを設定しますか？",
      ui.ButtonSet.YES_NO,
    );

    if (response === ui.Button.YES) {
      setWebAppUrl();
      url = getWebAppUrl();
    }
  }

  if (url) {
    // WebアプリのURLからcurriculum.htmlのURLを生成
    const curriculumUrl = url.replace(/\/exec$/, "/exec?page=curriculum");

    // 自動的に新しいタブで開く
    const html = `
      <div style="font-family: sans-serif; padding: 20px;">
        <h3>🎓 カリキュラム</h3>
        <p>カリキュラムを新しいタブで開いています...</p>
        <p><a href="${curriculumUrl}" target="_blank" style="color: #1a73e8; font-size: 14px;">${curriculumUrl}</a></p>
        <br>
        <p style="color: #666; font-size: 12px;">
          ※自動的に開かない場合は、上記のリンクをクリックしてください<br>
          ※段階的な学習プログラム
        </p>
      </div>
      <script>
        window.open('${curriculumUrl}', '_blank');
      </script>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(450)
      .setHeight(200);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "カリキュラム");
  } else {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "Webアプリがまだデプロイされていません。",
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

/**
 * 全データを削除
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ 警告",
    "本当にすべてのデータを削除しますか？\n\nこの操作は取り消せません。",
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // 各シートのデータをクリア
  Object.values(CONFIG.SHEETS).forEach((sheetName) => {
    db.clearData(sheetName);
  });

  ui.alert("✅ 完了", "すべてのデータを削除しました。", ui.ButtonSet.OK);
}

/**
 * 定期的に実行するトリガーを設定
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "updateDashboard") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // ダッシュボード更新トリガー（5分ごと）
  ScriptApp.newTrigger("updateDashboard")
    .timeBased()
    .everyMinutes(CONFIG.LIMITS.DASHBOARD_UPDATE_INTERVAL)
    .create();

  console.log("トリガーを設定しました");
}

/**
 * ダッシュボードを更新（トリガー用）
 */
function updateDashboard() {
  salesService.updateDashboard();
}

function handleError(error, context) {
  console.error(`Error in ${context}:`, error);

  // エラーログシートを取得または作成
  if (!db.hasSheet(CONFIG.SHEETS.ERROR_LOG)) {
    db.createSheet(CONFIG.SHEETS.ERROR_LOG, [
      "発生日時",
      "コンテキスト",
      "エラーメッセージ",
      "スタックトレース",
    ]);
  }

  // エラーログを記録
  db.addData(CONFIG.SHEETS.ERROR_LOG, [
    {
      発生日時: new Date(),
      コンテキスト: context,
      エラーメッセージ: error.toString(),
      スタックトレース: error.stack || "",
    },
  ]);

  // ユーザー向けエラーレスポンス
  return {
    success: false,
    message: "システムエラーが発生しました。管理者に連絡してください。",
    error: error.toString(),
  };
}
