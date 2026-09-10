/** @OfficeScript */
// i never thought id stoop so low as to be a "prompt engineer"
// but i hate to admit, this saved a LOT of time and works fine
// if i ever want to update it i will NEED ai though, 
// which is a worthy drawback ig for saving HOURS
// maybe ai is the future :pensive:

const INVOICE_OUTPUT_ROW = 3;
const DROPDOWN_COLUMN = 4; // E

const TARGET_HEADERS = [
  "NAME",
  "SHOP_NAME",
  "SHOP_QUANTITY",
  "PRICE CHANGE",
  "PRICE PER UNIT",
  "SHOP_PRICE",
  "MARKUP"
];

function main(workbook: ExcelScript.Workbook) {

  const inputSheet = workbook.getWorksheet("INPUT");
  const shopItemsSheet = workbook.getWorksheet("SHOP_ITEMS");

  if (!inputSheet) {
    throw new Error("INPUT sheet not found.");
  }

  if (!shopItemsSheet) {
    throw new Error("SHOP_ITEMS sheet not found.");
  }

  /*
   * Get existing company sheets.
   */
  const worksheetNames = workbook
    .getWorksheets()
    .map(sheet => sheet.getName())
    .filter(name =>
      name !== "INPUT" &&
      name !== "SHOP_ITEMS"
    );

  /*
   * Company selector.
   */
  const companyCell = inputSheet.getRange("B1");

  if (worksheetNames.length > 0) {
    companyCell.getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: worksheetNames.join(",")
      }
    });
  }

  const company =
    String(companyCell.getValue()).trim();

  const customName =
    String(inputSheet.getRange("B2").getValue()).trim();

  /*
   * Decide target sheet.
   */
  let targetSheetName = "";

  if (
    company &&
    company !== "Select a name"
  ) {
    targetSheetName = company;

  } else if (customName) {
    targetSheetName = customName;

  } else {
    throw new Error(
      "Please select a company name or enter a new company name in B2."
    );
  }

  /*
   * Create target sheet if necessary.
   */
  let targetSheet =
    workbook.getWorksheet(targetSheetName);

  if (!targetSheet) {

    targetSheet =
      workbook.addWorksheet(targetSheetName);

    targetSheet
      .getRange("A1:G1")
      .setValues([TARGET_HEADERS]);

    formatHeaders(targetSheet);

  } else {

    targetSheet
      .getRange("A1:G1")
      .setValues([TARGET_HEADERS]);

    formatHeaders(targetSheet);
  }

  /*
   * Read SHOP_ITEMS.
   *
   * A = Name
   * B = Variant
   * C = Price
   * D = Average Weight
   */
  const shopRange =
    shopItemsSheet.getUsedRange();

  if (!shopRange) {
    throw new Error("SHOP_ITEMS is empty.");
  }

  const shopValues =
    shopRange.getValues();

  const shopItems: ShopItem[] = [];

  for (
    let i = 0;
    i < shopValues.length;
    i++
  ) {

    const name =
      String(shopValues[i][0] ?? "").trim();

    const variant =
      String(shopValues[i][1] ?? "").trim();

    if (!name && !variant) {
      continue;
    }

    const fullName =
      `${name} - ${variant}`.trim();

    shopItems.push({
      name: fullName,
      price: toNumber(shopValues[i][2]),
      averageWeight: toNumber(shopValues[i][3])
    });
  }

  /*
   * Read invoice.
   *
   * B = PRODUCT_NAME
   * C = QUANTITY
   * D = TOTAL_PRICE
   * E = SHOP_NAME
   * F = SHOP_QUANTITY
   * G = AVERAGE_WEIGHT
   */
  const invoice =
    getSalesInvoiceOutput(inputSheet);

  if (invoice.length === 0) {
    throw new Error(
      "No invoice products were found in INPUT."
    );
  }

  /*
   * Get current target values.
   */
  let targetValues =
    getTargetValues(targetSheet);

  /*
   * Process every invoice product.
   */
  for (
    let i = 0;
    i < invoice.length;
    i++
  ) {

    const inputRow =
      INVOICE_OUTPUT_ROW + i;

    const productName =
      invoice[i].name;

    const invoiceQuantity =
      invoice[i].quantity;

    const totalPrice =
      invoice[i].totalPrice;

    /*
     * INPUT columns E, F and G.
     */
    const shopNameCell =
      inputSheet.getCell(
        inputRow,
        DROPDOWN_COLUMN
      );

    const shopQuantityCell =
      inputSheet.getCell(
        inputRow,
        DROPDOWN_COLUMN + 1
      );

    const averageWeightCell =
      inputSheet.getCell(
        inputRow,
        DROPDOWN_COLUMN + 2
      );

    /*
     * Check whether this product already
     * exists on the company sheet.
     */
    const existingRow =
      findTargetProductRow(
        targetValues,
        productName
      );

    /*
     * If it already exists, restore the
     * shop selection and quantity.
     */
    if (existingRow >= 0) {

      const existingShopName =
        String(
          targetValues[existingRow][1] ?? ""
        );

      const existingShopQuantity =
        toNumber(
          targetValues[existingRow][2]
        );

      if (
        !shopNameCell.getValue() &&
        existingShopName
      ) {
        shopNameCell.setValue(
          existingShopName
        );
      }

      if (
        !shopQuantityCell.getValue() &&
        existingShopQuantity
      ) {
        shopQuantityCell.setValue(
          existingShopQuantity
        );
      }
    }

    /*
     * If no shop product is selected,
     * perform fuzzy matching.
     */
    if (!shopNameCell.getValue()) {

      const matches =
        shopItems
          .map(shopItem => ({
            item: shopItem,
            score: fuzzyScore(
              productName,
              shopItem.name
            )
          }))
          .sort(
            (a, b) =>
              b.score - a.score
          )
          .slice(0, 20);

      if (matches.length === 0) {
        throw new Error(
          `No SHOP_ITEMS found for "${productName}".`
        );
      }

      const options =
        matches.map(match =>
          match.item.name
            .replace(/,/g, " - ")
        );

      shopNameCell
        .getFormat()
        .getFill()
        .setColor("yellow");

      shopNameCell
        .getDataValidation()
        .setRule({
          list: {
            inCellDropDown: true,
            source: options.join(",")
          }
        });

      /*
       * Automatically select best match.
       */
      shopNameCell.setValue(
        matches[0].item.name
      );
    }

    /*
     * Find selected shop item.
     */
    const selectedShopName =
      String(
        shopNameCell.getValue()
      ).trim();

    const selectedShopItem =
      findShopItem(
        shopItems,
        selectedShopName
      );

    /*
     * Fill average weight.
     */
    if (selectedShopItem) {

      averageWeightCell.setValue(
        selectedShopItem.averageWeight
      );

    } else {

      averageWeightCell.setFormula(
        `=IFERROR(XLOOKUP(E${inputRow + 1},SHOP_ITEMS!A:A&" - "&SHOP_ITEMS!B:B,SHOP_ITEMS!D:D),"")`
      );
    }

    /*
     * SHOP_QUANTITY is required.
     */
    if (!shopQuantityCell.getValue()) {

      shopQuantityCell
        .getFormat()
        .getFill()
        .setColor("yellow");

      continue;
    }

    /*
     * Clear highlighting.
     */
    shopNameCell
      .getFormat()
      .getFill()
      .clear();

    shopQuantityCell
      .getFormat()
      .getFill()
      .clear();

    const shopQuantity =
      toNumber(
        shopQuantityCell.getValue()
      );

    /*
     * Validate numbers.
     */
    if (
      invoiceQuantity <= 0 ||
      totalPrice < 0 ||
      shopQuantity <= 0
    ) {
      throw new Error(
        `Invalid quantity or price for "${productName}".`
      );
    }

    /*
     * ---------------------------------------------
     * PRICE PER SUPPLIER UNIT
     *
     * Example:
     *
     * Soursop:
     * 7.5 kg
     * £60 total
     *
     * £60 / 7.5 = £8 per kg
     * ---------------------------------------------
     */
    const pricePerSupplierUnit =
      totalPrice / invoiceQuantity;

    /*
     * ---------------------------------------------
     * PRICE PER SHOP UNIT
     *
     * Example:
     *
     * £8 per kg
     * 2 × 500g shop units per kg
     *
     * £8 / 2 = £4 per 500g
     * ---------------------------------------------
     */
    const pricePerUnit =
      pricePerSupplierUnit /
      shopQuantity;

    /*
     * Shop selling price.
     */
    const shopPrice =
      selectedShopItem
        ? selectedShopItem.price
        : 0;

    let profitMarkup = 0;

    if (pricePerUnit > 0) {
      profitMarkup =
        (
          shopPrice -
          pricePerUnit
        ) /
        pricePerUnit;
    }

    /*
     * Update target sheet.
     */
    writeTargetProduct(
      workbook,
      targetSheet,
      productName,
      selectedShopName,
      shopQuantity,
      pricePerUnit,
      shopPrice,
      profitMarkup,
      existingRow
    );

    /*
     * Refresh target values because the
     * sheet may have changed.
     */
    targetValues =
      getTargetValues(targetSheet);
  }

  /*
   * Final formatting.
   */
  formatTargetSheet(targetSheet);
}


/* =====================================================
   TYPES
   ===================================================== */

interface ShopItem {
  name: string;
  price: number;
  averageWeight: number;
}

interface InvoiceRow {
  name: string;
  quantity: number;
  totalPrice: number;
}


/* =====================================================
   INPUT
   ===================================================== */

function getSalesInvoiceOutput(
  inputSheet: ExcelScript.Worksheet
): InvoiceRow[] {

  const table: InvoiceRow[] = [];

  let row =
    INVOICE_OUTPUT_ROW;

  while (true) {

    const name =
      inputSheet
        .getCell(row, 1)
        .getValue();

    if (!name) {
      break;
    }

    const quantity =
      toNumber(
        inputSheet
          .getCell(row, 2)
          .getValue()
      );

    const totalPrice =
      toNumber(
        inputSheet
          .getCell(row, 3)
          .getValue()
      );

    table.push({
      name: String(name).trim(),
      quantity: quantity,
      totalPrice: totalPrice
    });

    row++;
  }

  return table;
}


/* =====================================================
   TARGET VALUES
   ===================================================== */

function getTargetValues(
  sheet: ExcelScript.Worksheet
): (string | number | boolean)[][] {

  const usedRange =
    sheet.getUsedRange();

  if (!usedRange) {
    return [];
  }

  return usedRange.getValues();
}


/* =====================================================
   FIND EXISTING PRODUCT
   ===================================================== */

function findTargetProductRow(
  targetValues:
    (string | number | boolean)[][],
  productName: string
): number {

  const wanted =
    normaliseTargetName(productName);

  /*
   * Row 0 is the header.
   */
  for (
    let row = 1;
    row < targetValues.length;
    row++
  ) {

    const existingName =
      String(
        targetValues[row][0] ?? ""
      );

    if (
      normaliseTargetName(existingName) ===
      wanted
    ) {
      return row;
    }
  }

  return -1;
}


function normaliseTargetName(
  value: string
): string {

  return value
    .toLowerCase()
    .replace(
      /[^\p{L}\p{N}\s]/gu,
      " "
    )
    .replace(
      /\s+/g,
      " "
    )
    .trim();
}


/* =====================================================
   WRITE PRODUCT
   ===================================================== */

function writeTargetProduct(
  workbook: ExcelScript.Workbook,
  targetSheet: ExcelScript.Worksheet,
  productName: string,
  shopName: string,
  shopQuantity: number,
  newPricePerUnit: number,
  shopPrice: number,
  profitMarkup: number,
  existingRow: number
) {

  let row: number;

  /*
   * Existing product.
   */
  if (existingRow >= 0) {

    row = existingRow;

  } else {

    /*
     * New product goes after the existing data.
     */
    const usedRange =
      targetSheet.getUsedRange();

    if (!usedRange) {
      row = 1;
    } else {
      row =
        usedRange.getRowCount();
    }
  }

  /*
   * ---------------------------------------------
   * Read previous PRICE PER UNIT.
   *
   * E = PRICE PER UNIT
   * ---------------------------------------------
   */
  let previousPricePerUnit = 0;

  if (existingRow >= 0) {

    previousPricePerUnit =
      toNumber(
        targetSheet
          .getCell(row, 4)
          .getValue()
      );
  }

  /*
   * ---------------------------------------------
   * PRICE CHANGE
   * ---------------------------------------------
   */
  let priceChange = 0;

  if (existingRow >= 0) {

    priceChange =
      newPricePerUnit -
      previousPricePerUnit;
  }

  /*
   * A - NAME
   */
  targetSheet
    .getCell(row, 0)
    .setValue(productName);

  /*
   * B - SHOP_NAME
   */
  targetSheet
    .getCell(row, 1)
    .setValue(shopName);

  /*
   * C - SHOP_QUANTITY
   */
  targetSheet
    .getCell(row, 2)
    .setValue(shopQuantity);

  /*
   * D - PRICE CHANGE
   */
  targetSheet
    .getCell(row, 3)
    .setValue(priceChange);

  /*
   * E - PRICE PER UNIT
   */
  targetSheet
    .getCell(row, 4)
    .setValue(newPricePerUnit);

  /*
   * F - SHOP PRICE
   */
  targetSheet
    .getCell(row, 5)
    .setValue(shopPrice);

  targetSheet
    .getCell(row, 6)
    .setValue(profitMarkup);

  /*
   * Currency formatting.
   */
  targetSheet
    .getRange(
      `D${row + 1}:F${row + 1}`
    )
    .setNumberFormatLocal("£0.00");

  /*
   * Percentage formatting.
   */
  targetSheet
    .getCell(row, 6)
    .setNumberFormatLocal("0.0%");

  /*
   * ---------------------------------------------
   * PRICE HISTORY COMMENT
   * ---------------------------------------------
   *
   * The comment is attached to PRICE PER UNIT.
   */
  updatePriceHistoryComment(
    workbook,
    targetSheet,
    row,
    newPricePerUnit,
    existingRow >= 0,
    previousPricePerUnit
  );

  colourProfitMarkup(
    targetSheet.getCell(row, 6),
    shopPrice,
    newPricePerUnit
  );
}


/* =====================================================
   PRICE HISTORY
   ===================================================== */

function updatePriceHistoryComment(
  workbook: ExcelScript.Workbook,
  targetSheet: ExcelScript.Worksheet,
  row: number,
  newPrice: number,
  existingProduct: boolean,
  previousPrice: number
) {

  const priceCell =
    targetSheet.getCell(row, 4);

  /*
   * Get existing comment if one exists.
   */
  let existingComment:
    ExcelScript.Comment | undefined;

  try {

    existingComment =
      workbook.getCommentByCell(
        priceCell
      );

  } catch {
    existingComment = undefined;
  }

  /*
   * Get existing history.
   */
  let history: string[] = [];

  if (existingComment) {

    const content =
      existingComment
        .getContent();

    history =
      parsePriceHistory(content);

    /*
     * Delete the old comment before recreating it.
     */
    existingComment.delete();
  }

  const today =
    formatDate(new Date());

  /*
   * Only add a new history entry when
   * the price has actually changed.
   *
   * For a brand new product, add the
   * initial price.
   */
  let shouldAddEntry = false;

  if (!existingProduct) {

    shouldAddEntry = true;

  } else if (
    Math.abs(
      newPrice -
      previousPrice
    ) > 0.000001
  ) {

    shouldAddEntry = true;
  }

  /*
   * Avoid adding the same price/date twice.
   */
  if (shouldAddEntry) {

    const newEntry =
      `${today} — £${newPrice.toFixed(2)}`;

    const alreadyExists =
      history.some(
        entry =>
          entry === newEntry
      );

    if (!alreadyExists) {
      history.push(newEntry);
    }
  }

  /*
   * If there is history, recreate the comment.
   */
  if (history.length > 0) {

    const commentText =
      "Price history\n\n" +
      history.join("\n");

    workbook.addComment(
      priceCell,
      commentText,
      ExcelScript.ContentType.plain
    );
  }
}


function parsePriceHistory(
  content: string
): string[] {

  const lines =
    content
      .split("\n")
      .map(line => line.trim())
      .filter(line =>
        line.length > 0
      );

  /*
   * Remove the title.
   */
  return lines.filter(
    line =>
      line.toLowerCase() !==
      "price history"
  );
}


/* =====================================================
   DATE
   ===================================================== */

function formatDate(
  date: Date
): string {

  const day =
    String(
      date.getDate()
    ).padStart(2, "0");

  const month =
    String(
      date.getMonth() + 1
    ).padStart(2, "0");

  const year =
    date.getFullYear();

  return `${day}/${month}/${year}`;
}


/* =====================================================
   SHOP ITEM LOOKUP
   ===================================================== */

function findShopItem(
  shopItems: ShopItem[],
  name: string
): ShopItem | undefined {

  const wanted =
    normaliseShopName(name);

  /*
   * Exact normalised match.
   */
  for (const item of shopItems) {

    if (
      normaliseShopName(item.name) ===
      wanted
    ) {
      return item;
    }
  }

  /*
   * Handle comma replacement from
   * Excel's dropdown.
   */
  const wantedWithoutComma =
    normaliseShopName(
      name.replace(
        /,/g,
        " - "
      )
    );

  for (const item of shopItems) {

    if (
      normaliseShopName(item.name) ===
      wantedWithoutComma
    ) {
      return item;
    }
  }

  return undefined;
}


function normaliseShopName(
  value: string
): string {

  return value
    .toLowerCase()
    .replace(
      /[^\p{L}\p{N}\s]/gu,
      " "
    )
    .replace(
      /\s+/g,
      " "
    )
    .trim();
}

function colourProfitMarkup(
  cell: ExcelScript.Range,
  shopPrice: number,
  costPerUnit: number
) {

  cell
    .getFormat()
    .getFill()
    .clear();

  if (
    shopPrice <= 0 ||
    costPerUnit <= 0
  ) {
    return;
  }

  /*
   * Green:
   * Selling price is at least double cost.
   */
  if (
    shopPrice >=
    costPerUnit * 2
  ) {

    cell
      .getFormat()
      .getFill()
      .setColor("#c6efce");

    /*
     * Red:
     * Breaking even or losing money.
     */
  } else if (
    shopPrice <= costPerUnit
  ) {

    cell
      .getFormat()
      .getFill()
      .setColor("#ffc7ce");

    /*
     * Yellow:
     * Profitable, but less than 2x cost.
     */
  } else {

    cell
      .getFormat()
      .getFill()
      .setColor("#ffeb9c");
  }
}


/* =====================================================
   FORMATTING
   ===================================================== */

function formatHeaders(
  sheet: ExcelScript.Worksheet
) {

  const header =
    sheet.getRange("A1:G1");

  header
    .getFormat()
    .getFont()
    .setBold(true);

  header
    .getFormat()
    .getFill()
    .setColor("#d9e1f2");
}


function formatTargetSheet(
  sheet: ExcelScript.Worksheet
) {

  const usedRange =
    sheet.getUsedRange();

  if (!usedRange) {
    return;
  }

  formatHeaders(sheet);

  const rowCount =
    usedRange.getRowCount();

  if (rowCount > 1) {

    /*
     * D = Price Change
     * E = Price Per Unit
     * F = Shop Price
     */
    sheet
      .getRange(
        `D2:F${rowCount}`
      )
      .setNumberFormatLocal("£0.00");

    sheet
      .getRange(
        `G2:G${rowCount}`
      )
      .setNumberFormatLocal("0.0%");
  }

  /*
   * Automatically size columns.
   */
  sheet
    .getRange("A:G")
    .getFormat()
    .autofitColumns();
}


/* =====================================================
   FUZZY MATCHING
   ===================================================== */

function fuzzyScore(
  input: string,
  candidate: string
): number {

  const inputWords =
    normalise(input);

  const candidateWords =
    normalise(candidate);

  if (
    inputWords.length === 0 ||
    candidateWords.length === 0
  ) {
    return -Infinity;
  }

  let exactMatches = 0;
  let fuzzyMatches = 0;
  let fuzzyScore = 0;

  const used: boolean[] =
    candidateWords.map(
      () => false
    );

  /*
   * Exact matches.
   */
  for (const inputWord of inputWords) {

    for (
      let i = 0;
      i < candidateWords.length;
      i++
    ) {

      if (used[i]) {
        continue;
      }

      if (
        inputWord ===
        candidateWords[i]
      ) {

        used[i] = true;
        exactMatches++;

        break;
      }
    }
  }

  /*
   * Fuzzy matches.
   */
  for (const inputWord of inputWords) {

    let alreadyMatched =
      false;

    for (
      let i = 0;
      i < candidateWords.length;
      i++
    ) {

      if (
        candidateWords[i] === inputWord &&
        used[i]
      ) {

        alreadyMatched = true;
        break;
      }
    }

    if (alreadyMatched) {
      continue;
    }

    let best = 0;
    let bestIndex = -1;

    for (
      let i = 0;
      i < candidateWords.length;
      i++
    ) {

      if (used[i]) {
        continue;
      }

      const similarity =
        wordSimilarity(
          inputWord,
          candidateWords[i]
        );

      if (similarity > best) {

        best = similarity;
        bestIndex = i;
      }
    }

    if (
      best >= 0.75 &&
      bestIndex >= 0
    ) {

      used[bestIndex] = true;
      fuzzyMatches++;
      fuzzyScore += best;
    }
  }

  let score = 0;

  score +=
    exactMatches * 10000;

  score +=
    fuzzyMatches * 100;

  score +=
    fuzzyScore * 10;

  const extraWords =
    candidateWords.length -
    exactMatches -
    fuzzyMatches;

  score -=
    extraWords * 2000;

  score -=
    candidateWords.length * 10;

  return score;
}


/* =====================================================
   NORMALISE FUZZY SEARCH
   ===================================================== */

function normalise(
  value: string
): string[] {

  const words =
    value
      .toLowerCase()

      /*
       * Ignore weights when finding
       * the best shop product.
       */
      .replace(
        /\d+(?:\.\d+)?\s*(?:kg|kgs|g|gm|gms|gram|grams)\b/g,
        " "
      )

      /*
       * Ignore standalone numbers.
       */
      .replace(
        /\d+(?:\.\d+)?/g,
        " "
      )

      .replace(
        /[^\p{L}\s]/gu,
        " "
      )

      .split(/\s+/)
      .filter(
        word =>
          word.length > 0
      );

  const uniqueWords: string[] = [];

  for (const word of words) {

    if (
      uniqueWords.indexOf(word) === -1
    ) {

      uniqueWords.push(word);
    }
  }

  return uniqueWords;
}


/* =====================================================
   WORD SIMILARITY
   ===================================================== */

function wordSimilarity(
  a: string,
  b: string
): number {

  if (a === b) {
    return 1;
  }

  const distance =
    levenshtein(a, b);

  const length =
    Math.max(
      a.length,
      b.length
    );

  if (length === 0) {
    return 0;
  }

  return (
    1 -
    distance / length
  );
}


/* =====================================================
   LEVENSHTEIN
   ===================================================== */

function levenshtein(
  a: string,
  b: string
): number {

  const matrix: number[][] = [];

  for (
    let i = 0;
    i <= b.length;
    i++
  ) {

    matrix[i] = [i];
  }

  for (
    let j = 0;
    j <= a.length;
    j++
  ) {

    matrix[0][j] = j;
  }

  for (
    let i = 1;
    i <= b.length;
    i++
  ) {

    for (
      let j = 1;
      j <= a.length;
      j++
    ) {

      if (
        b.charAt(i - 1) ===
        a.charAt(j - 1)
      ) {

        matrix[i][j] =
          matrix[i - 1][j - 1];

      } else {

        matrix[i][j] =
          Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
      }
    }
  }

  return matrix[b.length][a.length];
}


/* =====================================================
   NUMBER CONVERSION
   ===================================================== */

function toNumber(
  value: string | number | boolean
): number {

  if (
    typeof value === "number"
  ) {
    return value;
  }

  const parsed =
    Number(
      String(value)
        .replace(/£/g, "")
        .replace(/,/g, "")
        .trim()
    );

  return isNaN(parsed)
    ? 0
    : parsed;
}