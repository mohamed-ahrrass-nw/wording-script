import ExcelJS from "exceljs";
const workbook = new ExcelJS.Workbook();

const FILE_NAME = "resource/r7_wordings.xlsx";
const SHEET = "Sheet1";
const SHEET_MAX_ROWS = 44;

interface IDBRow {
  table: string; // can be 'long' | 'short' | 'rich'
  id: number;
  fr: string;
  en: string;
  ar: string;
  it: string;
}

// const fr_cell = ;
var problems = [];
var wording_problems = [];
var final_queries = [];

workbook.xlsx
  .readFile(FILE_NAME)
  .then(() => {
    var worksheet = workbook.getWorksheet(SHEET);

    [...Array(SHEET_MAX_ROWS).keys()].map((key) => {
      const line = worksheet.getRow(key + 1);
      getDbRowFromExcelLine(line);
    });

    console.log("------------WORDING ISSUES--------------");
    console.log(wording_problems);

    console.log("----------Results------------");
    console.log(final_queries);
  })
  .catch((error) => {
    console.error("Error reading file:", error);
  });

function checkListLength(tokens_length, list, col, line, language) {
  if (tokens_length !== list.length) {
    wording_problems.push(
      `${language} wordings in ${col}${line.number} not separated correctly | Lengths : Tokens = ${tokens_length} Wording list = ${list.length}`
    );
    return false;
  }
  return true;
}

function validateToken(input) {
  const regex = /^(\w+#\d+)(,\w+#\d+)*$/;
  return regex.test(input);
}

const getDbRowFromExcelLine = (line: ExcelJS.Row) => {
  console.log(`Processing line : ${line.number} - ${line.getCell("A")}`);

  var lineProcessable = true;
  const SEPARATOR = "\\\\";
  const TABLE_IDS_SEP = "#";
  const TOKENS_SEP = ",";

  const GUIDANCE_COL = "C";

  const FR_COL = "D";
  const EN_COL = "E";
  const AR_COL = "G";
  const IT_COL = "I";

  const tokens = line.getCell(GUIDANCE_COL).toString().split(SEPARATOR);
  const tokens_length = tokens.length;

  const fr_list = line.getCell(FR_COL).toString().split(SEPARATOR);
  const en_list = line.getCell(EN_COL).toString().split(SEPARATOR);
  const ar_list = line.getCell(AR_COL).toString().split(SEPARATOR);
  const it_list = line.getCell(IT_COL).toString().split(SEPARATOR);

  const languages = [
    { list: fr_list, col: FR_COL, language: "French" },
    { list: en_list, col: EN_COL, language: "English" },
    { list: ar_list, col: AR_COL, language: "Arabic" },
    { list: it_list, col: IT_COL, language: "Italian" },
  ];

  // Check for errors and issues
  // => Langs wordings length checks
  languages.forEach(({ list, col, language }) => {
    if (!checkListLength(tokens_length, list, col, line, language)) {
      lineProcessable = false;
    }
  });

  // => Tokens check
  var filtered_tokens: { token: string; index: number }[] = [];

  tokens.forEach((token, index) => {
    const tokenF = token.replace(/\n/g, "");
    // console.log(`Processing token : ${tokenF}`);
    if (!validateToken(tokenF)) {
      problems.push(
        `token ${index + 1} in ${GUIDANCE_COL}${
          line.number
        } is not respecting the convention`
      );
    } else {
      filtered_tokens.push({ token: tokenF, index });
    }
  });

  console.log(`Filtered tokens : ${JSON.stringify(filtered_tokens)}`);

  // if there is no problem with the wordings
  var db_rows = [];
  if (lineProcessable) {
    filtered_tokens.forEach((token) => {
      token.token.split(TOKENS_SEP).forEach((subtoken) => {
        const table = subtoken.split(TABLE_IDS_SEP)[0];
        const id = subtoken.split(TABLE_IDS_SEP)[1];
        const lang_in = token.index;
        db_rows.push({
          table,
          id,
          fr: fr_list[lang_in],
          en: en_list[lang_in],
          ar: ar_list[lang_in],
          it: it_list[lang_in],
        });
      });
    });
    
  }
  console.log(getSqlQueriesFromDbRows(db_rows));
};

// IGNORE THE RICH TEXT BECAUSE OF HTML TAGS
// TRIM THE WORDINGS TO DELETE SPACES AND \n from the first line
//  After the first trim if \n still exists in the wording and it's short || long 
      // Error 
      // 

const getSqlQueriesFromDbRows = (db_rows: IDBRow[]) => {
  return db_rows.map((row) => {
    const { table, id, fr, en, ar, it } = row;
    return `UPDATE ${table} SET fr = '${fr}', en = '${en}', ar = '${ar}', it = '${it}' WHERE id = ${id};`;
  });
};
