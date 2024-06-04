import ExcelJS from "exceljs";
const workbook = new ExcelJS.Workbook();

import fs from "node:fs";

const FILE_NAME = "resource/r7_wordings.xlsx";
const SHEET = "Sheet1";
const SHEET_MAX_ROWS = 45;

interface IDBRow {
  table: string; // can be 'long' | 'short' | 'rich'
  id: number;
  fr: string;
  en: string;
  es: string;
  ar: string;
  nl: string;
  it: string;
  de: string;
}

// const fr_cell = ;
var problems = [];
var wording_problems = [];
var final_queries = [];
var ignored_lines = [];

workbook.xlsx
  .readFile(FILE_NAME)
  .then(() => {
    var worksheet = workbook.getWorksheet(SHEET);

    [...Array(SHEET_MAX_ROWS).keys()].map((key) => {
      const line = worksheet.getRow(key + 1);
      getDbRowFromExcelLine(line).map((result) => {
        fs.appendFile("out/default.json", JSON.stringify(result) + "\n", (err) => {
          if (err) {
            console.log("Writing Query Error : " + err);
          }
        });
      });
      final_queries.push();
    });

    console.log("------------WORDING ISSUES--------------");
    console.log(wording_problems);

    console.log("------------Other ISSUES--------------");
    console.log(problems);

    console.log("------------Ignored LINES --------------");
    console.log(ignored_lines);

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

function checkNewLine(input) {
  const regex = /\n/;
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
  const ES_COL = "F";
  const AR_COL = "G";
  const NL_COL = "H";
  const IT_COL = "I";
  const DE_COL = "J";

  const toBeCrawledTokens = ["long", "short", "rich"];

  const tokens = line.getCell(GUIDANCE_COL).toString().split(SEPARATOR);
  const tokens_length = tokens.length;

  const fr_list = line.getCell(FR_COL).toString().split(SEPARATOR);
  const en_list = line.getCell(EN_COL).toString().split(SEPARATOR);
  const es_list = line.getCell(ES_COL).toString().split(SEPARATOR);
  const ar_list = line.getCell(AR_COL).toString().split(SEPARATOR);
  const nl_list = line.getCell(NL_COL).toString().split(SEPARATOR);
  const it_list = line.getCell(IT_COL).toString().split(SEPARATOR);
  const de_list = line.getCell(DE_COL).toString().split(SEPARATOR);

  const languages = [
    { list: fr_list, col: FR_COL, language: "French" },
    { list: en_list, col: EN_COL, language: "English" },
    { list: es_list, col: ES_COL, language: "Spanish" },
    { list: ar_list, col: AR_COL, language: "Arabic" },
    { list: nl_list, col: NL_COL, language: "Neerlandais" },
    { list: it_list, col: IT_COL, language: "Italian" },
    { list: de_list, col: DE_COL, language: "Germany" },
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
  var json_results = [];

  if (lineProcessable) {
    filtered_tokens.forEach((token) => {
      token.token.split(TOKENS_SEP).forEach((subtoken) => {
        const table = subtoken.split(TABLE_IDS_SEP)[0];
        const id = subtoken.split(TABLE_IDS_SEP)[1];
        const lang_in = token.index;

        const fr_wording = fr_list[lang_in].trim();
        const en_wording = en_list[lang_in].trim();
        const es_wording = es_list[lang_in].trim();
        const ar_wording = ar_list[lang_in].trim();
        const nl_wording = nl_list[lang_in].trim();
        const it_wording = it_list[lang_in].trim();
        const de_wording = de_list[lang_in].trim();

        if (checkNewLine(fr_wording)) {
          problems.push(
            `FR in ${FR_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(en_wording)) {
          problems.push(
            `EN in ${EN_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(es_wording)) {
          problems.push(
            `ES in ${ES_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(ar_wording)) {
          problems.push(
            `AR in ${AR_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(nl_wording)) {
          problems.push(
            `NL in ${NL_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(it_wording)) {
          problems.push(
            `IT in ${IT_COL}${line.number} contains a new line in the middle`
          );
        }
        if (checkNewLine(de_wording)) {
          problems.push(
            `DE in ${DE_COL}${line.number} contains a new line in the middle`
          );
        }
        if (toBeCrawledTokens.includes(table)) {
          json_results.push({
            fr: fr_wording,
            en: en_wording,
            es: es_wording,
            ar: ar_wording,
            nl: nl_wording,
            it: it_wording,
            de: de_wording,
          });
        }
      });
    });
  } else {
    ignored_lines.push("Line " + line.number);
  }
  return json_results;
};

// IGNORE THE RICH TEXT BECAUSE OF HTML TAGS => DONE
// TRIM THE WORDINGS TO DELETE SPACES AND \n from the first line => DONE
//  After the first trim if \n still exists in the wording and it's short || long
// Error
// ISO DATABASE to facilitate insertion
// Add a sheet for the to delete wordings
// Map to the table => DONE

const tablesMapping = {
  short: "components_language_language_short_texts",
  long: "components_language_language_long_texts",
  rich: "components_language_language_rich_texts",
};

const getSqlQueriesFromDbRows = (db_rows: IDBRow[]) => {
  return db_rows.map((row) => {
    const { table, id, fr, en, es, ar, nl, it, de } = row;
    return `UPDATE ${tablesMapping[table]} SET fr = \'${fr}\', en = \'${en}\', es = \'${es}\', ar = \'${ar}\', nl = \'${nl}\', it = \'${it}\', de = \'${de}\' WHERE id = ${id};`;
  });
};
