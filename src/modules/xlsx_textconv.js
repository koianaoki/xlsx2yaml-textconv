const xlsx = require("../modules/xlsx_ext.js");

function xlsx_to_yaml(file, field_row, select_sheet_name) {
  let output_text = "";
  let book = xlsx.readFile(file);
  book.SheetNames.forEach(function( sheet_name ) {
    if (select_sheet_name) {
      regexp = new RegExp(select_sheet_name);
      if (!sheet_name.match(regexp)) return;
    }

    let sheet = book.Sheets[sheet_name];
    let yaml;
    if (field_row !== false) {
      yaml = xlsx.utils.sheet_to_yaml(sheet, {range: xlsx.utils.table_range(sheet, field_row), raw: true});
    } else {
      yaml = xlsx.utils.sheet_to_yaml(sheet, {header: "A"});
    }

    let blank_size = sheet_name.length;
    yaml = yaml.replace(/\n-/g, "\n\n-")
               .replace(/\n\s\s/g, `\n${" ".repeat(blank_size + 1)}`)
               .replace(/\n-\s/g, `\n${sheet_name} `)
               .replace(/^-\s/g, `\n${sheet_name} `);

    output_text += yaml
  });
  console.log(output_text)
}

module.exports.xlsx_to_yaml = xlsx_to_yaml
