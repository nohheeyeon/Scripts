////////////////////////////////////////////////////////////////////////////////////
// "PMS SW 패치 현황 및 검증서"를 최신화시키는 과정에서 사용

// 코드 실행 과정 :
// 1. "PMS SW 패치 현황 및 검증서" 접근
// 2. 첫 번째 시트(SW현황)과 정규표현식으로 찾은 최신 검증서 시트의 제품명을 비교
// 3. 동일한 제품명이 있다면, 최신 검증서 시트의 데이터를 SW현황 시트로 복사(복사 항목 : 발표일, 버전, 패치 파일명)
////////////////////////////////////////////////////////////////////////////////////
function main(workbook: ExcelScript.Workbook) {
  // "PMS SW 패치 현황 및 검증서" 불러오기
  let sheets = workbook.getWorksheets();

  // SW현황 시트를 가져오기
  let sw_summary_sheet = find_sheet_by_regex(sheets, /^SW현황$/);

  if (!sw_summary_sheet) {
    console.log("SW현황 시트를 찾을 수 없습니다.");
    return;
  }

  // 최신 검증서 시트를 가져오기
  let latest_validation_sheet = find_sheet_by_regex(sheets, /^\d{4}-\d{2}-\d{2} 검증서$/);

  if (!latest_validation_sheet) {
    console.log("최신검증서 시트를 찾을 수 없습니다.");
    return;
  }

  // 최신 검증서 시트에서 데이터 추출 및 업데이트
  let validationData = extract_validation_data(latest_validation_sheet);
  
  if (!validationData) {
    console.log("검증서 시트에서 데이터를 추출하는데 실패했습니다.");
    return;
  }
  update_sw_sheet(sw_summary_sheet, validationData);
}

function extract_validation_data(sheet: ExcelScript.Worksheet): {
  productNames: string[], 
  releaseDates: string[], 
  versions: string[], 
  patchFiles: string[]
} | null {

  let productCellAddress = find_cell_with_regex(sheet, /^제품명$/);
  let releaseDateCellAddress = find_cell_with_regex(sheet, /^발표일$/);
  let versionCellAddress = find_cell_with_regex(sheet, /^버전$/);
  let patchFileCellAddress = find_cell_with_regex(sheet, /^패치 파일$/);

  if (!productCellAddress || !releaseDateCellAddress || !versionCellAddress || !patchFileCellAddress) {
    return null;
  }

  let productNames: string[] = get_column_data(sheet, productCellAddress.row + 1, productCellAddress.column);
  let releaseDates: string[] = get_column_data(sheet, releaseDateCellAddress.row + 1, releaseDateCellAddress.column);
  let versions: string[] = get_column_data(sheet, versionCellAddress.row + 1, versionCellAddress.column);
  let patchFiles: string[] = get_column_data(sheet, patchFileCellAddress.row + 1, patchFileCellAddress.column);

  return {
    productNames,
    releaseDates,
    versions,
    patchFiles
  };
}

function update_sw_sheet(sheet: ExcelScript.Worksheet, validationData: {
  productNames: string[], 
  releaseDates: string[], 
  versions: string[], 
  patchFiles: string[]
}): void {

  let productCellAddress = find_cell_with_regex(sheet, /^제품명$/);
  let releaseDateCellAddress = find_cell_with_regex(sheet, /^발표일$/);
  let versionCellAddress = find_cell_with_regex(sheet, /^버전$/);
  let patchFileCellAddress = find_cell_with_regex(sheet, /^패치 파일$/);

  if (!productCellAddress || !releaseDateCellAddress || !versionCellAddress || !patchFileCellAddress) {
    return;
  }

  let swProductNames: string[] = get_column_data(sheet, productCellAddress.row + 1, productCellAddress.column+1);
  let commonProducts: string[] = find_common_values(swProductNames, validationData.productNames);

  let logMessages: string[] = [];

  for (let product of commonProducts) {
    let swProductIndex: number = swProductNames.indexOf(product);
    let validationProductIndex: number = validationData.productNames.indexOf(product);

    if (swProductIndex !== -1 && validationProductIndex !== -1) {
      let swReleaseDateCell = sheet.getCell(releaseDateCellAddress.row + swProductIndex, releaseDateCellAddress.column);
      let swVersionCell = sheet.getCell(versionCellAddress.row + swProductIndex, versionCellAddress.column);
      let swPatchFileCell = sheet.getCell(patchFileCellAddress.row + swProductIndex, patchFileCellAddress.column);
      swReleaseDateCell.setValue(validationData.releaseDates[validationProductIndex]);
      swVersionCell.setValue(validationData.versions[validationProductIndex]);
      swPatchFileCell.setValue(validationData.patchFiles[validationProductIndex]);

      logMessages.push(`[제품명: ${product}, 발표일: ${validationData.releaseDates[validationProductIndex]}, 패치 파일: ${validationData.patchFiles[validationProductIndex]}, 버전: ${validationData.versions[validationProductIndex]}]`);
    }
  }

  console.log("복사 완료된 제품:", logMessages);
}

function find_sheet_by_regex(sheets: ExcelScript.Worksheet[], regex: RegExp): ExcelScript.Worksheet | null {
  for (let sheet of sheets) {
    if (regex.test(sheet.getName())) {
      return sheet;
    }
  }
  return null;
}

function find_cell_with_regex(sheet: ExcelScript.Worksheet, regex: RegExp): { row: number, column: number } | null {
  let used_range = sheet.getUsedRange();
  let values = used_range.getValues();

  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (regex.test(String(values[row][col]))) {
        return { row: row + 1, column: col + 1 };
      }
    }
  }

  return null;
}

function get_column_data(sheet: ExcelScript.Worksheet, start_row: number, column: number): string[] {
  let column_data: string[] = [];
  let total_row_count = sheet.getUsedRange().getRowCount();
  let column_range = sheet.getRangeByIndexes(start_row - 1, column - 1, total_row_count - start_row + 1, 1);

  column_range.getValues().forEach(row => {
    let cell_value = row[0];
    if (cell_value !== null && cell_value !== undefined && cell_value !== "") {
      column_data.push(String(cell_value).trim());
    }
  });

  return column_data;
}

function find_common_values(array1: string[], array2: string[]): string[] {
  let common_values: string[] = [];

  array1.forEach(value => {
    if (array2.includes(value)) {
      common_values.push(value);
    }
  });

  return common_values;
}