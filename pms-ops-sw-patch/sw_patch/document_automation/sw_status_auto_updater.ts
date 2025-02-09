function main(workbook: ExcelScript.Workbook) {
  // "PSW 패치 현황 및 검증서" 불러오기
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

  // "제품"/"제품명"이라는 값을 가진 셀을 정규표현식을 사용하여 찾기
  let sw_product_cell_address = find_cell_with_regex(sw_summary_sheet, /^제품$/);
  let validation_product_cell_address = find_cell_with_regex(latest_validation_sheet, /^제품명$/);

  // "발표일" 셀 찾기
  let sw_release_date_cell_address = find_cell_with_regex(sw_summary_sheet, /^발표일$/);
  let validation_release_date_cell_address = find_cell_with_regex(latest_validation_sheet, /^발표일$/);

  // "버전" 셀 찾기
  let sw_version_cell_address = find_cell_with_regex(sw_summary_sheet, /^버전\(G\)$/);
  let validation_version_cell_address = find_cell_with_regex(latest_validation_sheet, /^버전$/);

  // "패치 파일(H)"/"패치 파일" 셀 찾기
  let sw_patch_file_cell_address = find_cell_with_regex(sw_summary_sheet, /^패치 파일\(H\)$/);
  let validation_patch_file_cell_address = find_cell_with_regex(latest_validation_sheet, /^패치 파일$/);

  if (sw_product_cell_address && validation_product_cell_address && sw_release_date_cell_address && validation_release_date_cell_address && sw_version_cell_address && validation_version_cell_address && sw_patch_file_cell_address && validation_patch_file_cell_address) {

    // "제품" 및 "제품명" 셀 하위의 모든 열 데이터를 추출
    let sw_product_names = get_column_data(sw_summary_sheet, sw_product_cell_address.row + 1, sw_product_cell_address.column + 1);
    let validation_product_names = get_column_data(latest_validation_sheet, validation_product_cell_address.row + 1, validation_product_cell_address.column);

    // "발표일" 셀 하위의 모든 열 데이터를 추출
    let validation_release_dates = get_column_data(latest_validation_sheet, validation_release_date_cell_address.row + 1, validation_release_date_cell_address.column);
    
    // "버전" 셀 하위의 모든 열 데이터를 추출
    let validation_versions = get_column_data(latest_validation_sheet, validation_version_cell_address.row + 1, validation_version_cell_address.column);
   
    // "패치 파일(H)" 및 "패치 파일" 셀 하위의 모든 열 데이터를 추출
    let validation_patch_file = get_column_data(latest_validation_sheet, validation_patch_file_cell_address.row + 1, validation_patch_file_cell_address.column);

    // 공통된 제품명 찾기
    let common_values = find_common_values(sw_product_names, validation_product_names);
    console.log("공통된 제품명:", common_values);

    // 로그를 저장할 배열 생성
    let logMessages: string[] = [];

    // 공통된 제품명과 같은 행에 있는 발표일 데이터를 SW현황 시트에 복사
    for (let product of common_values) {
      let product_index_in_sw = sw_product_names.indexOf(product);
      let product_index_in_validation = validation_product_names.indexOf(product);

      if (product_index_in_validation !== -1 && product_index_in_sw !== -1) {

        let validation_release_date = validation_release_dates[product_index_in_validation];
        let sw_release_date_cell = sw_summary_sheet.getCell(sw_release_date_cell_address.row + product_index_in_sw, sw_release_date_cell_address.column);
        let validation_patch_files = validation_patch_file[product_index_in_validation];
        let sw_patch_file_cell = sw_summary_sheet.getCell(sw_patch_file_cell_address.row + product_index_in_sw, sw_patch_file_cell_address.column);
        let validation_product_version = validation_versions[product_index_in_validation];
        let sw_product_version_cell = sw_summary_sheet.getCell(sw_version_cell_address.row + product_index_in_sw, sw_version_cell_address.column);
        
        // 발표일 데이터를 SW현황 시트에 복사
        sw_release_date_cell.setValue(validation_release_date);
        sw_patch_file_cell.setValue(validation_patch_files);
        sw_product_version_cell.setValue(validation_product_version);

        // 로그 메시지를 배열에 추가
        logMessages.push(`[제품명: ${product}, 발표일: ${validation_release_date}, 패치 파일: ${validation_patch_files}, 버전: ${validation_product_version}]`);
      }
    }

    // 배열 형태로 복사 완료 로그 메시지 출력
    console.log("복사 완료된 제품:", logMessages);

  } else {
    console.log("SW현황 시트 또는 최신 검증서 시트에서 '제품' 셀을 찾을 수 없습니다.");
  }

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

// "제품" 및 "제품명"에서 같은 값을 찾는 함수
function find_common_values(array1: string[], array2: string[]): string[] {
  let common_values: string[] = [];

  array1.forEach(value => {
    if (array2.includes(value)) {
      common_values.push(value);
    }
  });

  return common_values;
}
