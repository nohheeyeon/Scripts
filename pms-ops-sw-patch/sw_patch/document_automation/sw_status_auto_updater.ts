function main(workbook: ExcelScript.Workbook) {
    // "SW 패치 현황 및 검증서" 불러오기
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
    console.log("최신검증서 시트를 찾을 수 없습니다.")
      return;
    }
  
    // "제품"/"제품명"이라는 값을 가진 셀을 정규표현식을 사용하여 찾기
    let sw_product_cell_address = find_cell_with_regex(sw_summary_sheet, /제품/);
    let validation_product_cell_address = find_cell_with_regex(latest_validation_sheet, /제품명/);
  
    if (sw_product_cell_address && validation_product_cell_address) {
      console.log(`SW현황 "제품" 셀이 위치한 행과 열: ${sw_product_cell_address.row}, ${sw_product_cell_address.column}`);
      console.log(`최신 검증서 "제품명" 셀이 위치한 행과 열: ${validation_product_cell_address.row}, ${validation_product_cell_address.column}`);
  
      // "제품" 및 "제품명" 셀 하위의 모든 열 데이터를 추출
      let sw_product_names = get_column_data(sw_summary_sheet, sw_product_cell_address.row + 1, sw_product_cell_address.column+1);
      let validation_product_names = get_column_data(latest_validation_sheet, validation_product_cell_address.row + 1, validation_product_cell_address.column);
  
      // 데이터 출력
      console.log("SW현황 시트 제품명 데이터:", sw_product_names);
      console.log("최신 검증서 시트 제품명 데이터:", validation_product_names);
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
  