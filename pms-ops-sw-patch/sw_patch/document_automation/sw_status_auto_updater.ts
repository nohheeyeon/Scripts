function main(workbook: ExcelScript.Workbook) {
    // "PMS SW 패치 현황 및 검증서" 불러오기
    let sheets = workbook.getWorksheets();
  
    // SW현황 시트를 가져오기
    let sw_summary_sheet = find_sheet_by_regex(sheets, /^SW현황$/);

    if (!sw_summary_sheet) {
        console.log("SW현황 시트를 찾을 수 없습니다.")
        return;
    }

    // 최신 검증서 시트를 가져오기
    let latest_validation_sheet = find_sheet_by_regex(sheets, /^\d{4}-\{2}-\d{2} 검증서$/);

    if (!latest_validation_sheet) {
        console.log("최신 검증서 시트를 찾을 수 없습니다.")
        return;
    }

    function find_sheet_by_regex(sheets: ExcelScript.Worksheet[], regex: RegExp): ExcelScript.Worksheet | null {
        for (let sheet of sheets) {
            if (regex.test(sheet.getName())) {
                return sheet;
            }
        }
        return null;
    }

    enum SWColumns {
        kProductname = "제품명",
        kReleaseDate = "발표일",
        KVersion = "버전",
        KPatchFileName = "패치 파일명"
    }


    enum ValidationColumns {
        kProductname = "제품명",
        kReleaseDat가e = "발표일",
        KVersion = "버전",
        KPatchFileName = "패치 파일명"
    }

    // SW현황의 제품명 데이터를 가져오기
    let sw_patch_names = get_patch_names(sw_summary_sheet, 'D');
  
    // 최신 검증서의 제품명 데이터를 가져오기
    let validation_patch_names = get_patch_names(latest_validation_sheet, 'F');
    
    // 제품명을 비교하여 데이터를 업데이트
    update_sw_summary_sheet(sw_summary_sheet, latest_validation_sheet, sw_patch_names, validation_patch_names);
  }
  
  // SW현황의 D열(제품명) 데이터를 딕셔너리로 저장
  function get_patch_names(sheet: ExcelScript.Worksheet, column: string): { [key: string]: number } {
    let patch_name_column_values: { [key: string]: number } = {};
    let total_row_count = sheet.getUsedRange().getRowCount();
    let patch_name_column_range = sheet.getRange(`${column}2:${column}${total_row_count}`);
    
    patch_name_column_range.getValues().forEach((row, index) => {
      let patch_name_values = row[0];
      if (patch_name_values !== null && patch_name_values !== undefined && patch_name_values !== "") {
        patch_name_column_values[String(patch_name_values)] = index + 2;
      }
    });
    
    return patch_name_column_values;
  }
  
  // 최신 검증서의 F열(제품명) 데이터와 비교하여 동일한 값이 있는 행의 데이터를 SW현황으로 복사
  function update_sw_summary_sheet(
    sw_summary_sheet: ExcelScript.Worksheet,
    latest_validation_sheet: ExcelScript.Worksheet,
    sw_patch_names: { [key: string]: number },
    validation_patch_names: { [key: string]: number }
  ) {
    Object.keys(validation_patch_names).forEach(patch_name_values => {
      if (sw_patch_names[patch_name_values] !== undefined) {
        let row_in_sw_summary_sheet: number = sw_patch_names[patch_name_values];
        let row_in_latest_validation_sheet: number = validation_patch_names[patch_name_values];
  
        // 해당하는 셀 복사 작업
        sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 5).setValue(latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 3).getValue()); // D열(발표일) 값 복사
        sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 6).setValue(latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 8).getValue()); // I열(버전) 값 복사
        sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 7).setValue(latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 9).getValue()); // J열(패치 파일명) 값 복사
  
        // 로그 출력
        console.log(`동일한 값 발견: ${patch_name_values} (SW현황의 ${row_in_sw_summary_sheet}행, 최신 검증서의 ${row_in_latest_validation_sheet}행)`);
        console.log(`최신 검증서의 D열 값 '${latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 3).getValue()}'을 SW현황의 F열로 복사`);
        console.log(`최신 검증서의 I열 값 '${latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 8).getValue()}'을 SW현황의 G열로 복사`);
        console.log(`최신 검증서의 J열 값 '${latest_validation_sheet.getCell(row_in_latest_validation_sheet - 1, 9).getValue()}'을 SW현황의 H열로 복사`);
      }
    });
  }
  