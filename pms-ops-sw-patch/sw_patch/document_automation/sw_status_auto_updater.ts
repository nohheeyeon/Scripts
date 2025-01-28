function main(workbook: ExcelScript.Workbook) {
    // SW 현황 및 검증서 불러오기
    let sheets = workbook.getWorksheets();
    
    // SW현황과 최신 검증서 가져오기
    let sw_summary_sheet = sheets[0]; // SW현황
    let lastest_validation_sheet = sheets[3]; // 최신 검증서

    // SW현황의 D열(제품명) 데이터를 딕셔너리로 저장
    let patch_name_column_values: { [key: string]: number } = {};
    let lastest_validation_sheet_row_sw_summary_sheet = sw_summary_sheet.getUsedRange().getRowCount();
    let patch_name_column_range = sw_summary_sheet.getRange(`D2:D${lastest_validation_sheet_row_sw_summary_sheet}`);
    patch_name_column_range.getValues().forEach((row, index) => {
        let patch_name_values = row[0];
        if (patch_name_values !== null && patch_name_values !== undefined && patch_name_values !== "") {
            patch_name_column_values[String(patch_name_values) ] = index + 2;
        }
    });

    // 최신 검증서의 E열(제품명) 데이터와 비교하여 동일한 값이 있는 행의 데이터를 SW현황으로 복사
    let lastest_validation_sheet_row_lastest_validation_sheet = lastest_validation_sheet.getUsedRange(). getRowCount();
    let e_column_range = lastest_validation_sheet.getRange(`E2:E${lastest_validation_sheet_row_lastest_validation_sheet}`);
    e_column_range.getValues().forEach((row, index) => {
        let patch_name_values = row[0];
        if (patch_name_values !== null && patch_name_values !== undefined && patch_name_values !== "" && patch_name_values[String(patch_name_values)] !== undefined) {
            let row_in_sw_summary_sheet: number = patch_name_column_values[String(patch_name_values)];
            let row_in_lastest_validation_sheet: number = index + 2;

            // 해당하는 셀 복사 작업
            sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 5).setValue(lastest_validation_sheet.getCell(row_in_ws4 - 1, 3).getValue()); // D열(발표일) 값 복사
            sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 6).setValue(lastest_validation_sheet.getCell(row_in_ws4 - 1, 7).getValue()); // H열(버전) 값 복사
            sw_summary_sheet.getCell(row_in_sw_summary_sheet - 1, 7).setValue(lastest_validation_sheet.getCell(row_in_ws4 - 1, 8).getValue()); // I열(패치 파일명) 값 복사

            // 로그 출력
            console.log(`동일한 값 발견: ${patch_name_values} (SW현황의 ${row_in_sw_summary_sheet}행, 최신 검증서의 ${row_in_lastest_validation_sheet}행)`)
            console.log(`최신 검증서의 D열 값 '${lastest_validation_sheet.getCell(row_in_lastest_validation_sheet - 1, 3).getValue()}'을 SW현황의 F열로 복사`);
            console.log(`최신 검증서의 H열 값 '${lastest_validation_sheet.getCell(row_in_lastest_validation_sheet - 1, 7).getValue()}'을 SW현황의 G열로 복사`);
            console.log(`최신 검증서의 I열 값 '${lastest_validation_sheet.getCell(row_in_lastest_validation_sheet - 1, 8).getValue()}'을 SW현황의 H열로 복사`);
        }
    })
}