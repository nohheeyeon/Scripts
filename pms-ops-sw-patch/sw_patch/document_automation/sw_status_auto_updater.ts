function main(workbook: ExcelScript.Workbook) {
    // SW 현황 및 검증서 불러오기
    let sheets = workbook.getWorksheets();
    
    // SW현황과 최신 검증서 가져오기
    let ws1 = sheets[0]; // SW현황
    let ws4 = sheets[3]; // 최신 검증서

    // SW현황의 D열(제품명) 데이터를 딕셔너리로 저장
    let d_column_values: { [key: string]: number } = {};
    let last_row_ws1 = ws1.getUsedRange().getRowCount();
    let d_column_range = ws1.getRange(`D2:D${last_row_ws1}`);
    d_column_range.getValues().forEach((row, index) => {
        let cell_values = row[0];
        if (cell_values !== null && cell_values !== undefined && cell_values !== "") {
            d_column_values[String(cell_values) ] = index + 2;
        }
    });

    // 최신 검증서의 E열(제품명) 데이터와 비교하여 동일한 값이 있는 행의 데이터를 SW현황으로 복사
    let last_row_ws4 = ws4.getUsedRange(). getRowCount();
    let e_column_range = ws4.getRange(`E2:E${last_row_ws4}`);
    e_column_range.getValues().forEach((row, index) => {
        let cell_values = row[0];
        if (cell_values !== null && cell_values !== undefined && cell_values !== "" && d_column_values[String(cell_values)] !== undefined) {
            let row_in_ws1: number = d_column_values[String(cell_values)];
            let row_in_ws4: number = index + 2;

            // 해당하는 셀 복사 작업
            ws1.getCell(row_in_ws1 - 1, 5).setValue(ws4.getCell(row_in_ws4 - 1, 3).getValue()); // D열(제품명) 값 복사
            ws1.getCell(row_in_ws1 - 1, 6).setValue(ws4.getCell(row_in_ws4 - 1, 7).getValue()); // H열(버전) 값 복사
            ws1.getCell(row_in_ws1 - 1, 7).setValue(ws4.getCell(row_in_ws4 - 1, 8).getValue()); // I열(패치 파일명) 값 복사

            // 로그 출력
            console.log(`동일한 값 발견: ${cell_values} (SW현황의 ${row_in_ws1}행, 최신 검증서의 ${row_in_ws4}행)`)
            console.log(`최신 검증서의 D열 값 '${ws4.getCell(row_in_ws4 - 1, 3).getValue()}'을 SW현황의 F열로 복사`);
            console.log(`최신 검증서의 H열 값 '${ws4.getCell(row_in_ws4 - 1, 7).getValue()}'을 SW현황의 G열로 복사`);
            console.log(`최신 검증서의 I열 값 '${ws4.getCell(row_in_ws4 - 1, 8).getValue()}'을 SW현황의 H열로 복사`);
        }
    })
}