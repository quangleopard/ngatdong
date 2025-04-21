async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    console.log(`Script đang chạy trên sheet: ${sheet.name}`);

    // Thiết lập sự kiện onChanged cho ô A1
    sheet.onChanged.add(async (event) => {
      if (event.address === `A1`) {
        console.log("Ô A1 đã thay đổi. Bắt đầu thực hiện logic...");
        await Excel.run(async (innerContext) => {
          const innerSheet = innerContext.workbook.worksheets.getActiveWorksheet();

          async function selectCell(range: Excel.Range, value: number) {
            range.load("address");
            range.load("values");
            await innerContext.sync();

            const rangeValues = range.values;
            if (rangeValues) {
              for (let i = 0; i < rangeValues.length; i++) {
                if (rangeValues[i][0] === value) {
                  const cell = range.getCell(i, 0);
                  cell.select();
                  return;
                }
              }
            }
          }
          const horizontalPageBreaks = innerSheet.horizontalPageBreaks.load("items");
          await innerContext.sync();
          horizontalPageBreaks.items.forEach(pageBreak => {
            if (!pageBreak) {
              pageBreak.delete();
            }
          });
          const verticalPageBreaks = innerSheet.verticalPageBreaks.load("items");
          await innerContext.sync();
          verticalPageBreaks.items.forEach(pageBreak => {
            if (!pageBreak) {
              pageBreak.delete();
            }
          });

          await context.sync();

          const range = innerSheet.getRange("Q5:Q4000");
          innerSheet.autoFilter.apply(range, 0, { criterion1: "1", filterOn: Excel.FilterOn.custom });
          await innerContext.sync();
          const I1 = innerSheet.getRange("I1");
          I1.load("values");
          await innerContext.sync();
          const I1value = I1.values[0][0];
          console.log(`Giá trị của I1 là: ${I1value}`);

          let rowstart = 1;
          const r1 = innerSheet.getRange("S1:S400");

          while (I1value - 30 > rowstart) {
            let rownext = rowstart + 31;
            await selectCell(r1, rownext);
            const jcell = innerContext.workbook.getActiveCell();
            jcell.load("address");
            jcell.load("values");
            await innerContext.sync();
            console.log(`Đã chọn ô: ${jcell.address}`);

            const xkcell = jcell.getOffsetRange(0, -11);
            xkcell.load("values");
            await innerContext.sync();
            const xkvalue = xkcell.values ? xkcell.values[0][0] : undefined;
            console.log(`Giá trị tại offset -11: ${xkvalue}`);

            if (!xkvalue || !String(xkvalue).startsWith("XK")) {
              console.log("Ô không bắt đầu bằng XK. Bắt đầu kiểm tra các hàng phía trên...");
              for (let m = 1; m <= 30; m++) {
                const rowup = rownext - 1;
                await selectCell(r1, rowup);
                const breakcell = innerContext.workbook.getActiveCell();
                breakcell.load("address");
                breakcell.load("values");
                await innerContext.sync();
                console.log(`Đã chọn ô để kiểm tra XK: ${breakcell.address}`);

                const checkbreakcell = breakcell.getOffsetRange(0, -11);
                checkbreakcell.load("values");
                await innerContext.sync();
                const xkbreak = checkbreakcell.values ? checkbreakcell.values[0][0] : undefined;
                console.log(`Giá trị XK ở hàng trên: ${xkbreak}`);

                if (xkbreak && String(xkbreak).startsWith("XK")) {
                  innerSheet.horizontalPageBreaks.add(breakcell);
                  await innerContext.sync();
                  console.log(`Đã thêm dấu ngắt trang tại: ${breakcell.address}`);
                  break;
                }
                rownext = rowup;
                await innerContext.sync();
              }
              console.log("Đã hoàn thành kiểm tra các hàng phía trên.");
            }
            rowstart = rownext;
            await innerContext.sync();
          }
          console.log("Hoàn thành logic xử lý.");
        });
      }
    });

    console.log("Sự kiện onChanged đã được thiết lập cho ô A1.");
    await context.sync();
  });
}

run();