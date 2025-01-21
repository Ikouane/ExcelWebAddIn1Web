// 每次加载新页面时都必须运行初始化函数。
Office.onReady(() => {
        // 如果你需要初始化，可以在此进行。
});

async function sampleFunction(event) { 
const values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
        await Excel.run(async (context) => {
                // Write sample values to a range in the active worksheet.
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("B3:D5").values = values;
                await context.sync();
        });
        } catch (error) {
        console.log(error.message);
        }
        // 需要调用 event.completed。event.completed 会让平台知道处理已完成。
        event.completed();
}
