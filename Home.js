
    let cellToHighlight;
    let messageBanner;

    // Office JS 和 JQuery 准备就绪时初始化。
    Office.onReady(() => {
        $(() => {
            // 初始化并隐藏 Office Fabric UI 通知机制。
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // 如果未使用 Excel 2016 或更高版本，请使用回退逻辑。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("此示例将显示电子表格中选定单元格的值。");
                $('#button-text').text("显示!");
                $('#button-desc').text("显示所选内容");

                $('#highlight-button').on('click',displaySelectedCells);
                return;
            }

            $("#template-description").text("此示例将突出显示电子表格中选定单元格的最低值。");
            $('#button-text').text("突出显示!");
            $('#button-desc').text("突出显示最小数字。");
                
            loadSampleData();

            // 为突出显示按钮添加单击事件处理程序。
            $('#highlight-button').on('click',highlightHighestValue);
        });
    });

    async function loadSampleData() {
        const values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                // 将示例值写入活动工作表中的范围
                sheet.getRange("B3:D5").values = values;
                await context.sync();
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function highlightHighestValue() {
        try {
            await Excel.run(async (context) => {
                const sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");

                await context.sync();
                let highestRow = 0;
                let highestCol = 0;
                let highestValue = sourceRange.values[0][0];

                // 找到要突出显示的单元格
                for (let i = 0; i < sourceRange.rowCount; i++) {
                    for (let j = 0; j < sourceRange.columnCount; j++) {
                        if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] < highestValue) {
                            highestRow = i;
                            highestCol = j;
                            highestValue = sourceRange.values[i][j];
                        }
                    }
                }

                cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                sourceRange.worksheet.getUsedRange().format.fill.clear();
                sourceRange.worksheet.getUsedRange().format.font.bold = false;

                // 突出显示该单元格
                cellToHighlight.format.fill.color = "orange";
                cellToHighlight.format.font.bold = true;
                await context.sync;
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function displaySelectedCells() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("text");
                await context.sync();
                const textValue = range.text.toString();
                showNotification('选定的文本为:', '"' + textvalue + '"');
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    // 处理错误的帮助程序函数
    function errorHandler(error) {
        // 请务必捕获 Excel.run 执行过程中出现的所有累积错误
        showNotification("错误", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 用于显示通知的帮助程序函数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
