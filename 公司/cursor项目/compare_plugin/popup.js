// 检查 ExcelJS 是否正确加载
window.addEventListener('load', function() {
    if (typeof ExcelJS === 'undefined') {
        console.error('ExcelJS 未能正确加载');
        document.getElementById('status').textContent = 'ExcelJS 加载失败，请刷新页面重试';
        return;
    }
    console.log('ExcelJS 加载成功');
    initializeApp();
});

function initializeApp() {
    console.log('DOM Content Loaded');
    
    const compareButton = document.getElementById('compareButton');
    const status = document.getElementById('status');
    const resultArea = document.getElementById('result');

    // 读取Excel文件
    async function readExcelFile(file) {
        console.log('开始读取文件:', file.name);
        try {
            const workbook = new ExcelJS.Workbook();
            const arrayBuffer = await file.arrayBuffer();
            console.log('文件已转换为 ArrayBuffer');
            const result = await workbook.xlsx.load(arrayBuffer);
            console.log('文件加载成功');
            return result;
        } catch (error) {
            console.error('读取文件失败:', error);
            throw error;
        }
    }

    async function compareExcelFiles(inputWb, outputWb, inputCol, outputCol, tagValue) {
        const inputSheet = inputWb.worksheets[0];
        const outputSheet = outputWb.worksheets[0];
        
        // 收集输入表中的值
        const inputValues = new Set();
        let inputTotal = 0;
        
        inputSheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // 跳过标题行
                const cell = row.getCell(inputCol);
                if (cell && cell.value !== null && cell.value !== undefined) {
                    inputValues.add(cell.value.toString());
                    inputTotal++;
                }
            }
            updateProgress('收集数据', (rowNumber / inputSheet.rowCount) * 100);
        });

        // 创建新的工作簿
        const newWb = new ExcelJS.Workbook();
        // 复制工作簿属性
        newWb.creator = outputWb.creator;
        newWb.lastModifiedBy = outputWb.lastModifiedBy;
        newWb.created = outputWb.created;
        newWb.modified = outputWb.modified;

        // 创建新工作表
        const newSheet = newWb.addWorksheet(outputSheet.name, {
            properties: { ...outputSheet.properties },
            views: [...outputSheet.views],
            pageSetup: { ...outputSheet.pageSetup }
        });

        // 复制列宽
        outputSheet.columns.forEach((col, index) => {
            if (col) {
                newSheet.getColumn(index + 1).width = col.width || 10;
            }
        });
        
        // 设置新列的宽度
        newSheet.getColumn(outputSheet.columnCount + 1).width = 10;

        // 先创建所有行
        for (let i = 1; i <= outputSheet.rowCount; i++) {
            newSheet.addRow([]);
        }

        // 复制所有行和单元格（包括样式和值）
        outputSheet.eachRow((row, rowNumber) => {
            const newRow = newSheet.getRow(rowNumber);
            
            // 复制行高和行格式
            newRow.height = row.height;
            newRow.hidden = row.hidden;
            newRow.outlineLevel = row.outlineLevel;
            
            // 复制每个单元格
            row.eachCell((cell, colNumber) => {
                try {
                    const newCell = newRow.getCell(colNumber);
                    
                    // 复制值和类型
                    newCell.value = cell.value;
                    
                    // 复制样式
                    if (cell.style) {
                        newCell.style = JSON.parse(JSON.stringify(cell.style));
                    }
                } catch (error) {
                    console.error(`复制单元格出错: 行=${rowNumber}, 列=${colNumber}`, error);
                }
            });
        });

        // 添加标题
        try {
            const headerRowNumber = 1;
            const headerColNumber = outputSheet.columnCount + 1;
            const headerRow = newSheet.getRow(headerRowNumber);
            const headerCell = headerRow.getCell(headerColNumber);
            const sourceHeaderCell = outputSheet.getRow(1).getCell(1);
            
            headerCell.value = '标记';
            if (sourceHeaderCell.style) {
                headerCell.style = JSON.parse(JSON.stringify(sourceHeaderCell.style));
            }
        } catch (error) {
            console.error('添加标题出错:', error);
        }

        // 处理匹配项
        let outputTotal = 0;
        let matchCount = 0;

        outputSheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // 跳过标题行
                const checkCell = row.getCell(outputCol);
                if (checkCell && checkCell.value !== null && checkCell.value !== undefined) {
                    outputTotal++;
                    if (inputValues.has(checkCell.value.toString())) {
                        try {
                            const newRow = newSheet.getRow(rowNumber);
                            const targetCell = newRow.getCell(outputSheet.columnCount + 1);
                            targetCell.value = tagValue;
                            if (checkCell.style) {
                                targetCell.style = JSON.parse(JSON.stringify(checkCell.style));
                            }
                            matchCount++;
                        } catch (error) {
                            console.error(`处理匹配项出错: 行=${rowNumber}`, error);
                        }
                    }
                }
            }
            updateProgress('标记数据', (rowNumber / outputSheet.rowCount) * 100);
        });

        // 显示结果统计
        showResults({
            inputTotal,
            outputTotal,
            matchCount
        });

        return newWb;
    }

    async function downloadExcel(workbook, originalFilename) {
        const buffer = await workbook.xlsx.writeBuffer({
            filename: originalFilename.replace(/\.xlsx?$/, '_已标记.xlsx'),
            useStyles: true,
            useSharedStrings: true
        });
        
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = originalFilename.replace(/\.xlsx?$/, '_已标记.xlsx');
        a.click();
        window.URL.revokeObjectURL(url);
    }

    function updateProgress(text, percent) {
        status.textContent = `${text} (${Math.round(percent)}%)`;
    }

    function showResults(stats) {
        resultArea.innerHTML = `
            <div class="result-item">输入文件数据：
                <span class="result-highlight">${stats.inputTotal}</span> 行
            </div>
            <div class="result-item">输出文件数据：
                <span class="result-highlight">${stats.outputTotal}</span> 行
            </div>
            <div class="result-item">匹配成功：
                <span class="result-highlight">${stats.matchCount}</span> 项
            </div>
        `;
    }

    compareButton.addEventListener('click', async function() {
        console.log('点击了比对按钮');
        try {
            status.textContent = '正在处理...';
            resultArea.innerHTML = '';
            
            const inputFile = document.getElementById('inputFile').files[0];
            const outputFile = document.getElementById('outputFile').files[0];
            
            console.log('输入文件:', inputFile?.name);
            console.log('输出文件:', outputFile?.name);
            
            if (!inputFile || !outputFile) {
                throw new Error('请选择输入和输出文件');
            }

            const inputColumn = parseInt(document.getElementById('inputColumn').value);
            const outputColumn = parseInt(document.getElementById('outputColumn').value);
            const tagValue = document.getElementById('tagValue').value || 'T';

            console.log('开始处理文件...');
            console.log('输入列:', inputColumn);
            console.log('输出列:', outputColumn);
            console.log('标记值:', tagValue);

            const inputData = await readExcelFile(inputFile);
            console.log('输入文件读取完成');
            
            const outputData = await readExcelFile(outputFile);
            console.log('输出文件读取完成');
            
            const processedWorkbook = await compareExcelFiles(inputData, outputData, inputColumn, outputColumn, tagValue);
            console.log('文件比对完成');
            
            await downloadExcel(processedWorkbook, outputFile.name);
            console.log('文件下载完成');
            
            status.textContent = '处理完成！';
        } catch (error) {
            console.error('处理过程出错:', error);
            status.textContent = '错误：' + error.message;
        }
    });

    console.log('事件监听器已设置');
}
