// 首先需要安装 xlsx 库
// npm install xlsx --save

import * as XLSX from 'xlsx';

// 导出函数
function exportToExcel(data, fileName = 'export.xlsx') {
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 将数据转换为工作表
    const ws = XLSX.utils.json_to_sheet(data);
    
    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    // 生成文件并下载
    XLSX.writeFile(wb, fileName);
}

// 使用示例
async function handleExport() {
    try {
        // 显示加载提示
        showLoading();
        
        // 获取所有数据（如果需要的话）
        const allData = await getAllData();
        
        // 导出
        exportToExcel(allData, '影片数据.xlsx');
    } catch (error) {
        console.error('导出失败:', error);
        // 显示错误提示
        showError('导出失败，请重试');
    } finally {
        // 隐藏加载提示
        hideLoading();
    }
}

function getCurrentFilters() {
    // 这里需要根据实际的页面元素ID或类名来获取值
    return {
        cid: document.querySelector('#cid')?.value || '',
        title: document.querySelector('#title')?.value || '',
        channel_name: document.querySelector('#channel_name')?.value || '',
        // ... 其他筛选条件
    };
}

async function getAllData() {
    try {
        // 获取最后一次请求的参数，只修改 offset 和 limit
        const lastRequestBody = {
            cid: "",                // 使用当前页面上的值
            title: "",             // 使用当前页面上的值
            channel_name: "",      // 使用当前页面上的值
            m4_status: "",         // 使用当前页面上的值
            m8_status: "",         // 使用当前页面上的值
            m6_status: "",         // 使用当前页面上的值
            m2_status: "",         // 使用当前页面上的值
            batch: "",             // 使用当前页面上的值
            is_effective: "",      // 使用当前页面上的值
            is_finished: "",       // 使用当前页面上的值
            offset: 0,
            limit: 99999
        };
        
        const response = await fetch('http://cms.enjoy-tv.cn/query/cover/list', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(lastRequestBody)
        });
        
        const data = await response.json();
        return data.list || data.data || data;
    } catch (error) {
        console.error('获取数据失败:', error);
        throw error;
    }
} 