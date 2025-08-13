document.getElementById('modify-cookie').addEventListener('click', () => {
    const currentCookie = localStorage.getItem('cookie') || '';
    const newCookie = prompt('当前Cookie:\n' + currentCookie + '\n\n请输入新的Cookie:');
    if (newCookie !== null) {
        localStorage.setItem('cookie', newCookie);
        alert('Cookie已更新！');
    }
});

document.getElementById('import-button').addEventListener('click', () => {
    const fileInput = document.getElementById('file-input');
    const file = fileInput.files[0];
    if (!file) {
        alert('请先选择一个Excel文件。');
        return;
    }

    const reader = new FileReader();
    reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

        // 处理数据
        await processData(jsonData);
    };
    reader.readAsArrayBuffer(file);
});

async function processData(data) {
    const project = document.getElementById('project').value;
    const cookie = localStorage.getItem('cookie') || '';

    for (const row of data) {
        const cid = row.cid;
        if (!cid) continue;

        const params = {
            cid: cid,
            partner_code: project,
            ...row // 其他字段
        };

        const response = await fetch('http://cms.enjoy-tv.cn/api/project/cover/edit', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Cookie': cookie
            },
            body: JSON.stringify(params)
        });

        if (response.ok) {
            document.getElementById('status').innerText += `处理成功: ${cid}\n`;
        } else {
            document.getElementById('status').innerText += `处理失败: ${cid}\n`;
        }

        // 更新进度条
        const progress = Math.round((data.indexOf(row) + 1) / data.length * 100);
        document.getElementById('progress').value = progress;
    }
}