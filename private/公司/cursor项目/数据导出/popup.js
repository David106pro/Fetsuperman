document.getElementById('api_type').addEventListener('change', function() {
    const totalCoverConditions = document.getElementById('total_cover_conditions');
    const timeRangeGroup = document.getElementById('time_range_group');
    const partnerGroup = document.getElementById('partner_group');
    const projectVideoFields = document.getElementById('project_video_fields');
    
    // 控制总库特有条件的显示
    totalCoverConditions.style.display = this.value === 'total_cover' ? 'block' : 'none';
    
    // 控制时间范围的显示（只在注入库时显示）
    timeRangeGroup.style.display = 
        this.value === 'inject_cover' || this.value === 'inject_video' 
        ? 'block' 
        : 'none';

    // 控制合作伙伴的显示（在总库时隐藏）
    partnerGroup.style.display = this.value === 'total_cover' ? 'none' : 'block';

    // 控制项目库-子集字段选择的显示
    projectVideoFields.style.display = this.value === 'project_video' ? 'block' : 'none';
});

// 将 processData 函数移到事件处理函数外部
const processData = (list, api_type) => {
    const status = document.getElementById('status');
    status.textContent = '正在处理数据...';
    
    switch (api_type) {
        case 'project_cover':
            status.textContent = '正在处理项目库-剧头数据...';
            // 项目库-剧头只保留指定字段
            return list.map(item => {
                // 只保留指定的7个字段
                const result = {
                    cid: item.cid || '',
                    title: item.title || '',
                    channel_name: item.channel_name || '',
                    is_online: item.is_online || '',
                    task_status: item.task_status || '',
                    c_time: item.c_time || '',
                    m_time: item.m_time || ''
                };
                // 确保只返回这7个字段
                return Object.fromEntries(
                    Object.entries(result).filter(([key]) => 
                        ['cid', 'title', 'channel_name', 'is_online', 'task_status', 'c_time', 'm_time'].includes(key)
                    )
                );
            });
        
        case 'project_video':
            status.textContent = '正在处理项目库-子集数据...';
            // 项目库-子集保留基础字段
            return list.map(item => {
                const baseFields = {
                    cid: item.cid || '',
                    vid: item.vid || '',
                    title: item.title || '',
                    channel_name: item.channel_name || '',
                    tv_is_online: item.tv_is_online || '',
                    task_status: item.task_status || '',
                    c_time: item.c_time || '',
                    m_time: item.m_time || ''
                };

                // 根据选择添加额外字段
                const includeInjectId = document.getElementById('field_inject_id').checked;
                const includeInjectTime = document.getElementById('field_inject_time').checked;

                if (includeInjectId) {
                    baseFields.series_id = item.series_id || '';
                    baseFields.program_id = item.program_id || '';
                    baseFields.movie_id = item.movie_id || '';
                }
                
                if (includeInjectTime) {
                    baseFields.inject_send_time = item.inject_send_time || '';
                    baseFields.inject_receive_time = item.inject_receive_time || '';
                }
                
                return baseFields;
            });
        
        default:
            // 其他数据来源保持原样
            return list;
    }
};

document.getElementById('exportBtn').addEventListener('click', async () => {
    const status = document.getElementById('status');
    status.textContent = '正在导出...';
    
    try {
        const api_type = document.getElementById('api_type').value;
        if (!api_type) {
            throw new Error('请选择数据来源');
        }

        const partner_code = document.getElementById('partner_code').value;
        const channel_name = document.getElementById('channel_name').value;
        
        // 处理创建时间范围
        let c_start_time = document.getElementById('c_start_time').value;
        let c_end_time = document.getElementById('c_end_time').value;
        if (c_start_time) {
            c_start_time = c_start_time + ' 00:00:00';
        }
        if (c_end_time) {
            c_end_time = c_end_time + ' 23:59:59';
        }

        // 处理修改时间范围
        let m_start_time = document.getElementById('m_start_time').value;
        let m_end_time = document.getElementById('m_end_time').value;
        if (m_start_time) {
            m_start_time = m_start_time + ' 00:00:00';
        }
        if (m_end_time) {
            m_end_time = m_end_time + ' 23:59:59';
        }

        // 处理CID列表
        const cidText = document.getElementById('cid_list').value;
        const cidList = cidText
            .split('\n')
            .map(cid => cid.trim())
            .filter(cid => cid !== '')
            .slice(0, 100);

        const filters = {
            channel_name,
            offset: 0,
            limit: 1000  // 统一设置为1000条
        };

        // 只有非总库时才添加合作伙伴条件
        if (api_type !== 'total_cover') {
            filters.partner_code = partner_code;
        }

        // 初始化文件名
        let fileName = getFileNamePrefix(api_type);
        // 只有非总库时才在文件名中添加合作伙伴
        if (api_type !== 'total_cover' && partner_code) {
            fileName += `_${partner_code}`;
        }
        if (channel_name) {
            fileName += `_${channel_name}`;
        }
        if (cidList.length > 0) {
            fileName += `_${cidList.length}个CID`;
        }

        // 只有在注入库时才添加时间范围和状态条件
        if (api_type.startsWith('inject_')) {
            filters.c_start_time = c_start_time;
            filters.c_end_time = c_end_time;
            filters.m_start_time = m_start_time;
            filters.m_end_time = m_end_time;

            // 添加注入状态和上线状态
            const task_status = document.getElementById('task_status').value;
            const is_online = document.getElementById('is_online').value;

            if (task_status) {
                filters.task_status = task_status === 'other' ? '-1' : task_status;
            }
            if (is_online) {
                filters.is_online = is_online;
            }

            // 在文件名中添加状态信息
            if (task_status) {
                const statusMap = {
                    '1': '待注入',
                    '4': '注入中',
                    '5': '注入成功',
                    'other': '其它状态'
                };
                fileName += `_${statusMap[task_status]}`;
            }
            if (is_online) {
                fileName += `_${is_online === '1' ? '上线' : '下线'}`;
            }
        }

        // 如果是总库-专辑，添加特定条件
        if (api_type === 'total_cover') {
            const is_effective = document.getElementById('is_effective').value;
            const m4_status = document.getElementById('m4_status').value;
            const m8_status = document.getElementById('m8_status').value;
            const is_finished = document.getElementById('is_finished').value;
            const batch = document.getElementById('batch').value;

            if (is_effective) filters.is_effective = is_effective;
            if (m4_status) filters.m4_status = m4_status;
            if (m8_status) filters.m8_status = m8_status;
            if (is_finished) filters.is_finished = is_finished;
            if (batch) filters.batch = batch;

            // 添加条件到文件名
            const conditions = [];
            if (filters.is_effective) {
                conditions.push(filters.is_effective === '1' ? '有效' : '无效');
            }
            if (filters.m4_status) {
                conditions.push('4M' + (filters.m4_status === '1' ? '有效' : '无效'));
            }
            if (filters.m8_status) {
                conditions.push('8M' + (filters.m8_status === '1' ? '有效' : '无效'));
            }
            if (filters.is_finished) {
                const finishedMap = { '0': '无', '1': '一般', '2': '完整' };
                conditions.push('完整性' + finishedMap[filters.is_finished]);
            }
            if (filters.batch) {
                const batchMap = {
                    '240626': '批次1',
                    '240725': '批次2',
                    '240816': '批次3',
                    '240905': '批次4',
                    '241015': '批次5',
                    '241122': '批次6',
                    '241212': '批次7',
                    '240101': '新片'
                };
                conditions.push(batchMap[filters.batch]);
            }
            if (conditions.length > 0) {
                fileName += '_' + conditions.join('_');
            }
        }

        // 添加时间范围到文件名
        if (c_start_time || c_end_time) {
            // 只保留日期部分 YYYY-MM-DD
            const startDate = c_start_time ? c_start_time.split(' ')[0] : '';
            const endDate = c_end_time ? c_end_time.split(' ')[0] : '';
            fileName += '_创建' + startDate + (endDate ? '至' + endDate : '');
        }
        if (m_start_time || m_end_time) {
            // 只保留日期部分 YYYY-MM-DD
            const startDate = m_start_time ? m_start_time.split(' ')[0] : '';
            const endDate = m_end_time ? m_end_time.split(' ')[0] : '';
            fileName += '_修改' + startDate + (endDate ? '至' + endDate : '');
        }
        
        fileName += '_' + new Date().toISOString().split('T')[0];
        fileName += '.xlsx';

        if (cidList.length > 0) {
            filters.cid_list = cidList;
        }

        // 使用新的 API 地址获取函数
        const apiUrl = getApiUrl(api_type);

        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Origin': 'http://cms.enjoy-tv.cn'
            },
            credentials: 'include',
            mode: 'cors',
            body: JSON.stringify(filters)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        
        if (data.code !== 'A000000') {
            throw new Error('API返回错误: ' + data.code);
        }

        const list = data.rows;
        if (!Array.isArray(list)) {
            throw new Error('数据格式不正确');
        }

        if (list.length === 0) {
            throw new Error('没有找到符合条件的数据');
        }

        if (data.total > filters.limit) {
            status.textContent = `注意：共找到 ${data.total} 条数据，当前仅导出前1000条`;
            await new Promise(resolve => setTimeout(resolve, 2000));
        }

        // 处理数据，只保留需要的字段
        const processedList = processData(list, api_type);
        console.log('Final processed data first item:', processedList[0]);

        // 修改列宽设置
        const getColumnWidths = (api_type) => {
            switch (api_type) {
                case 'project_cover':
                    return [
                        {wch: 15}, // cid
                        {wch: 30}, // title
                        {wch: 15}, // channel_name
                        {wch: 10}, // is_online
                        {wch: 12}, // task_status
                        {wch: 20}, // c_time
                        {wch: 20}  // m_time
                    ];
                
                case 'project_video':
                    let widths = [
                        {wch: 15}, // cid
                        {wch: 15}, // vid
                        {wch: 30}, // title
                        {wch: 15}, // channel_name
                        {wch: 10}, // tv_is_online
                        {wch: 12}, // task_status
                        {wch: 20}, // c_time
                        {wch: 20}  // m_time
                    ];

                    // 根据选择添加额外列宽
                    if (document.getElementById('field_inject_id').checked) {
                        widths = widths.concat([
                            {wch: 20}, // series_id
                            {wch: 20}, // program_id
                            {wch: 20}  // movie_id
                        ]);
                    }
                    
                    if (document.getElementById('field_inject_time').checked) {
                        widths = widths.concat([
                            {wch: 20}, // inject_send_time
                            {wch: 20}  // inject_receive_time
                        ]);
                    }
                    
                    return widths;
                
                default:
                    return [
                        // 原有的列宽设置
                        {wch: 10}, // id
                        {wch: 15}, // cid
                        {wch: 30}, // title
                        {wch: 8},  // rate
                        {wch: 10}, // is_effective
                        {wch: 10}, // is_online
                        {wch: 15}, // channel_name
                        {wch: 12}, // task_priority
                        {wch: 12}, // task_status
                        {wch: 20}, // series_id
                        {wch: 30}, // xml_name
                        {wch: 50}, // receive_xml_url
                        {wch: 20}, // inject_send_time
                        {wch: 20}, // inject_receive_time
                        {wch: 20}, // c_time
                        {wch: 20}  // m_time
                    ];
            }
        };

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(processedList);
        ws['!cols'] = getColumnWidths(api_type);

        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, fileName);

        status.textContent = `导出成功！共 ${list.length} 条数据${data.total > filters.limit ? '（总共 ' + data.total + ' 条）' : ''}`;
    } catch (error) {
        status.textContent = '导出失败：' + error.message;
        console.error('导出失败:', error);
    }
}); 

// 在 try 块内修改 API 地址判断部分
const getApiUrl = (type) => {
    console.log('Getting URL for type:', type);
    switch (type) {
        case 'inject_cover':
            return 'http://cms.enjoy-tv.cn/zinject/inject/cover/list';
        case 'inject_video':
            return 'http://cms.enjoy-tv.cn/zinject/inject/video/list';
        case 'project_cover':
            return 'http://cms.enjoy-tv.cn/query/project/cover/list';
        case 'project_video':
            return 'http://cms.enjoy-tv.cn/query/project/video/list';
        case 'total_cover':
            return 'http://cms.enjoy-tv.cn/query/cover/list';
        default:
            throw new Error('未知的数据来源类型');
    }
};

// 获取文件名前缀
const getFileNamePrefix = (type) => {
    switch (type) {
        case 'inject_cover':
            return '注入库剧头数据';
        case 'inject_video':
            return '注入库子集数据';
        case 'project_cover':
            return '项目库专辑数据';
        case 'project_video':
            return '项目库子集数据';
        case 'total_cover':
            return '总库专辑数据';
        default:
            return '导出数据';
    }
}; 