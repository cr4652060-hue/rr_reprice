(function () {
    const $ = (id) => document.getElementById(id);

    const valOrPlaceholder = (id, placeholder) => {
        const el = $(id);
        if (!el) return placeholder;
        const v = (el.value ?? "").toString().trim();
        return v === "" ? placeholder : v;
    };

    function exportToExcel() {
        const loanProduct = document.getElementById('product').value || '{loanProduct}';
        const guarantee = document.getElementById('guarantee').value || '{guarantee}';
        const amount = document.getElementById('amount').value || '{amount}';
        const term = document.getElementById('term').value || '{term}';
        const origRate = document.getElementById('origRate').value || '{origRate}';
        const execRate = document.getElementById('execRate').value || '{execRate}';

        // 假设最低制度要求的利率是来自 RATE_TABLE 中的无定价利率
        const minRate = 3.45; // 示例数据，实际使用时根据 RATE_TABLE 计算
        const finalRate = 6.00; // 最终建议利率，来自计算

        // 三条件是否需要判断
        const needCond = (amount >= 100 && term <= 36) ? "是" : "否";  // 示例条件判断

        // 创建 Excel 数据，符合模板格式
        const ws_data = [
            ['贷款产品', loanProduct],
            ['担保方式', guarantee],
            ['额度（万元）', amount],
            ['期限（月）', term],
            ['原合同/当前利率', origRate],
            ['拟执行利率', execRate],
            ['最低利率', minRate], // 填入最低利率
            ['最终建议执行利率', finalRate], // 填入最终建议执行利率
            ['是否需上报', '{needReport}'],  // 根据需要的逻辑填充
            ['三条件是否需要判断', needCond],  // 填充三条件判断
            ['近一年逾期次数', '{overdue}'], // 填入占位符
            ['资金归行率', '{flow}'],
            ['互联网负面信息', '{negative}'],
            ['三条件是否满足', '{condOk}'],
            // 公式计算（根据模板需求）：以下是填充占位符的公式
            ['首次判定', 'IF(AND(C5>=F5, C6<=C5), "不再上报定价审核", "需要继续写下项")'],
            ['二次判定', 'IF(AND(C7>=F7, C8<=C7), "不再上报定价审核", "需要继续写下项")'],
        ];

        // 生成 Excel 工作表
        const ws = XLSX.utils.aoa_to_sheet([
            ['贷款产品', '担保方式', '最小金额', '最大金额', '最小期限', '最大期限', '最低利率', '最高利率'],
            ...ws_data,
        ]);

        // 设置列宽（使其看起来更加整齐）
        ws["!cols"] = [
            { wch: 16 },
            { wch: 20 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 },
            { wch: 12 }
        ];

        // 创建工作簿并生成文件
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '贷款利率数据');

        // 文件名使用时间戳
        const now = new Date();
        const filename = "贷款利率数据_" + now.getFullYear() + (now.getMonth() + 1).toString().padStart(2, "0") + now.getDate().toString().padStart(2, "0") + ".xlsx";

        // 导出文件
        XLSX.writeFile(wb, filename);
    }

    // 绑定导出按钮
    document.getElementById('btnExportExcel').addEventListener('click', exportToExcel);

    fetch('src/main/resources/re_price_template.xlsx')
        .then(response => response.blob())
        .then(blob => {
            const reader = new FileReader();
            reader.onload = function() {
                const data = new Uint8Array(reader.result);
                const wb = XLSX.read(data, { type: 'array' });
                // 在这里填充数据到模板
            };
            reader.readAsArrayBuffer(blob);
        });

    window.exportToExcel = exportToExcel;
})();
