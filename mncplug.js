const hlColor = 65535;
const noColor = 16777215;
const largeFileRowCount = 10000;
const cnTextCol = 6;
const projectSheetName = "new_translate";
const obsolateCol = 5;
function updateHighlight() {

    function getNewShtArr(shtPath, loadColor = true) {
        // 获取返稿文本和颜色信息
        if (shtPath === "") {
            MsgBox("未输入文件路径，已退出程序。");
            return;
        }

        let newWorkbook;
        try {
            newWorkbook = Workbooks.Open(shtPath);
        }
        catch (error) {
            MsgBox("无法打开工作簿，请确认选择了正确的表格：" + error.message);
            return;
        }

        let newShtName = InputBox("输入使用的页签名称或顺序编号。", "选择页签", `${newWorkbook.ActiveSheet.Name}`);
        let newSht;
        try {
            if (!isNaN(Number(newShtName))) {
                newShtName = +newShtName;
            }
            newSht = newWorkbook.Sheets.Item(newShtName);
            if (!newSht) {
                throw Error("未找到更新页签！")
            }
        }
        catch (error) {
            MsgBox("无法访问新工作簿中的工作表：" + error.message);
            newWorkbook.Close(false);  // 关闭新工作簿，不保存更改
            return [false, false, false];
        }
        const getInfo = MsgBox(`确定从返稿的页签 ${newSht.Name} 中获取信息吗？`, jsOKCancel);
        if (getInfo === 2) {
            newWorkbook.Close(false);
            return [false, false, false];
        }
        const _newArr = newSht.UsedRange.Value2;
        let _langCol = 0;
        if (_newArr.length > largeFileRowCount) {
            _langCol = InputBox("由于返稿文本量巨大，为避免卡死，请先输入一列进行更新 (如C列输入数字3)", "输入列号", "3");
            if (_langCol === "") {
                newWorkbook.Close(false);
                return [false, false, false];
            }
            _langCol = Number(_langCol);
        }

        const _colorDict = {};
        if (loadColor) {
            const lr = newSht.UsedRange.Item(newSht.UsedRange.Count).Row;
            // const lc = newSht.UsedRange.Item(newSht.UsedRange.Count).Column;
            if (_langCol < cnTextCol + 1) {
                for (let r = 1; r <= lr; r++) {
                    for (let c = 1; c <= _langCol; c++) {
                        _colorDict[`${r - 1}:${c - 1}`] = newSht.UsedRange.Item(r, c).Interior.Color;
                    }
                }
            }
            else {
                for (let r = 1; r <= lr; r++) {
                    _colorDict[`${r - 1}:${_langCol - 1}`] = newSht.UsedRange.Item(r, _langCol).Interior.Color;
                }
            }
        }

        newWorkbook.Close(false);
        return [_newArr, _colorDict, _langCol];
    }

    // 获取原表Key, 有key的情况取key，无key的情况用中文作为索引，对应列为表格行
    // 1. 总表有key，返稿有key (Key匹配)
    // 2. 总表有key，返稿无key (中文匹配)
    // 3. 总表无key，返稿无key (中文匹配)
    // 4. 总表无key，返稿有key (异常排除)

    function getKeyDict(shtArr) {
        // 记录所有中文对应的行，记录所有key对应的行
        _dict = {};
        _cnDict = {};
        for (let i = 0; i < shtArr.length; i++) {
            let _cn = shtArr[i][cnTextCol];
            if (_cn) {
                if (_cn in _cnDict) {
                    _cnDict[_cn].push(i + 1);
                }
                else {
                    _cnDict[_cn] = [i + 1];
                }
            }

            let _key = `${shtArr[i][0]}_${shtArr[i][1]}_${shtArr[i][2]}`;
            if (_key) {
                if (_key in _dict) {
                    if (shtArr[i][obsolateCol - 1] !== "已废弃") {
                        MsgBox(`总表发现重复Key: ${_key}，请检查总表`);
                        return [false, false];
                    }
                }
                else {
                    _dict[_key] = i + 1;
                }
            }
        }
        return [_dict, _cnDict];
    }

    // 对单列进行更新
    // 精简为两种情况: 返稿有key则key匹配(1), 返稿无key则中文匹配空行(2,3)
    function updateOneColumn(uArr, uNewArr, uCol, uKeyDict, uCnDict, uColorDict, uIsHighlight, uSht) {
        if (uIsHighlight === true) {
            for (let row = 1; row < uNewArr.length; row++) {
                const newKey = `${uNewArr[row][0]}_${uNewArr[row][1]}_${uNewArr[row][2]}`;
                const newCn = uNewArr[row][cnTextCol];

                const shtRow = uKeyDict[newKey];
                const cnRows = uCnDict[newCn];
                if (shtRow) {
                    if (newColorDict[`${row}:${uCol}`] !== noColor) {
                        const _newVal = uNewArr[row][uCol];
                        if (uArr[shtRow - 1][uCol] !== _newVal) {
                            uSht.Cells.Item(shtRow, uCol + 1).Value2 = _newVal;
                        }
                    }
                }
                else if (cnRows) {
                    if (newColorDict[`${row}:${uCol}`] !== noColor) {
                        for (const cnR of cnRows) {
                            const _oriVal = uArr[cnR - 1][uCol];
                            const _newVal = uNewArr[row][uCol];
                            if (!_oriVal) {
                                uSht.Cells.Item(cnR, uCol + 1).Value2 = _newVal;
                            }
                        }
                    }
                }
            }
        }
        else {
            for (let row = 1; row < uNewArr.length; row++) {
                const newKey = `${uNewArr[row][0]}_${uNewArr[row][1]}_${uNewArr[row][2]}`;
                const newCn = uNewArr[row][cnTextCol];

                const shtRow = uKeyDict[newKey];
                const cnRows = uCnDict[newCn];
                if (shtRow) {
                    const _newVal = uNewArr[row][uCol];
                    if (uArr[shtRow - 1][uCol] !== _newVal) {
                        uSht.Cells.Item(shtRow, uCol + 1).Value2 = _newVal;
                    }
                }
                else if (cnRows) {
                    for (const cnR of cnRows) {
                        const _oriVal = uArr[cnR - 1][uCol];
                        const _newVal = uNewArr[row][uCol];
                        if (!_oriVal) {
                            uSht.Cells.Item(cnR, uCol + 1).Value2 = _newVal;
                        }
                    }
                }
            }
        }
    }

    // 总表处理部分
    const sht = Application.ActiveSheet;
    if (sht.Name !== projectSheetName) {
        MsgBox(`页签名不为${projectSheetName}，请确认选择了正确的总表！`);
        return;
    }
    const arr = sht.UsedRange.Value2;
    const [keyDict, cnDict] = getKeyDict(arr);
    if (keyDict === false) {
        return;
    }

    if (arr[0].length < cnTextCol + 1) {
        MsgBox("总表列数太少，请确认选择了正确的总表！")
        return;
    }

    // 返稿处理部分
    const fDialog = Application.FileDialog(msoFileDialogFilePicker);
    fDialog.Filters.Clear();
    fDialog.Filters.Add("表格文件", "*.xls*")
    fDialog.Title = "选择返稿文件"
    let accept = fDialog.Show();
    if (accept === 0)
        return;
    const newShtPath = fDialog.SelectedItems.Item(1);

    // 高亮处理部分
    let highlightChoice = MsgBox("是否只从高亮的单元格更新", jsYesNo);
    const isHighlightOnly = highlightChoice === 6 ? true : false

    Application.ScreenUpdating = false;
    const [newArr, newColorDict, langCol] = getNewShtArr(newShtPath, isHighlightOnly);

    if (!newArr) {
        return;
    }
    if (newArr.length < 2) {
        MsgBox("更新表需未包含更新内容，程序已退出")
        return;
    }
    if (arr[0].length !== newArr[0].length) {
        let willUpdate = MsgBox("总表和更新表的列数不一致，确定要更新吗？", jsOKCancel);
        if (willUpdate == 2)
            return;
    }


    //更新逻辑
    if (langCol < 3) {
        for (let col = 2; col < newArr[0].length; col++) {
            updateOneColumn(arr, newArr, col, keyDict, cnDict, newColorDict, isHighlightOnly, sht)
        }
    }
    else {
        updateOneColumn(arr, newArr, langCol - 1, keyDict, cnDict, newColorDict, isHighlightOnly, sht)
    }

    Application.ScreenUpdating = true;
    MsgBox("更新完成！");
}
