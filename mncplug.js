const hlColor = 65535;
const noColor = 16777215;
const largeFileRowCount = 10000;
const cnVTextCol = 6;
const projectSheetName = "new_translate";
const obsolateVCol = 5;
function updateHighlight() {

    function getNewShtArr(shtPath, loadColor = true) {
        // 返回值_newArr全表内容, _colorDict单元格颜色信息, _langCol需要更新的列
        // _langVCol为0时默认全部更新
        // 获取返稿文本和颜色信息
        const EXP_MSG = [false, false, false]
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
            return EXP_MSG;
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
            return EXP_MSG;
        }
        const getInfo = MsgBox(`确定从返稿的页签 ${newSht.Name} 中获取信息吗？`, jsOKCancel);
        if (getInfo === 2) {
            newWorkbook.Close(false);
            return EXP_MSG;
        }
        const _newArr = newSht.UsedRange.Value2;
        let _langVCol = 0;
        if (_newArr.length > largeFileRowCount) {
            _langVCol = InputBox("由于返稿文本量巨大，为避免卡死，请先输入一列进行更新 (如C列输入数字3)", "输入列号", "7");
            if (_langVCol === "") {
                newWorkbook.Close(false);
                return EXP_MSG;
            }
            _langVCol = Number(_langVCol);
        }

        const _colorDict = {};
        if (loadColor) {
            const lr = newSht.UsedRange.Item(newSht.UsedRange.Count).Row;
            const lc = newSht.UsedRange.Item(newSht.UsedRange.Count).Column;
            if (_langVCol === 0) {
                for (let r = 1; r <= lr; r++) {
                    for (let c = 1; c <= lc; c++) {
                        _colorDict[`${r - 1}:${c - 1}`] = newSht.UsedRange.Item(r, c).Interior.Color;
                    }
                }
            }
            else if (_langVCol < cnVTextCol + 1) {
                MsgBox(`输入的列数过小，列${_langVCol}`)
                return EXP_MSG
            }
            else {
                for (let r = 1; r <= lr; r++) {
                    _colorDict[`${r - 1}:${_langVCol - 1}`] = newSht.UsedRange.Item(r, _langVCol).Interior.Color;
                }
            }
        }

        newWorkbook.Close(false);
        return [_newArr, _colorDict, _langVCol];
    }

    // 获取原表Key, 有key的情况取key，无key的情况用中文作为索引，对应列为表格行
    // 1. 总表有key，返稿有key (Key匹配)
    // 2. 总表有key，返稿无key (中文匹配)
    // 3. 总表无key，返稿无key (中文匹配)
    // 4. 总表无key，返稿有key (异常排除)

    function getKeyDict(shtArr) {
        // 记录所有中文对应的行，记录所有key对应的行
        // VDict, cnVDict记录视觉上的行号
        _VDict = {};
        _cnVDict = {};
        for (let i = 0; i < shtArr.length; i++) {
            let _cn = shtArr[i][cnVTextCol - 1];
            if (_cn) {
                if (_cn in _cnVDict) {
                    _cnVDict[_cn].push(i + 1);
                }
                else {
                    _cnVDict[_cn] = [i + 1];
                }
            }

            let _key = `${shtArr[i][0]}_${shtArr[i][1]}_${shtArr[i][2]}`;
            let _empty_key = 'undefined_undefined_undefined'
            if (_key) {
                if (_key === _empty_key) {
                    continue;
                }
                else if (_key in _VDict) {
                    if (shtArr[i][obsolateVCol - 1] !== "已废弃") {
                        MsgBox(`总表发现重复Key: ${_key}，请检查总表`);
                        return [false, false];
                    }
                }
                else {
                    _VDict[_key] = i + 1;
                }
            }
        }
        return [_VDict, _cnVDict];
    }


    function updateOneColumn(uArr, uNewArr, uCol, uKeyVDict, uCnVDict, uColorDict, uIsHighlight, uSht) {
        // 对单列进行更新
        // 精简为两种情况: 返稿有key则key匹配(1), 返稿无key则中文匹配空行(2,3)
        // uCol传入的是数组中的列号，而非视觉上的列号
        if (uIsHighlight === true) {
            for (let row = 1; row < uNewArr.length; row++) {
                const newKey = `${uNewArr[row][0]}_${uNewArr[row][1]}_${uNewArr[row][2]}`;
                const newCn = uNewArr[row][cnVTextCol - 1];

                const shtVRow = uKeyVDict[newKey];
                const cnVRows = uCnVDict[newCn];
                if (shtVRow) {
                    if (uColorDict[`${row}:${uCol}`] !== noColor) {
                        const _newVal = uNewArr[row][uCol];
                        if (uArr[shtVRow - 1][uCol] !== _newVal) {
                            uSht.Cells.Item(shtVRow, uCol + 1).Value2 = _newVal;
                        }
                    }
                }
                else if (cnVRows) {
                    if (uColorDict[`${row}:${uCol}`] !== noColor) {
                        for (const cnVR of cnVRows) {
                            const _oriVal = uArr[cnVR - 1][uCol];
                            const _newVal = uNewArr[row][uCol];
                            if (!_oriVal) {
                                uSht.Cells.Item(cnVR, uCol + 1).Value2 = _newVal;
                            }
                        }
                    }
                }
            }
        }
        else {
            for (let row = 1; row < uNewArr.length; row++) {
                const newKey = `${uNewArr[row][0]}_${uNewArr[row][1]}_${uNewArr[row][2]}`;
                const newCn = uNewArr[row][cnVTextCol - 1];

                const shtVRow = uKeyVDict[newKey];
                const cnVRows = uCnVDict[newCn];
                if (shtVRow) {
                    const _newVal = uNewArr[row][uCol];
                    if (uArr[shtVRow - 1][uCol] !== _newVal) {
                        uSht.Cells.Item(shtVRow, uCol + 1).Value2 = _newVal;
                    }
                }
                else if (cnVRows) {
                    for (const cnVR of cnVRows) {
                        const _oriVal = uArr[cnVR - 1][uCol];
                        const _newVal = uNewArr[row][uCol];
                        if (!_oriVal) {
                            uSht.Cells.Item(cnVR, uCol + 1).Value2 = _newVal;
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
    const [keyVDict, cnVDict] = getKeyDict(arr);
    if (keyVDict === false) {
        return;
    }

    if (arr[0].length < cnVTextCol + 1) {
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
    const [newArr, newColorDict, langVCol] = getNewShtArr(newShtPath, isHighlightOnly);

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
    if (langVCol === 0) {
        let enVTextCol = cnVTextCol + 1;
        for (let col = enVTextCol - 1; col < newArr[0].length; col++) {
            updateOneColumn(arr, newArr, col, keyVDict, cnVDict, newColorDict, isHighlightOnly, sht)
        }
    }
    else {
        updateOneColumn(arr, newArr, langVCol - 1, keyVDict, cnVDict, newColorDict, isHighlightOnly, sht)
    }

    Application.ScreenUpdating = true;
    MsgBox("更新完成！");
}
