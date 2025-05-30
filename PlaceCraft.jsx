/*
 * PlaceCraft.jsx
 * InDesign オブジェクト自動配置スクリプト
 * 
 * このスクリプトは、Adobe InDesign上で指定フォルダ内のオブジェクト（画像、PDFなど）を
 * アートボードに自動配置するツールです。
 * 
 * バージョン: 1.0.0
 * 日付: 2025-05-30
 */

// エラー時にスタックトレースを表示するためのグローバル変数
$.level = 2;

// 定数
var SCRIPT_NAME = "PlaceCraft";
var SCRIPT_VERSION = "1.0.0";
var DEFAULT_OBJECT_LIMIT = 100;
var DEFAULT_GRID_COLUMNS = 3;
var DEFAULT_GRID_MARGIN = 10; // mm
var DEFAULT_RANDOM_ROTATION = 30; // ±30度
var DEFAULT_RANDOM_SCALE_MIN = 80; // 80%
var DEFAULT_RANDOM_SCALE_MAX = 120; // 120%
var SUPPORTED_FILE_TYPES = ["jpg", "jpeg", "png", "tif", "tiff", "pdf", "psd", "ai", "eps"];

// メイン処理
function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。先にドキュメントを開いてください。");
        return;
    }

    var doc = app.activeDocument;
    
    // アートボードが存在するか確認
    if (doc.pages.length === 0) {
        alert("アートボードを作成してください。");
        return;
    }

    // ダイアログを表示
    showDialog(doc);
}

// ダイアログを表示する関数
function showDialog(doc) {
    var dialog = app.dialogs.add({
        name: SCRIPT_NAME + " " + SCRIPT_VERSION,
        canCancel: true
    });

    // ダイアログの構築
    var column = dialog.dialogColumns.add();
    
    // 1. フォルダ選択
    var folderRow = column.dialogRows.add();
    folderRow.staticTexts.add({staticLabel: "📁 フォルダ選択:"});
    
    // ボタン用の列を追加
    var buttonColumn = folderRow.dialogColumns.add();
    var selectGroup = buttonColumn.borderPanels.add();
    var folderSelectButton = selectGroup.enablingGroups.add({staticLabel: "選択"});
    folderSelectButton.minWidth = 80;
    var folderPathText = folderRow.dialogColumns.add().staticTexts.add({
        staticLabel: "フォルダが選択されていません",
        minWidth: 200
    });
    
    // 2. ファイルタイプ選択
    var fileTypeRow = column.dialogRows.add();
    fileTypeRow.staticTexts.add({staticLabel: "📄 ファイルタイプ"});
    
    var fileTypeGroup = column.dialogRows.add().borderPanels.add().dialogColumns.add();
    var allFilesRadio = fileTypeGroup.dialogRows.add().radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "全て 📂", checkedState: true}]
    });
    var specificTypeRow = fileTypeGroup.dialogRows.add();
    var specificTypeRadio = specificTypeRow.radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "特定タイプ 📑", checkedState: false}]
    });
    
    var fileTypeDropdown = specificTypeRow.dropdowns.add({
        stringList: SUPPORTED_FILE_TYPES,
        selectedIndex: 0,
        minWidth: 150
    });
    fileTypeDropdown.enabled = false;
    
    // 3. 配置スタイル
    var placementStyleRow = column.dialogRows.add();
    placementStyleRow.staticTexts.add({staticLabel: "🎲 配置スタイル"});
    
    var styleGroup = column.dialogRows.add().borderPanels.add().dialogColumns.add();
    var gridStyleRow = styleGroup.dialogRows.add();
    var gridRadio = gridStyleRow.radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "グリッド 📐", checkedState: true}]
    });
    
    var gridSettingsRow = styleGroup.dialogRows.add();
    gridSettingsRow.staticTexts.add({staticLabel: "    列数:"});
    var columnsField = gridSettingsRow.integerEditboxes.add({
        editValue: DEFAULT_GRID_COLUMNS,
        minWidth: 50
    });
    gridSettingsRow.staticTexts.add({staticLabel: "余白:"});
    var marginField = gridSettingsRow.measurementEditboxes.add({
        editValue: DEFAULT_GRID_MARGIN,
        editUnits: MeasurementUnits.MILLIMETERS,
        minWidth: 50
    });
    gridSettingsRow.staticTexts.add({staticLabel: "mm"});
    
    var randomStyleRow = styleGroup.dialogRows.add();
    var randomRadio = randomStyleRow.radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "ランダム 🎲", checkedState: false}]
    });
    
    var randomSettingsRow = styleGroup.dialogRows.add();
    randomSettingsRow.staticTexts.add({staticLabel: "    回転:"});
    var rotationField = randomSettingsRow.angleEditboxes.add({
        editValue: DEFAULT_RANDOM_ROTATION,
        minWidth: 50
    });
    randomSettingsRow.staticTexts.add({staticLabel: "°"});
    randomSettingsRow.staticTexts.add({staticLabel: "スケール:"});
    var scaleMinField = randomSettingsRow.percentEditboxes.add({
        editValue: DEFAULT_RANDOM_SCALE_MIN,
        minWidth: 50
    });
    randomSettingsRow.staticTexts.add({staticLabel: "-"});
    var scaleMaxField = randomSettingsRow.percentEditboxes.add({
        editValue: DEFAULT_RANDOM_SCALE_MAX,
        minWidth: 50
    });
    randomSettingsRow.staticTexts.add({staticLabel: "%"});
    randomSettingsRow.enabled = false;
    
    // 4. 配置順
    var orderRow = column.dialogRows.add();
    orderRow.staticTexts.add({staticLabel: "🔄 配置順:"});
    var orderDropdown = orderRow.dropdowns.add({
        stringList: ["ファイル名", "ファイルサイズ", "作成日", "ランダム"],
        selectedIndex: 0,
        minWidth: 150
    });
    
    // 5. 配置サイズ
    var sizeRow = column.dialogRows.add();
    sizeRow.staticTexts.add({staticLabel: "🔲 配置サイズ"});
    
    var sizeGroup = column.dialogRows.add().borderPanels.add().dialogColumns.add();
    var originalSizeRadio = sizeGroup.dialogRows.add().radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "原寸 📏", checkedState: true}]
    });
    var fitRadio = sizeGroup.dialogRows.add().radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "アートボードにフィット 🔲", checkedState: false}]
    });
    
    // 6. 配置オブジェクト数
    var limitRow = column.dialogRows.add();
    limitRow.staticTexts.add({staticLabel: "🔢 配置オブジェクト数:"});
    var limitField = limitRow.integerEditboxes.add({
        editValue: DEFAULT_OBJECT_LIMIT,
        minWidth: 80
    });
    limitRow.staticTexts.add({staticLabel: "(0=無制限)"});
    
    // 7. 配置方法
    var methodRow = column.dialogRows.add();
    methodRow.staticTexts.add({staticLabel: "🔗 配置方法"});
    
    var methodGroup = column.dialogRows.add().borderPanels.add().dialogColumns.add();
    var embedRadio = methodGroup.dialogRows.add().radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "リアル挿入 🖼️", checkedState: true}]
    });
    var linkRadio = methodGroup.dialogRows.add().radiobuttonGroups.add({
        radiobuttonControls: [{staticLabel: "リンク 🔗", checkedState: false}]
    });
    
    // 8. ボタン（プレビュー、キャンセル、OK）
    var buttonRow = column.dialogRows.add();
    var previewGroup = buttonRow.borderPanels.add();
    var previewButton = previewGroup.enablingGroups.add({staticLabel: "プレビュー 👁️"});
    previewButton.minWidth = 100;
    
    // イベントリスナー
    var selectedFolder = null;
    
    // フォルダ選択ボタンのイベント
    folderSelectButton.onClick = function() {
        selectedFolder = Folder.selectDialog("オブジェクトフォルダを選択");
        if (selectedFolder) {
            folderPathText.staticLabel = decodeURI(selectedFolder.fsName);
        }
    };
    
    // ファイルタイプラジオボタンのイベント
    allFilesRadio.onClick = function() {
        fileTypeDropdown.enabled = false;
    };
    
    specificTypeRadio.onClick = function() {
        fileTypeDropdown.enabled = true;
    };
    
    // 配置スタイルラジオボタンのイベント
    gridRadio.onClick = function() {
        gridSettingsRow.enabled = true;
        randomSettingsRow.enabled = false;
    };
    
    randomRadio.onClick = function() {
        gridSettingsRow.enabled = false;
        randomSettingsRow.enabled = true;
    };
    
    // プレビューボタンのイベント
    previewButton.onClick = function() {
        if (!selectedFolder) {
            alert("フォルダを選択してください。");
            return;
        }
        
        // プレビュー処理（簡易実装）
        try {
            alert("プレビュー機能は実装中です。将来のバージョンで利用可能になります。");
        } catch (e) {
            alert("プレビューを生成できませんでした: " + e.message);
        }
    };
    
    // ダイアログを表示
    var result = dialog.show();
    
    // ダイアログの結果を処理
    if (result) {
        // 設定を収集
        var settings = {
            folder: selectedFolder,
            fileType: allFilesRadio.selectedButton === 0 ? "all" : "specific",
            specificType: SUPPORTED_FILE_TYPES[fileTypeDropdown.selectedIndex],
            placementStyle: gridRadio.selectedButton === 0 ? "grid" : "random",
            gridColumns: columnsField.editValue,
            gridMargin: marginField.editValue,
            randomRotation: rotationField.editValue,
            randomScaleMin: scaleMinField.editValue,
            randomScaleMax: scaleMaxField.editValue,
            placementOrder: orderDropdown.selectedIndex,
            sizingOption: originalSizeRadio.selectedButton === 0 ? "original" : "fit",
            objectLimit: limitField.editValue,
            placementMethod: embedRadio.selectedButton === 0 ? "embed" : "link"
        };
        
        // 入力検証
        if (!validateSettings(settings)) {
            return;
        }
        
        // オブジェクト配置処理を実行
        placeObjects(doc, settings);
    }
    
    // ダイアログを破棄
    dialog.destroy();
}

// 設定の検証
function validateSettings(settings) {
    // フォルダ選択の検証
    if (!settings.folder) {
        alert("フォルダを選択してください。");
        return false;
    }
    
    // 特定タイプ選択時の検証
    if (settings.fileType === "specific" && !settings.specificType) {
        alert("少なくとも1つのファイルタイプを選択してください。");
        return false;
    }
    
    // グリッド設定の検証
    if (settings.placementStyle === "grid") {
        if (settings.gridColumns < 1 || settings.gridColumns > 10) {
            alert("列数は1～10の範囲で指定してください。");
            return false;
        }
        
        if (settings.gridMargin < 0 || settings.gridMargin > 100) {
            alert("余白は0～100mmの範囲で指定してください。");
            return false;
        }
    }
    
    // ランダム設定の検証
    if (settings.placementStyle === "random") {
        if (settings.randomRotation < 0 || settings.randomRotation > 180) {
            alert("回転は0～180°の範囲で指定してください。");
            return false;
        }
        
        if (settings.randomScaleMin < 50 || settings.randomScaleMin > 200) {
            alert("スケール最小値は50～200%の範囲で指定してください。");
            return false;
        }
        
        if (settings.randomScaleMax < 50 || settings.randomScaleMax > 200) {
            alert("スケール最大値は50～200%の範囲で指定してください。");
            return false;
        }
        
        if (settings.randomScaleMin > settings.randomScaleMax) {
            alert("スケール最小値は最大値より小さくしてください。");
            return false;
        }
    }
    
    // 配置オブジェクト数の検証
    if (settings.objectLimit < 0) {
        alert("配置オブジェクト数は0以上の整数を入力してください。");
        return false;
    }
    
    return true;
}

// オブジェクト配置処理
function placeObjects(doc, settings) {
    // ファイル一覧を取得
    var files = getFilesFromFolder(settings.folder, settings);
    
    if (files.length === 0) {
        alert("指定されたフォルダには配置可能なファイルがありません。");
        return;
    }
    
    // 配置制限の適用
    if (settings.objectLimit > 0 && files.length > settings.objectLimit) {
        files = files.slice(0, settings.objectLimit);
    }
    
    // プログレスバー表示判断
    var showProgress = files.length > 1000;
    var progressBar = null;
    
    if (showProgress) {
        progressBar = createProgressBar("オブジェクトを配置中...", files.length);
    }
    
    try {
        // 配置処理
        app.doScript(function() {
            var placedCount = 0;
            
            // 現在のアートボード（ページ）情報
            var currentPage = 0;
            var totalPages = doc.pages.length;
            
            // グリッド配置用の変数
            var gridX = 0;
            var gridY = 0;
            var gridItemWidth = 0;
            var gridItemHeight = 0;
            
            if (settings.placementStyle === "grid") {
                // グリッドアイテムのサイズを計算
                calculateGridItemSize(doc, settings);
            }
            
            // 各ファイルを配置
            for (var i = 0; i < files.length; i++) {
                // プログレスバー更新
                if (showProgress) {
                    progressBar.value = i + 1;
                }
                
                // 配置位置を計算
                var placementPosition;
                if (settings.placementStyle === "grid") {
                    placementPosition = calculateGridPosition(doc, settings, i, gridX, gridY);
                    
                    // グリッド位置を更新
                    gridX++;
                    if (gridX >= settings.gridColumns) {
                        gridX = 0;
                        gridY++;
                    }
                } else {
                    placementPosition = calculateRandomPosition(doc, settings);
                }
                
                // オブジェクト配置
                try {
                    var placedItem = placeFile(doc, files[i], placementPosition, settings);
                    if (placedItem) {
                        placedCount++;
                    }
                } catch (e) {
                    // エラーが発生しても次のファイルを処理
                    $.writeln("ファイル配置エラー: " + e.message);
                }
            }
            
            // 配置完了メッセージ
            alert("配置完了: " + placedCount + "個のオブジェクトを配置しました。");
            
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT);
    } catch (e) {
        alert("エラーが発生しました: " + e.message);
    } finally {
        // プログレスバーを閉じる
        if (progressBar) {
            progressBar.close();
        }
    }
}

// フォルダからファイル一覧を取得
function getFilesFromFolder(folder, settings) {
    var files = [];
    
    // フォルダ内のファイルを取得
    var folderFiles = folder.getFiles();
    
    // ファイルをフィルタリング
    for (var i = 0; i < folderFiles.length; i++) {
        var file = folderFiles[i];
        
        // ディレクトリはスキップ
        if (file instanceof Folder) {
            continue;
        }
        
        // ファイル拡張子をチェック
        var extension = getFileExtension(file.name).toLowerCase();
        
        // ファイルタイプフィルタ
        if (settings.fileType === "all") {
            if (isSupportedFileType(extension)) {
                files.push(file);
            }
        } else if (settings.fileType === "specific") {
            if (extension === settings.specificType.toLowerCase()) {
                files.push(file);
            }
        }
    }
    
    // 配置順でソート
    sortFiles(files, settings.placementOrder);
    
    return files;
}

// ファイル拡張子を取得
function getFileExtension(filename) {
    var dotIndex = filename.lastIndexOf(".");
    if (dotIndex !== -1) {
        return filename.substring(dotIndex + 1);
    }
    return "";
}

// サポートされているファイルタイプかチェック
function isSupportedFileType(extension) {
    for (var i = 0; i < SUPPORTED_FILE_TYPES.length; i++) {
        if (extension === SUPPORTED_FILE_TYPES[i].toLowerCase()) {
            return true;
        }
    }
    return false;
}

// ファイルリストをソート
function sortFiles(files, orderIndex) {
    switch (orderIndex) {
        case 0: // ファイル名
            files.sort(function(a, b) {
                return a.name.localeCompare(b.name);
            });
            break;
        case 1: // ファイルサイズ
            files.sort(function(a, b) {
                return a.length - b.length;
            });
            break;
        case 2: // 作成日
            files.sort(function(a, b) {
                return a.created.getTime() - b.created.getTime();
            });
            break;
        case 3: // ランダム
            shuffleArray(files);
            break;
    }
}

// 配列をシャッフル（Fisher-Yates）
function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = array[i];
        array[i] = array[j];
        array[j] = temp;
    }
}

// グリッドアイテムのサイズを計算
function calculateGridItemSize(doc, settings) {
    // 最初のページのサイズを取得
    var page = doc.pages[0];
    var pageWidth = page.bounds[3] - page.bounds[1];
    var pageHeight = page.bounds[2] - page.bounds[0];
    
    // 余白をポイントに変換
    var marginPoints = convertUnits(settings.gridMargin, "mm", "pt");
    
    // グリッドアイテムのサイズを計算
    settings.gridItemWidth = (pageWidth - (marginPoints * (settings.gridColumns + 1))) / settings.gridColumns;
    settings.gridItemHeight = settings.gridItemWidth; // 正方形アイテムを仮定
}

// グリッド配置位置を計算
function calculateGridPosition(doc, settings, index, gridX, gridY) {
    var page = doc.pages[0];
    var pageX = page.bounds[1];
    var pageY = page.bounds[0];
    
    // 余白をポイントに変換
    var marginPoints = convertUnits(settings.gridMargin, "mm", "pt");
    
    // 位置を計算
    var x = pageX + marginPoints + (gridX * (settings.gridItemWidth + marginPoints));
    var y = pageY + marginPoints + (gridY * (settings.gridItemHeight + marginPoints));
    
    return [x, y, x + settings.gridItemWidth, y + settings.gridItemHeight];
}

// ランダム配置位置を計算
function calculateRandomPosition(doc, settings) {
    var page = doc.pages[0];
    var pageWidth = page.bounds[3] - page.bounds[1];
    var pageHeight = page.bounds[2] - page.bounds[0];
    
    // 仮のオブジェクトサイズ（平均的なサイズを仮定）
    var objectWidth = pageWidth / 10;
    var objectHeight = objectWidth;
    
    // ランダム位置を生成（オブジェクトがアートボードからはみ出さないように）
    var x = page.bounds[1] + Math.random() * (pageWidth - objectWidth);
    var y = page.bounds[0] + Math.random() * (pageHeight - objectHeight);
    
    return [x, y, x + objectWidth, y + objectHeight];
}

// ファイルを配置
function placeFile(doc, file, position, settings) {
    // 配置オプションを設定
    var placeOptions = {
        showingOptions: false,
        autoSize: false,
        position: position,
        link: settings.placementMethod === "link"
    };
    
    // ファイルを配置
    var placedItem = doc.pages[0].place(file, position)[0];
    
    // 配置スタイルに応じた処理
    if (settings.placementStyle === "random") {
        // ランダム回転
        var rotation = (Math.random() * 2 - 1) * settings.randomRotation;
        placedItem.rotationAngle = rotation;
        
        // ランダムスケール
        var scale = settings.randomScaleMin + (Math.random() * (settings.randomScaleMax - settings.randomScaleMin));
        placedItem.transform(
            CoordinateSpaces.INNER_COORDINATES,
            AnchorPoint.CENTER_ANCHOR,
            new HomogeneousMatrix(scale / 100, 0, 0, scale / 100, 0, 0)
        );
    }
    
    // サイズオプション適用
    if (settings.sizingOption === "fit") {
        var page = doc.pages[0];
        var pageWidth = page.bounds[3] - page.bounds[1];
        var pageHeight = page.bounds[2] - page.bounds[0];
        
        var itemWidth = placedItem.geometricBounds[3] - placedItem.geometricBounds[1];
        var itemHeight = placedItem.geometricBounds[2] - placedItem.geometricBounds[0];
        
        var widthRatio = pageWidth / itemWidth;
        var heightRatio = pageHeight / itemHeight;
        var scaleRatio = Math.min(widthRatio, heightRatio);
        
        placedItem.transform(
            CoordinateSpaces.INNER_COORDINATES,
            AnchorPoint.CENTER_ANCHOR,
            new HomogeneousMatrix(scaleRatio, 0, 0, scaleRatio, 0, 0)
        );
    }
    
    return placedItem;
}

// プログレスバーを作成
function createProgressBar(title, max) {
    var progressBar = new Window("palette", title, undefined, {closeButton: false});
    progressBar.orientation = "column";
    progressBar.alignChildren = ["center", "center"];
    progressBar.spacing = 10;
    progressBar.margins = 16;
    
    progressBar.staticText = progressBar.add("statictext", undefined, "処理中... 0/" + max);
    progressBar.staticText.alignment = "center";
    
    progressBar.progressBar = progressBar.add("progressbar", undefined, 0, max);
    progressBar.progressBar.preferredSize.width = 300;
    
    // プログレスバーの値を更新するためのプロパティとメソッド
    progressBar.__defineGetter__("value", function() {
        return this.progressBar.value;
    });
    
    progressBar.__defineSetter__("value", function(val) {
        this.progressBar.value = val;
        this.staticText.text = "処理中... " + val + "/" + max;
        this.update();
    });
    
    progressBar.show();
    return progressBar;
}

// 単位変換
function convertUnits(value, fromUnit, toUnit) {
    var unitMap = {
        "pt": 1,
        "mm": 2.83465,
        "in": 72,
        "cm": 28.3465
    };
    
    return value * unitMap[fromUnit] / unitMap[toUnit];
}

// スクリプト実行
main();
