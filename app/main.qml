import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Window
import QtMultimedia
import "controls"
import "components"

Window {
    id: rootWindow
    visible: true
    width: 400
    height: 500
    maximumHeight: 600
    minimumWidth: 400
    minimumHeight: 500
    title: "CubeFlow"
    color: "#2b2b36"
    property color themeBg: "#2b2b36"
    property color themePanel: "#3d3d4d"
    property color themeLayer1: "#52525e"
    property color themeLayer2: "#6b6b7a"
    property color themeLayer3: "#7d7d8a"
    property color themeText: "white"
    property color themeTextSecondary: "#b8b8c4"
    property color themeInset: "#2b2b36"
    flags: Qt.Window | Qt.CustomizeWindowHint | Qt.WindowTitleHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint
    onProcessStateChanged: {
        requestActivate()
        if (processState !== "batchReview") {
            batchFileNameDrafts = ({})
            batchRowHighlightIndex = -1
        }
        if (processState === "converting") {
            Qt.callLater(function() {
                anchorConvertingStatusToRecent()
            })
        }
        if (processState !== "formatCreate") {
            formatRowsHighlightIndex = -1
            formatEditorFocusType = ""
            formatRowsJumpToIndexNextChange = -1
            formatCompactLabelsMode = false
        }
        if (processState !== "formatDesigner" && processState !== "formatCreate") {
            rootWindow.formatDesignerSelectedFormatIndex = 0
            rootWindow.formatDesignerSelectedRowIndex = -1
        }
    }
    onFormatCompactLabelsModeChanged: {
        if (!formatCompactLabelsMode) {
            clearFormatCreateEditorFocus()
        }
    }

    property string processState: "idle"
    property string selectedFile: ""
    property string selectionType: ""
    property int progress: 0
    property string fileSize: ""
    property var selectedFiles: []
    property var batchFileStatuses: []
    property bool isBatch: false
    property int currentBatchIndex: 0
    property int totalBatchFiles: 0
    property int currentBatchSaveCount: 0
    property string currentFileName: ""
    property var batchOutputs: []
    property var batchFileNameDrafts: ({})
    property int formatDesignerSelectedFormatIndex: 0
    property int formatDesignerSelectedRowIndex: -1
    property string formatEditorFocusType: ""
    property bool formatCompactLabelsMode: false
    property real primaryControlHeight: 38 * scaleFactor
    property real primaryControlFontSize: 12 * scaleFactor
    property bool compactBatchControls: width <= 540
    property real batchControlHeight: (compactBatchControls ? 30 : 32) * scaleFactor
    property real batchOutputTextSize: (compactBatchControls ? 9 : 10) * scaleFactor
    property real formatRowsSavedContentY: 0
    property real batchOutputsSavedContentY: 0
    property bool formatRowsRestorePending: false
    property bool batchOutputsRestorePending: false
    property bool formatRowsScrollToBottomNextChange: false
    property int formatRowsJumpToIndexNextChange: -1
    property int formatRowsHighlightIndex: -1
    property int batchRowHighlightIndex: -1
    property string confirmAction: ""
    property string confirmTitle: "Confirm"
    property string confirmMessage: ""
    property int confirmRequestToken: -1
    property int pendingDeleteFormatIndex: -1
    property string infoMessage: ""
    property string completionDetailMessage: ""
    property bool previewSelectMode: false
    property int previewSelectFormatIndex: -1
    property int previewSelectRowIndex: -1
    property var backendSafe: (backend !== null && backend !== undefined) ? backend : backendNull
    property bool formatDragGhostVisible: false
    property real formatDragGhostY: 0
    property string formatDragGhostText: ""

    QtObject {
        id: backendNull
        property var formatModel: []
        property var xmlTypeOptions: []
        property var customLabelOptions: []
        property string formatDesignerStatus: ""
        property var xmlPreviewHeaders: []
        property var xmlPreviewRows: []
        property string xmlPreviewStatus: ""

        function validateOutputDirectory(_path) { return "" }
        function estimateBatchOutputConflicts(_outputs) { return [] }
        function getFileSize(_path) { return "" }
        function isBuiltinFormat(_name) { return false }

        function setFormatRowFromPreview(_a, _b, _c) { return -1 }
        function updateFormatRow(_a, _b, _c, _d) { return -1 }
        function createFormatDraft() { return -1 }
        function addFormatRow(_i) { return -1 }
        function moveFormatRow(_a, _b, _c) { return -1 }

        function setSelectionType(_t) {}
        function selectFile() {}
        function setDroppedPaths(_p) {}
        function openFormatDesigner() {}
        function removeSelectedFile(_i) {}
        function confirmAndConvert() {}
        function openFormatForEdit(_i) {}
        function duplicateFormatAndOpen(_i) {}
        function deleteFormatDefinition(_i) {}
        function importFormatModelFromFile() {}
        function closeFormatDesigner() {}
        function renameFormatDefinition(_i, _n) {}
        function selectPreviewXmlFile() {}
        function selectAnotherPreviewXmlFile() {}
        function deleteFormatRow(_a, _b) {}
        function saveFormatByName(_i) {}
        function commitFormatEdit() {}
        function cancelFormatEdit() {}
        function confirmDiscardFormatEdit() { return false }
        function resolveInAppConfirm(_token, _accepted) {}
        function applyBatchOutputDirectoryToAll(_p) {}
        function browseBatchOutputDirectoryForAll() {}
        function updateBatchOutputFileName(_i, _n) {}
        function updateBatchOutputDirectory(_i, _p) {}
        function browseBatchOutputDirectory(_i) {}
        function saveAllBatchOutputs() {}
        function convertAnotherFile() {}
        function cancelCurrentOperation() {}
        function selectDifferentFile() {}
    }

    onFormatDesignerSelectedFormatIndexChanged: {
        formatRowsSavedContentY = 0
        formatRowsRestorePending = false
        formatRowsHighlightIndex = -1
    }

    onBatchOutputsChanged: {
        var keepY = batchOutputsSavedContentY
        queueRestoreListContentY(batchOutputsListView, keepY)
    }

    onBatchFileStatusesChanged: {
        if (processState === "converting") {
            Qt.callLater(function() {
                anchorConvertingStatusToRecent()
            })
        }
    }

    property real scaleFactor: Math.min(width / 400, height / 500)
    property bool windowSizeChanged: width > minimumWidth || height > minimumHeight
    property real fileInfoBaseHeight: (processState === "selecting" && isBatch && selectedFiles.length <= 2)
                                    ? 108 * scaleFactor
                                    : 58 * scaleFactor
    property int nonScrollableBatchFiles: 4
    property real fileInfoBatchScrollCapHeight: 230 * scaleFactor
    property real fileInfoMaxHeight: processState === "selecting"
                                   ? (isBatch
                                      ? (selectedFiles.length <= nonScrollableBatchFiles
                                         ? Math.max(fileInfoBaseHeight, fileInfoDesiredHeight)
                                         : fileInfoBatchScrollCapHeight)
                                      : 132 * scaleFactor)
                                   : 190
    property real fileInfoDesiredHeight: (isBatch ? batchFileListColumn.implicitHeight : singleFileInfoText.implicitHeight) + (16 * scaleFactor)
    property real fileInfoCardHeight: Math.min(fileInfoMaxHeight, Math.max(fileInfoBaseHeight, fileInfoDesiredHeight))
    property real compactTriggerPressure: 0.52
    property real fileInfoPressure: processState === "selecting"
                                  ? Math.max(0, Math.min(1, (fileInfoDesiredHeight - fileInfoBaseHeight) / (200 * scaleFactor)))
                                  : 0
    property real headerCompress: processState === "selecting"
                                 ? Math.max(
                                     Math.max(0, Math.min(1, (fileInfoCardHeight - (72 * scaleFactor)) / (72 * scaleFactor))),
                                     fileInfoPressure
                                   )
                                 : 0
    property real headerScale: Math.max(0.72, 1.0 - (0.35 * headerCompress))
    property bool compactHeaderMode: windowSizeChanged
                                  || (processState === "selecting" && fileInfoPressure >= compactTriggerPressure)
    property real compactGifSize: Math.max(130, 130 * scaleFactor)
    property real headerReservedTopSpace: (compactHeaderMode && processState !== "converting" && processState !== "formatDesigner" && processState !== "formatCreate") ? (compactGifSize + (20 * scaleFactor)) : 0

    function baseName(path) {
        var normalized = String(path).replace(/\\/g, "/")
        var parts = normalized.split("/")
        return parts.length ? parts[parts.length - 1] : normalized
    }

    function stripXlsx(name) {
        var value = String(name || "")
        return value.replace(/\.xlsx$/i, "")
    }

    function validateBatchBaseName(name) {
        var value = String(name || "").trim()
        if (value.length === 0) {
            return "File name cannot be empty."
        }
        if (/[<>:\"\/\\|?*\x00-\x1F]/.test(value)) {
            return "Use letters, numbers, spaces, '-', '_', '(', ')', '.'.\nNot allowed: < > : \" / \\ | ? *"
        }
        if (/[. ]$/.test(value)) {
            return "File name cannot end with a dot or space."
        }
        return ""
    }

    function validateBatchSaveDir(path) {
        return backendSafe.validateOutputDirectory(String(path || ""))
    }

    function restoreListContentY(listView, targetY) {
        if (!listView) {
            return
        }
        var maxY = Math.max(0, listView.contentHeight - listView.height)
        var clamped = Math.max(0, Math.min(targetY, maxY))
        listView.contentY = clamped
    }

    function queueRestoreListContentY(listView, targetY) {
        if (listView === formatRowsListView) {
            formatRowsRestorePending = true
        } else if (listView === batchOutputsListView) {
            batchOutputsRestorePending = true
        }
        Qt.callLater(function() {
            restoreListContentY(listView, targetY)
            Qt.callLater(function() {
                restoreListContentY(listView, targetY)
                if (listView === formatRowsListView) {
                    formatRowsRestorePending = false
                    formatRowsSavedContentY = formatRowsListView ? formatRowsListView.contentY : targetY
                } else if (listView === batchOutputsListView) {
                    batchOutputsRestorePending = false
                    batchOutputsSavedContentY = batchOutputsListView ? batchOutputsListView.contentY : targetY
                }
            })
        })
    }

    function rememberFormatRowsScroll() {
        if (formatRowsListView) {
            formatRowsSavedContentY = formatRowsListView.contentY
        }
    }

    function prepareFormatRowsModelChange(scrollToBottom) {
        var toBottom = !!scrollToBottom
        rememberFormatRowsScroll()
        formatRowsRestorePending = true
        formatRowsScrollToBottomNextChange = toBottom
        formatRowsJumpToIndexNextChange = -1
        if (toBottom) {
            formatRowsSavedContentY = 1000000000
        }
    }

    function markFormatRowHighlight(rowIndex) {
        formatRowsHighlightIndex = rowIndex >= 0 ? rowIndex : -1
    }

    function markAndJumpFormatRow(rowIndex) {
        markFormatRowHighlight(rowIndex)
        if (rowIndex >= 0) {
            formatRowsRestorePending = true
            formatRowsJumpToIndexNextChange = rowIndex
        }
    }

    function showColumnLettersOnlyPopup(targetItem) {
        if (targetItem) {
            var p = targetItem.mapToItem(rootWindow.contentItem, 0, targetItem.height + (4 * scaleFactor))
            var margin = 8 * scaleFactor
            var maxX = Math.max(margin, rootWindow.width - columnLettersPopup.width - margin)
            var maxY = Math.max(margin, rootWindow.height - columnLettersPopup.height - margin)
            columnLettersPopup.x = Math.max(margin, Math.min(p.x, maxX))
            columnLettersPopup.y = Math.max(margin, Math.min(p.y, maxY))
        }
        columnLettersPopup.close()
        columnLettersPopup.open()
        columnLettersPopupTimer.restart()
    }

    function showWidthNumbersOnlyPopup(targetItem) {
        if (targetItem) {
            var p = targetItem.mapToItem(rootWindow.contentItem, 0, targetItem.height + (4 * scaleFactor))
            var margin = 8 * scaleFactor
            var maxX = Math.max(margin, rootWindow.width - widthNumbersPopup.width - margin)
            var maxY = Math.max(margin, rootWindow.height - widthNumbersPopup.height - margin)
            widthNumbersPopup.x = Math.max(margin, Math.min(p.x, maxX))
            widthNumbersPopup.y = Math.max(margin, Math.min(p.y, maxY))
        }
        widthNumbersPopup.close()
        widthNumbersPopup.open()
        widthNumbersPopupTimer.restart()
    }

    function clearActiveEditorFocus() {
        var focusedItem = activeFocusItem
        if (focusedItem && focusedItem !== rootWindow && focusedItem.focus !== undefined) {
            focusedItem.focus = false
        }
        if (contentItem) {
            contentItem.forceActiveFocus(Qt.ShortcutFocusReason)
        } else {
            forceActiveFocus(Qt.ShortcutFocusReason)
        }
    }

    function applyPreviewColumnToActiveRow(columnIndex) {
        var targetRow = formatDesignerSelectedRowIndex >= 0
                      ? formatDesignerSelectedRowIndex
                      : formatRowsHighlightIndex
        if (formatDesignerSelectedFormatIndex < 0) {
            return
        }
        if (targetRow < 0
                && backendSafe.formatModel
                && formatDesignerSelectedFormatIndex < backendSafe.formatModel.length
                && backendSafe.formatModel[formatDesignerSelectedFormatIndex].columns
                && backendSafe.formatModel[formatDesignerSelectedFormatIndex].columns.length > 0) {
            targetRow = 0
        }
        if (targetRow < 0) {
            return
        }
        var updatedIndex = backendSafe.setFormatRowFromPreview(formatDesignerSelectedFormatIndex, targetRow, columnIndex)
        if (updatedIndex >= 0) {
            markAndJumpFormatRow(updatedIndex)
            formatDesignerSelectedRowIndex = updatedIndex
            formatEditorFocusType = "value"
        }
    }

    function openPreviewDialogReadOnly() {
        previewSelectMode = false
        previewSelectFormatIndex = -1
        previewSelectRowIndex = -1
        if (backendSafe.xmlPreviewHeaders && backendSafe.xmlPreviewHeaders.length > 0) {
            xmlPreviewDialog.open()
        }
    }

    function openPreviewDialogForIndex(formatIndex, rowIndex) {
        previewSelectMode = true
        previewSelectFormatIndex = formatIndex
        previewSelectRowIndex = rowIndex
        if (!backendSafe.xmlPreviewHeaders || backendSafe.xmlPreviewHeaders.length === 0) {
            backendSafe.selectPreviewXmlFile()
        }
        if (backendSafe.xmlPreviewHeaders && backendSafe.xmlPreviewHeaders.length > 0) {
            xmlPreviewDialog.open()
        }
    }

    function pickSourceIndexFromPreview(columnIndex) {
        if (!previewSelectMode || previewSelectFormatIndex < 0 || previewSelectRowIndex < 0) {
            return
        }
        prepareFormatRowsModelChange(false)
        var updatedIndex = backendSafe.updateFormatRow(previewSelectFormatIndex, previewSelectRowIndex, "value", String(columnIndex))
        markAndJumpFormatRow(updatedIndex)
        formatDesignerSelectedRowIndex = updatedIndex
        formatEditorFocusType = "value"
        previewSelectMode = false
        previewSelectFormatIndex = -1
        previewSelectRowIndex = -1
        xmlPreviewDialog.close()
    }

    function scrollPreviewHoriz(stepDir) {
        if (!xmlPreviewDialogFlick) {
            return
        }
        var step = 36 * scaleFactor
        var maxX = Math.max(0, xmlPreviewDialogFlick.contentWidth - xmlPreviewDialogFlick.width)
        var nextX = xmlPreviewDialogFlick.contentX + (stepDir > 0 ? step : -step)
        xmlPreviewDialogFlick.contentX = Math.max(0, Math.min(maxX, nextX))
    }

    function openConfirmation(actionId, messageText) {
        confirmRequestToken = -1
        confirmTitle = "Confirm"
        confirmAction = String(actionId || "")
        confirmMessage = String(messageText || "Are you sure?")
        confirmDialog.open()
    }

    function openBackendConfirmation(tokenValue, titleText, messageText) {
        confirmAction = ""
        confirmRequestToken = tokenValue
        confirmTitle = String(titleText || "Confirm")
        confirmMessage = String(messageText || "Are you sure?")
        confirmDialog.open()
    }

    function openInfoDialog(messageText) {
        infoMessage = String(messageText || "")
        infoDialog.open()
    }

    function runConfirmedAction() {
        var actionId = confirmAction
        confirmDialog.close()
        if (actionId === "cancelProcess") {
            backendSafe.cancelCurrentOperation()
        } else if (actionId === "headerBack") {
            if (processState === "selecting") {
                backendSafe.selectDifferentFile()
            } else {
                backendSafe.convertAnotherFile()
            }
        } else if (actionId === "batchReviewBack") {
            backendSafe.convertAnotherFile()
        } else if (actionId === "deleteFormat") {
            if (pendingDeleteFormatIndex >= 0) {
                backendSafe.deleteFormatDefinition(pendingDeleteFormatIndex)
            }
        } else if (actionId === "formatCreateDiscard") {
            backendSafe.cancelFormatEdit()
            processState = "formatDesigner"
        }
        pendingDeleteFormatIndex = -1
        confirmAction = ""
        confirmRequestToken = -1
    }

    function scrollConvertingStatusListToBottom() {
        if (!convertingBatchScroll || !convertingBatchScroll.contentItem) {
            return
        }
        var flick = convertingBatchScroll.contentItem
        if (flick.contentY === undefined || flick.contentHeight === undefined) {
            return
        }
        var viewportHeight = flick.height !== undefined ? flick.height : 0
        var maxY = Math.max(0, flick.contentHeight - viewportHeight)
        flick.contentY = maxY
    }

    function anchorConvertingStatusToRecent() {
        if (!convertingBatchScroll || !convertingBatchScroll.contentItem) {
            return
        }
        var flick = convertingBatchScroll.contentItem
        if (flick.contentY === undefined || flick.contentHeight === undefined) {
            return
        }
        var viewportHeight = flick.height !== undefined ? flick.height : 0
        var maxY = Math.max(0, flick.contentHeight - viewportHeight)
        var lastCompletedIndex = -1
        if (batchFileStatuses && batchFileStatuses.length > 0) {
            for (var i = 0; i < batchFileStatuses.length; i++) {
                var s = String(batchFileStatuses[i] || "")
                if (s === "Done" || s === "Failed" || s === "Cancelled") {
                    lastCompletedIndex = i
                }
            }
        }
        var targetIndex = lastCompletedIndex >= 0
                        ? lastCompletedIndex
                        : Math.max(0, Math.min(Math.max(0, selectedFiles.length - 1), currentBatchIndex))
        var rowHeight = 22 * scaleFactor
        if (convertingBatchContent && convertingBatchContent.children && convertingBatchContent.children.length > 0) {
            var firstRow = convertingBatchContent.children[0]
            if (firstRow && firstRow.height !== undefined && firstRow.height > 0) {
                rowHeight = firstRow.height + convertingBatchContent.spacing
            }
        }
        var rowStep = Math.max(10 * scaleFactor, rowHeight)
        var targetY = targetIndex * rowStep
        flick.contentY = Math.max(0, Math.min(maxY, targetY))
    }

    function queueJumpFormatRowsToIndex(targetIndex) {
        formatRowsRestorePending = true
        Qt.callLater(function() {
            if (formatRowsListView && formatRowsListView.count > 0) {
                var idx = Math.max(0, Math.min(targetIndex, formatRowsListView.count - 1))
                formatRowsListView.positionViewAtIndex(idx, ListView.Center)
            }
            Qt.callLater(function() {
                if (formatRowsListView && formatRowsListView.count > 0) {
                    var idx = Math.max(0, Math.min(targetIndex, formatRowsListView.count - 1))
                    formatRowsListView.positionViewAtIndex(idx, ListView.Center)
                    formatRowsSavedContentY = formatRowsListView.contentY
                }
                formatRowsRestorePending = false
            })
        })
    }

    function rememberBatchOutputsScroll() {
        if (batchOutputsListView) {
            batchOutputsSavedContentY = batchOutputsListView.contentY
        }
    }

    function hasInvalidBatchNamesInModel() {
        if (!batchOutputs) {
            return false
        }
        for (var i = 0; i < batchOutputs.length; i++) {
            var draftName = batchFileNameDrafts[i]
            var baseNameToCheck = draftName !== undefined ? draftName : stripXlsx(batchOutputs[i].fileName)
            if (validateBatchBaseName(baseNameToCheck).length > 0) {
                return true
            }
        }
        return false
    }

    function hasInvalidBatchSaveDirsInModel() {
        if (!batchOutputs) {
            return false
        }
        for (var i = 0; i < batchOutputs.length; i++) {
            if (validateBatchSaveDir(batchOutputs[i].saveDir).length > 0) {
                return true
            }
        }
        return false
    }

    function hasBatchOutputConflictsInModel() {
        if (!batchOutputs || batchOutputs.length === 0) {
            return false
        }
        return backendSafe.estimateBatchOutputConflicts(batchOutputs).length > 0
    }

    function batchOutputConflicts() {
        if (!batchOutputs || batchOutputs.length === 0) {
            return []
        }
        return backendSafe.estimateBatchOutputConflicts(batchOutputs)
    }

    function batchOutputConflictSummary(limit) {
        var conflicts = batchOutputConflicts()
        if (!conflicts || conflicts.length === 0) {
            return ""
        }
        var maxLines = Math.max(1, limit || 4)
        var lines = []
        var seen = {}
        for (var i = 0; i < conflicts.length; i++) {
            var path = String(conflicts[i].path || "")
            if (path.length === 0 || seen[path]) {
                continue
            }
            seen[path] = true
            lines.push("- " + path)
            if (lines.length >= maxLines) {
                break
            }
        }
        var hidden = Math.max(0, conflicts.length - lines.length)
        return lines.join("\n") + (hidden > 0 ? ("\n...and " + hidden + " more conflict(s)") : "")
    }

    function batchFilesDescription() {
        if (!selectedFiles || selectedFiles.length === 0) {
            return "No files selected"
        }
        return selectedFiles.map(function(f) {
            return baseName(f) + " (" + backendSafe.getFileSize(f) + ")"
        }).join("\n")
    }

    function clearFormatCreateEditorFocus() {
        if (processState !== "formatCreate") {
            return
        }
        var focusedItem = activeFocusItem
        if (focusedItem && focusedItem !== rootWindow && focusedItem.focus !== undefined) {
            focusedItem.focus = false
        }
        formatEditorFocusType = ""
        formatDesignerSelectedRowIndex = -1
        if (contentItem) {
            contentItem.forceActiveFocus(Qt.ShortcutFocusReason)
        } else {
            forceActiveFocus(Qt.ShortcutFocusReason)
        }
    }

    function customLabelOptionsList() {
        if (!backendSafe || !backendSafe.customLabelOptions) {
            return []
        }
        return backendSafe.customLabelOptions
    }

    function customLabelOptionsByType(rowType) {
        var typeName = String(rowType || "data").toLowerCase()
        var options = customLabelOptionsList()
        if (typeName === "formula") {
            var formulaAllowed = { "demand": true, "kwh": true, "kvarh": true }
            var filtered = []
            for (var i = 0; i < options.length; i++) {
                var k = String(options[i] && options[i].key ? options[i].key : "")
                if (formulaAllowed[k]) {
                    filtered.push(options[i])
                }
            }
            return filtered
        }
        return options
    }

    function customLabelModel(rowType) {
        var options = customLabelOptionsByType(rowType)
        var labels = ["(No Label)"]
        for (var i = 0; i < options.length; i++) {
            var row2 = String(options[i] && options[i].row2 ? options[i].row2 : "")
            labels.push(row2)
        }
        return labels
    }

    function customLabelKeyForIndex(comboIndex, rowType) {
        if (comboIndex <= 0) {
            return ""
        }
        var options = customLabelOptionsByType(rowType)
        var idx = comboIndex - 1
        if (idx < 0 || idx >= options.length) {
            return ""
        }
        return String(options[idx] && options[idx].key ? options[idx].key : "")
    }

    function customLabelIndexForKey(labelKey, rowType) {
        var key = String(labelKey || "")
        if (key.length === 0) {
            return 0
        }
        if (key.indexOf("custom:") === 0) {
            return 0
        }
        var options = customLabelOptionsByType(rowType)
        for (var i = 0; i < options.length; i++) {
            if (String(options[i] && options[i].key ? options[i].key : "") === key) {
                return i + 1
            }
        }
        return 0
    }

    function customLabelIsCustom(labelKey) {
        return String(labelKey || "").indexOf("custom:") === 0
    }

    function customLabelTextFromKey(labelKey) {
        var key = String(labelKey || "")
        return key.indexOf("custom:") === 0 ? key.slice(7) : ""
    }

    function customLabelKeyFromText(textValue) {
        var t = String(textValue || "").trim()
        return t.length > 0 ? ("custom:" + t) : ""
    }

    onSelectionTypeChanged: {
        if (selectionType === "") {
            if (typeComboBox.currentIndex !== -1) {
                typeComboBox.currentIndex = -1
            }
        }

        var idx = typeComboBox.model.indexOf(selectionType);
        if (idx >= 0 && typeComboBox.currentIndex !== idx) {
            typeComboBox.currentIndex = idx;
        }
    }

    Connections {
        target: backend
        function onFormatModelChanged() {
            if (processState !== "formatCreate") {
                return
            }
            if (formatRowsJumpToIndexNextChange >= 0) {
                var jumpIndex = formatRowsJumpToIndexNextChange
                formatRowsJumpToIndexNextChange = -1
                Qt.callLater(function() {
                    rootWindow.queueJumpFormatRowsToIndex(jumpIndex)
                })
                return
            }
            var keepY = formatRowsScrollToBottomNextChange ? 1000000000 : formatRowsSavedContentY
            formatRowsScrollToBottomNextChange = false
            formatRowsSavedContentY = keepY
            queueRestoreListContentY(formatRowsListView, keepY)
        }
    }


    NumberAnimation {
        id: convertAnimation
        target: this
        property: "progress"
        from: 0
        to: 100
        duration: 2000
        onStopped: {
            processState = "creating"
            progress = 0
            createAnimation.start()
        }
    }

    NumberAnimation {
        id: createAnimation
        target: this
        property: "progress"
        from: 0
        to: 100
        duration: 1500
        onStopped: {
            processState = "complete"
        }
    }


    Popup {
        id: errorDialog
        modal: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
        width: 300 * scaleFactor
        height: 150 * scaleFactor

        Rectangle {
            anchors.fill: parent
            color: "white"
            border.color: "black"
            border.width: 1
            radius: 5

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20 * scaleFactor
                spacing: 10 * scaleFactor

                Text {
                    text: "Invalid File Type"
                    font.bold: true
                    font.pixelSize: 16 * scaleFactor
                    Layout.alignment: Qt.AlignHCenter
                }

                Text {
                    text: "Only XML files are supported.\nPlease select a valid XML file."
                    font.pixelSize: 12 * scaleFactor
                    Layout.alignment: Qt.AlignHCenter
                    horizontalAlignment: Text.AlignHCenter
                }

                PixelButton {
                    sliceLeft: 5
                    sliceRight: 5
                    sliceTop: 4
                    sliceBottom: 4
                    text: "OK"
                    textPixelSize: 11 * scaleFactor
                    Layout.alignment: Qt.AlignHCenter
                    Layout.preferredWidth: 90 * scaleFactor
                    Layout.preferredHeight: 34 * scaleFactor
                    fallbackNormal: themeLayer3
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    borderColor: themeLayer3
                    onClicked: errorDialog.close()
                }
            }
        }
    }

    Popup {
        id: columnLettersPopup
        modal: false
        focus: false
        closePolicy: Popup.NoAutoClose
        padding: 0
        x: 16 * scaleFactor
        y: 16 * scaleFactor
        width: 220 * scaleFactor
        height: 48 * scaleFactor
        z: 300
        background: Item {}

        Rectangle {
            anchors.fill: parent
            color: "#3a1f1f"
            border.color: "#dc2626"
            border.width: 1
            radius: 5

            Text {
                anchors.centerIn: parent
                text: "You can only enter letters."
                color: "#fee2e2"
                font.pixelSize: 10 * scaleFactor
                font.bold: true
            }
        }
    }

    Timer {
        id: columnLettersPopupTimer
        interval: 1200
        repeat: false
        onTriggered: columnLettersPopup.close()
    }

    Popup {
        id: widthNumbersPopup
        modal: false
        focus: false
        closePolicy: Popup.NoAutoClose
        padding: 0
        x: 16 * scaleFactor
        y: 16 * scaleFactor
        width: 220 * scaleFactor
        height: 48 * scaleFactor
        z: 300
        background: Item {}

        Rectangle {
            anchors.fill: parent
            color: "#3a1f1f"
            border.color: "#dc2626"
            border.width: 1
            radius: 5

            Text {
                anchors.centerIn: parent
                text: "You can only enter numbers."
                color: "#fee2e2"
                font.pixelSize: 10 * scaleFactor
                font.bold: true
            }
        }
    }

    Timer {
        id: widthNumbersPopupTimer
        interval: 1200
        repeat: false
        onTriggered: widthNumbersPopup.close()
    }

    Popup {
        id: confirmDialog
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
        width: Math.min(rootWindow.width - (32 * scaleFactor), 320 * scaleFactor)
        height: 160 * scaleFactor
        padding: 0
        z: 350
        background: Item {}
        onClosed: {
            if (rootWindow.confirmRequestToken >= 0) {
                backendSafe.resolveInAppConfirm(rootWindow.confirmRequestToken, false)
                rootWindow.confirmRequestToken = -1
            }
        }

        Rectangle {
            anchors.fill: parent
            color: themePanel
            border.color: themeLayer2
            border.width: 1
            radius: 6

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 12 * scaleFactor
                spacing: 12 * scaleFactor

                Text {
                    text: rootWindow.confirmTitle
                    color: themeText
                    font.family: appFontFamily
                    font.pixelSize: 14 * scaleFactor
                    font.bold: true
                    Layout.fillWidth: true
                    horizontalAlignment: Text.AlignHCenter
                }

                Text {
                    text: confirmMessage
                    color: themeText
                    font.family: appFontFamily
                    font.pixelSize: 11 * scaleFactor
                    wrapMode: Text.Wrap
                    Layout.fillWidth: true
                    horizontalAlignment: Text.AlignHCenter
                }

                RowLayout {
                    Layout.fillWidth: true
                    spacing: 8 * scaleFactor

                    PixelButton {
                        sliceLeft: 5
                        sliceRight: 5
                        sliceTop: 4
                        sliceBottom: 4
                        Layout.fillWidth: true
                        Layout.preferredHeight: 36 * scaleFactor
                        text: "No"
                        textPixelSize: 11 * scaleFactor
                        fallbackNormal: themePanel
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        textColor: themeText
                        borderColor: themeLayer3
                        onClicked: {
                            if (rootWindow.confirmRequestToken >= 0) {
                                backendSafe.resolveInAppConfirm(rootWindow.confirmRequestToken, false)
                                rootWindow.confirmRequestToken = -1
                            }
                            confirmDialog.close()
                        }
                    }

                    PixelButton {
                        sliceLeft: 5
                        sliceRight: 5
                        sliceTop: 4
                        sliceBottom: 4
                        Layout.fillWidth: true
                        Layout.preferredHeight: 36 * scaleFactor
                        text: "Yes"
                        textPixelSize: 11 * scaleFactor
                        fallbackNormal: themeLayer3
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        textColor: themeText
                        borderColor: themeLayer3
                        onClicked: {
                            if (rootWindow.confirmRequestToken >= 0) {
                                backendSafe.resolveInAppConfirm(rootWindow.confirmRequestToken, true)
                                rootWindow.confirmRequestToken = -1
                                confirmDialog.close()
                            } else {
                                rootWindow.runConfirmedAction()
                            }
                        }
                    }
                }
            }
        }
    }

    Popup {
        id: xmlPreviewDialog
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
        width: Math.min(rootWindow.width - (20 * scaleFactor), 760 * scaleFactor)
        height: Math.min(rootWindow.height - (28 * scaleFactor), 420 * scaleFactor)
        padding: 0
        z: 360
        background: Item {}
        onClosed: {
            previewSelectMode = false
            previewSelectFormatIndex = -1
            previewSelectRowIndex = -1
        }
        onOpened: {
            forceActiveFocus()
            if (xmlPreviewDialogFlick) {
                xmlPreviewDialogFlick.forceActiveFocus()
            }
        }

        Rectangle {
            anchors.fill: parent
            color: themePanel
            border.color: themeLayer2
            border.width: 1
            radius: 6
            clip: true

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 8 * scaleFactor
                spacing: 6 * scaleFactor

                RowLayout {
                    Layout.fillWidth: true
                    spacing: 6 * scaleFactor

                    Text {
                        Layout.fillWidth: true
                        text: previewSelectMode ? "Select Source Index" : "XML Preview"
                        color: themeText
                        font.family: appFontFamily
                        font.pixelSize: 12 * scaleFactor
                        font.bold: true
                        elide: Text.ElideRight
                    }

                    PixelButton {
                        sliceLeft: 4
                        sliceRight: 4
                        sliceTop: 4
                        sliceBottom: 4
                        Layout.preferredWidth: 72 * scaleFactor
                        Layout.preferredHeight: 26 * scaleFactor
                        text: "Close"
                        textPixelSize: 10 * scaleFactor
                        fallbackNormal: themePanel
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        borderColor: themeLayer3
                        onClicked: xmlPreviewDialog.close()
                    }
                }

                Text {
                    Layout.fillWidth: true
                    text: backendSafe.xmlPreviewStatus
                    color: themeTextSecondary
                    font.pixelSize: 10 * scaleFactor
                    elide: Text.ElideRight
                }

                Rectangle {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    color: themeInset
                    border.color: themeLayer2
                    border.width: 1
                    radius: 5
                    clip: true

                    Flickable {
                        id: xmlPreviewDialogFlick
                        focus: true
                        anchors.fill: parent
                        anchors.margins: 6 * scaleFactor
                        clip: true
                        contentWidth: xmlPreviewDialogTable.implicitWidth
                        contentHeight: xmlPreviewDialogTable.implicitHeight

                        ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }
                        ScrollBar.horizontal: ScrollBar { policy: ScrollBar.AsNeeded }

                        WheelHandler {
                            onWheel: function(event) {
                                var delta = event.angleDelta.y !== 0 ? event.angleDelta.y : event.angleDelta.x
                                var step = 36 * scaleFactor
                                var maxX = Math.max(0, xmlPreviewDialogFlick.contentWidth - xmlPreviewDialogFlick.width)
                                var nextX = xmlPreviewDialogFlick.contentX + (delta < 0 ? step : -step)
                                xmlPreviewDialogFlick.contentX = Math.max(0, Math.min(maxX, nextX))
                                event.accepted = true
                            }
                        }

                        Column {
                            id: xmlPreviewDialogTable
                            spacing: 4 * scaleFactor

                            Row {
                                spacing: 4 * scaleFactor

                                Repeater {
                                    model: backendSafe.xmlPreviewHeaders

                                    Rectangle {
                                        width: 90 * scaleFactor
                                        height: 24 * scaleFactor
                                        color: themePanel
                                        border.color: themeLayer3
                                        border.width: 1
                                        radius: 3

                                        Text {
                                            anchors.centerIn: parent
                                            text: "[" + modelData.index + "]"
                                            color: themeText
                                            font.pixelSize: 9 * scaleFactor
                                            font.bold: true
                                        }

                                        MouseArea {
                                            anchors.fill: parent
                                            enabled: previewSelectMode
                                            cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                                            onClicked: rootWindow.pickSourceIndexFromPreview(modelData.index)
                                        }
                                    }
                                }
                            }

                            Repeater {
                                model: backendSafe.xmlPreviewRows

                                Row {
                                    spacing: 4 * scaleFactor

                                    Repeater {
                                        model: modelData

                                        Rectangle {
                                            width: 90 * scaleFactor
                                            height: 22 * scaleFactor
                                            color: themePanel
                                            border.color: themeLayer2
                                            border.width: 1
                                            radius: 2

                                            Text {
                                                anchors.verticalCenter: parent.verticalCenter
                                                anchors.left: parent.left
                                                anchors.leftMargin: 4 * scaleFactor
                                                anchors.right: parent.right
                                                anchors.rightMargin: 4 * scaleFactor
                                                text: String(modelData)
                                                color: themeText
                                                font.pixelSize: 9 * scaleFactor
                                                elide: Text.ElideRight
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    Shortcut {
        sequence: "Left"
        context: Qt.ApplicationShortcut
        enabled: xmlPreviewDialog.visible
        onActivated: {
            rootWindow.scrollPreviewHoriz(-1)
        }
    }

    Shortcut {
        sequence: "Right"
        context: Qt.ApplicationShortcut
        enabled: xmlPreviewDialog.visible
        onActivated: {
            rootWindow.scrollPreviewHoriz(1)
        }
    }


    ColumnLayout {
        anchors.fill: parent
        anchors.leftMargin: 20 * scaleFactor
        anchors.rightMargin: 20 * scaleFactor
        anchors.bottomMargin: ((processState === "converting" || processState === "creating")
                               ? 72
                               : (processState === "selecting" ? 18 : 20)) * scaleFactor
        anchors.topMargin: ((20 - (12 * headerCompress)) * scaleFactor) + headerReservedTopSpace
        spacing: (20 - (8 * headerCompress)) * scaleFactor


        ColumnLayout {
            Layout.alignment: Qt.AlignHCenter
            spacing: (10 - (4 * headerCompress)) * scaleFactor
            visible: processState !== "batchReview" && processState !== "converting" && processState !== "creating" && processState !== "complete" && processState !== "formatDesigner" && processState !== "formatCreate" && !compactHeaderMode

            AnimatedImage {
                id: copywriting
                Layout.preferredWidth: Math.min(340 * scaleFactor * headerScale, rootWindow.width - (80 * scaleFactor))
                Layout.preferredHeight: Math.min(200 * scaleFactor * headerScale, rootWindow.height * 0.24)
                source: "images/copywriting.gif"
                speed: 1.0
                fillMode: Image.PreserveAspectFit
                transformOrigin: Item.Center
                scale: 1.25
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: qsTr("CubeFlow")
                color: "white"
                font.family: appFontFamily
                font.pixelSize: 32 * scaleFactor * headerScale
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                color: "white"
                text: "Convert your XML files with ease"
                font.pixelSize: 16 * scaleFactor * headerScale
                font.family: appFontFamily
                Layout.alignment: Qt.AlignHCenter
                visible: !compactHeaderMode && headerCompress < 0.45
            }

            Text {
                color: themeTextSecondary
                text: "by wahchachaps"
                font.pixelSize: 12 * scaleFactor * headerScale
                font.family: appFontFamily
                Layout.alignment: Qt.AlignHCenter
            }
        }

        Rectangle {
            Layout.fillWidth: true
            Layout.preferredHeight: 165 * scaleFactor
            color: themePanel
            border.color: dropArea.containsDrag ? themeLayer3 : themeLayer2
            border.width: dropArea.containsDrag ? 3 : 2
            radius: 10
            visible: processState === "idle"


            MouseArea {
                id: selectFileArea
                anchors.fill: parent
                hoverEnabled: true
                cursorShape: Qt.PointingHandCursor
                onClicked: {
                    backendSafe.selectFile()
                    selectionType = ""
                    typeComboBox.currentIndex = -1
                }
            }


            DropArea {
                id: dropArea
                anchors.fill: parent
                keys: ["text/uri-list"]

                onEntered: {

                    console.log("Drag entered drop zone")
                }

                onExited: {

                    console.log("Drag exited drop zone")
                }

                onDropped: function(drop) {
                    if (!drop.hasUrls) {
                        return
                    }
                    var droppedPaths = []
                    for (var i = 0; i < drop.urls.length; i++) {
                        var raw = drop.urls[i].toString()
                        var path = decodeURIComponent(raw.replace("file:///", ""))
                        droppedPaths.push(path)
                    }
                    backendSafe.setDroppedPaths(droppedPaths)
                    selectionType = ""
                    typeComboBox.currentIndex = -1
                }
            }


            Rectangle {
                anchors.fill: parent
                color: dropArea.containsDrag ? "#cce7ff" : (selectFileArea.containsMouse ? "#e0e7ff" : "transparent")
                border.color: dropArea.containsDrag ? themeLayer3 : (selectFileArea.containsMouse ? themeLayer3 : themeLayer2)
                border.width: dropArea.containsDrag ? 3 : (selectFileArea.containsMouse ? 2 : 2)
                radius: 10
                opacity: dropArea.containsDrag ? 0.4 : (selectFileArea.containsMouse ? 0.5 : 0)
                Behavior on opacity { NumberAnimation { duration: 200 } }
                Behavior on color { ColorAnimation { duration: 200 } }
                Behavior on border.width { NumberAnimation { duration: 200 } }
            }

            AnimatedImage {
                anchors.centerIn: parent
                width: parent.width + (15 * scaleFactor)
                height: parent.height + (15 * scaleFactor)
                source: "images/ui/dropzone_fx.gif"
                playing: processState === "idle" && !selectFileArea.containsMouse
                fillMode: Image.Stretch
                smooth: false
                mipmap: false
                opacity: dropArea.containsDrag ? 1.0 : 0.85
            }

            ColumnLayout {
                anchors.centerIn: parent
                spacing: 10 * scaleFactor



                Rectangle {
                    Layout.preferredWidth: 48 * scaleFactor
                    Layout.preferredHeight: 48 * scaleFactor
                    color: "transparent"
                    Layout.alignment: Qt.AlignHCenter

                    Image {
                        source: "images/file.png"
                        fillMode: Image.PreserveAspectFit
                        width: 75
                        height: 75
                        anchors.centerIn: parent
                    }
                }

                Text {
                    text: dropArea.containsDrag ? "Drop XML file(s) or folder here" : "Click to select XML file(s)"
                    color: "white"
                    font.pixelSize: 16 * scaleFactor
                    font.bold: true
                    Layout.alignment: Qt.AlignHCenter
                }

                Text {
                    text: "XML files only"
                    color: "white"
                    font.pixelSize: 12 * scaleFactor
                    Layout.alignment: Qt.AlignHCenter
                }
            }
        }

        PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
            Layout.fillWidth: true
            Layout.preferredHeight: 36 * scaleFactor
            visible: processState === "idle"
            text: "Open Format Designer"
            textPixelSize: 12 * scaleFactor
            normalSource: Qt.resolvedUrl("images/ui/button_normal.png")
            hoverSource: Qt.resolvedUrl("images/ui/button_hover.png")
            pressedSource: Qt.resolvedUrl("images/ui/button_pressed.png")
            disabledSource: Qt.resolvedUrl("images/ui/button_disabled.png")
            fallbackNormal: themeLayer3
            fallbackHover: themeLayer2
            fallbackPressed: themeLayer1
            fallbackDisabled: "#9ca3af"
            borderColor: themeLayer3
            onClicked: backendSafe.openFormatDesigner()
        }


        ColumnLayout {
            visible: processState === "selecting"
            spacing: (8 - (2 * headerCompress)) * scaleFactor
            Layout.fillHeight: true

            Text {
                visible: isBatch
                text: selectedFiles.length + " file(s) selected"
                font.pixelSize: 13 * scaleFactor
                font.bold: true
                color: themeText
            }

            Item {
                id: fileInfoContainer
                Layout.fillWidth: true
                Layout.fillHeight: true
                Layout.minimumHeight: fileInfoBaseHeight
                Layout.preferredHeight: fileInfoCardHeight
                Layout.maximumHeight: fileInfoMaxHeight

                Rectangle {
                    id: fileInfoCard
                    anchors.fill: parent
                    color: themeInset
                    border.color: themeLayer2
                    border.width: 1
                    radius: 5
                    clip: true

                    ScrollView {
                        id: fileInfoScroll
                        anchors.fill: parent
                        anchors.leftMargin: 8 * scaleFactor
                        anchors.topMargin: 8 * scaleFactor
                        anchors.bottomMargin: 8 * scaleFactor
                        anchors.rightMargin: 0
                        clip: true
                        ScrollBar.vertical.policy: ScrollBar.AsNeeded

                    Item {
                        width: Math.max(0, fileInfoScroll.availableWidth - (16 * scaleFactor))
                        implicitHeight: isBatch ? batchFileListColumn.implicitHeight : singleFileInfoText.implicitHeight

                        Column {
                            id: batchFileListColumn
                            width: parent.width
                            spacing: 6 * scaleFactor
                            visible: isBatch

                            Repeater {
                                model: selectedFiles.length

                                Rectangle {
                                    width: batchFileListColumn.width
                                    height: Math.max(30 * scaleFactor, fileNameText.implicitHeight + (8 * scaleFactor))
                                    color: themePanel
                                    border.color: themeLayer2
                                    border.width: 1
                                    radius: 4

                                    RowLayout {
                                        anchors.fill: parent
                                        anchors.margins: 4 * scaleFactor
                                        spacing: 6 * scaleFactor

                                        Text {
                                            id: fileNameText
                                            Layout.fillWidth: true
                                            text: baseName(selectedFiles[index]) + " (" + backendSafe.getFileSize(selectedFiles[index]) + ")"
                                            font.pixelSize: 12 * scaleFactor
                                            color: themeText
                                            elide: Text.ElideMiddle
                                            verticalAlignment: Text.AlignVCenter
                                        }

                                        Text {
                                            Layout.preferredWidth: 66 * scaleFactor
                                            text: (batchFileStatuses && batchFileStatuses.length > index) ? batchFileStatuses[index] : "Queued"
                                            font.pixelSize: 11 * scaleFactor
                                            color: text === "Done" ? "#059669"
                                                  : text === "Failed" ? "#dc2626"
                                                  : text === "Processing" ? themeLayer3
                                                  : "#6b7280"
                                            horizontalAlignment: Text.AlignRight
                                            verticalAlignment: Text.AlignVCenter
                                            elide: Text.ElideRight
                                        }

                                        Item {
                                            Layout.preferredWidth: 22 * scaleFactor
                                            Layout.preferredHeight: 22 * scaleFactor

                                            Text {
                                                anchors.centerIn: parent
                                                text: "x"
                                                color: "#b91c1c"
                                                font.pixelSize: 12 * scaleFactor
                                                font.bold: true
                                            }

                                            MouseArea {
                                                anchors.fill: parent
                                                cursorShape: Qt.PointingHandCursor
                                                onClicked: backendSafe.removeSelectedFile(index)
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        Text {
                            id: singleFileInfoText
                            width: parent.width
                            visible: !isBatch
                            text: "File: " + (selectedFile ? baseName(selectedFile) : "") + "\nSize: " + fileSize
                            color: "white"
                            font.pixelSize: 13 * scaleFactor
                            wrapMode: Text.Wrap
                            elide: Text.ElideNone
                            lineHeight: 1.15
                            lineHeightMode: Text.ProportionalHeight
                        }
                    }
                }
            }
            }

            Text {
                text: "Select Conversion Type"
                color: "white"
                font.pixelSize: 14 * scaleFactor
                font.bold: true
            }


            PixelComboBox {
                sliceLeft: 5
                sliceRight: 5
                sliceTop: 4
                sliceBottom: 4
                id: typeComboBox
                Layout.fillWidth: true
                Layout.preferredHeight: primaryControlHeight
                textPixelSize: primaryControlFontSize
                popupTextPixelSize: 11 * scaleFactor
                model: backendSafe.xmlTypeOptions
                currentIndex: -1
                displayText: currentIndex === -1 ? "Select XML type" : currentText
                fallbackNormal: themeInset
                fallbackFocus: themeLayer1
                fallbackOpen: themeLayer2
                fallbackDisabled: themeLayer2
                fallbackBorder: themeLayer2
                fallbackText: themeText
                fallbackPlaceholder: "white"
                fallbackPopup: themePanel
                onCurrentTextChanged: {
                    if (currentIndex >= 0 && currentText !== selectionType) {
                        backendSafe.setSelectionType(currentText)
                    }
                }
            }


            PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                Layout.fillWidth: true
                Layout.preferredHeight: primaryControlHeight
                text: "Confirm and Convert"
                textPixelSize: primaryControlFontSize
                enabled: selectionType !== ""
                fallbackNormal: themeLayer3
                fallbackHover: themeLayer2
                fallbackPressed: themeLayer1
                fallbackDisabled: "#9ca3af"
                borderColor: themeLayer3
                onClicked: {
                    if (typeComboBox.currentIndex >= 0) {
                        backendSafe.setSelectionType(typeComboBox.currentText)
                    }
                    backendSafe.confirmAndConvert()
                }
            }


            PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                Layout.fillWidth: true
                Layout.preferredHeight: primaryControlHeight
                text: "Select Different File"
                textPixelSize: primaryControlFontSize
                fallbackNormal: themeLayer3
                fallbackHover: themeLayer2
                fallbackPressed: themeLayer1
                borderColor: themeLayer3
                onClicked: backendSafe.selectFile()
            }
        }


        ColumnLayout {
            visible: processState === "formatDesigner"
            spacing: 8 * scaleFactor
            Layout.fillWidth: true
            Layout.fillHeight: true

            Text {
                text: "Select Format"
                color: "white"
                font.family: appFontFamily
                font.pixelSize: 22 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Click a format to edit it"
                color: themeTextSecondary
                font.pixelSize: 11 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                id: createFormatLink
                text: "Create Format"
                color: createFormatArea.containsMouse ? themeLayer3 : "white"
                font.family: appFontFamily
                font.pixelSize: 12 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignRight
                horizontalAlignment: Text.AlignRight

                MouseArea {
                    id: createFormatArea
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: {
                        var newIndex = backendSafe.createFormatDraft()
                        rootWindow.formatDesignerSelectedFormatIndex = Math.max(0, newIndex)
                        rootWindow.formatDesignerSelectedRowIndex = -1
                        processState = "formatCreate"
                    }
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 6
                clip: true

                ListView {
                    id: formatDesignerListView
                    property var windowRef: rootWindow
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 6 * scaleFactor
                    model: backendSafe.formatModel
                    clip: true

                    delegate: Rectangle {
                        width: ListView.view.width
                        height: 44 * scaleFactor
                        radius: 5
                        property bool isBuiltInFormat: backendSafe.isBuiltinFormat(modelData.name)
                        color: rootWindow.formatDesignerSelectedFormatIndex === index ? themeLayer1 : themePanel
                        border.color: rootWindow.formatDesignerSelectedFormatIndex === index ? themeLayer3 : themeLayer2
                        border.width: 1

                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 6 * scaleFactor
                            spacing: 6 * scaleFactor

                            Text {
                                Layout.fillWidth: true
                                text: modelData.name
                                color: themeText
                                font.pixelSize: 12 * scaleFactor
                                font.bold: true
                                elide: Text.ElideRight
                            }

                            PixelButton {
                        sliceLeft: 4
                        sliceRight: 4
                        sliceTop: 4
                        sliceBottom: 4
                                Layout.preferredWidth: 60 * scaleFactor
                                Layout.preferredHeight: 30 * scaleFactor
                                text: isBuiltInFormat ? "Open" : "Edit"
                                textPixelSize: 10 * scaleFactor
                                fallbackNormal: themeLayer3
                                fallbackHover: themeLayer2
                                fallbackPressed: themeLayer1
                                borderColor: themeLayer3
                                onClicked: backendSafe.openFormatForEdit(index)
                            }

                            PixelButton {
                                sliceLeft: 4
                                sliceRight: 4
                                sliceTop: 4
                                sliceBottom: 4
                                Layout.preferredWidth: 74 * scaleFactor
                                Layout.preferredHeight: 30 * scaleFactor
                                text: "Duplicate"
                                textPixelSize: 9 * scaleFactor
                                fallbackNormal: themeLayer3
                                fallbackHover: themeLayer2
                                fallbackPressed: themeLayer1
                                borderColor: themeLayer3
                                onClicked: backendSafe.duplicateFormatAndOpen(index)
                            }

                            PixelButton {
                                sliceLeft: 4
                                sliceRight: 4
                                sliceTop: 4
                                sliceBottom: 4
                                visible: !isBuiltInFormat
                                Layout.preferredWidth: 60 * scaleFactor
                                Layout.preferredHeight: 30 * scaleFactor
                                text: "Delete"
                                textPixelSize: 10 * scaleFactor
                                enabled: backendSafe.formatModel.length > 1
                                fallbackNormal: "#dc2626"
                                fallbackHover: "#b91c1c"
                                fallbackPressed: "#991b1b"
                                fallbackDisabled: "#9ca3af"
                                borderColor: "#dc2626"
                                onClicked: {
                                    rootWindow.pendingDeleteFormatIndex = index
                                    rootWindow.openConfirmation("deleteFormat", "Delete this format?")
                                }
                            }
                        }
                    }
                }
            }

            Text {
                Layout.fillWidth: true
                text: backendSafe.formatDesignerStatus
                color: backendSafe.formatDesignerStatus.indexOf("Failed") === 0 ? "#dc2626" : themeLayer3
                font.pixelSize: 10 * scaleFactor
                wrapMode: Text.Wrap
                visible: backendSafe.formatDesignerStatus.length > 0
            }

            PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                Layout.fillWidth: true
                Layout.preferredHeight: 38 * scaleFactor
                text: "Open File"
                textPixelSize: 12 * scaleFactor
                fallbackNormal: themeLayer3
                fallbackHover: themeLayer2
                fallbackPressed: themeLayer1
                borderColor: themeLayer3
                onClicked: backendSafe.importFormatModelFromFile()
            }

            PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                Layout.fillWidth: true
                Layout.preferredHeight: 36 * scaleFactor
                text: "Back"
                textPixelSize: 12 * scaleFactor
                fallbackNormal: themePanel
                fallbackHover: themeLayer2
                fallbackPressed: themeLayer1
                textColor: themeText
                borderColor: themeLayer3
                onClicked: backendSafe.closeFormatDesigner()
            }
        }

        ColumnLayout {
            id: formatCreatePanel
            visible: processState === "formatCreate"
            property int selectedFormatIndex: rootWindow.formatDesignerSelectedFormatIndex
            property bool selectedBuiltInFormat: (
                backendSafe.formatModel.length > 0
                && selectedFormatIndex >= 0
                && selectedFormatIndex < backendSafe.formatModel.length
                && backendSafe.isBuiltinFormat(backendSafe.formatModel[selectedFormatIndex].name)
            )
            spacing: 8 * scaleFactor
            Layout.fillWidth: true
            Layout.fillHeight: true

            Shortcut {
                sequence: "Esc"
                enabled: formatCreatePanel.visible
                context: Qt.WindowShortcut
                onActivated: rootWindow.clearFormatCreateEditorFocus()
            }

            Text {
                text: formatCreatePanel.selectedBuiltInFormat ? "View Format" : "Create Format"
                color: "white"
                font.family: appFontFamily
                font.pixelSize: 22 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            PixelTextField {
                id: formatNameField
                Layout.fillWidth: true
                Layout.preferredHeight: 50 * scaleFactor
                font.pixelSize: 16 * scaleFactor
                text: (backendSafe.formatModel.length > 0 && rootWindow.formatDesignerSelectedFormatIndex >= 0 && rootWindow.formatDesignerSelectedFormatIndex < backendSafe.formatModel.length)
                    ? backendSafe.formatModel[rootWindow.formatDesignerSelectedFormatIndex].name
                    : ""
                placeholderText: "Format name"
                readOnly: formatCreatePanel.selectedBuiltInFormat
                onEditingFinished: {
                    if (!formatCreatePanel.selectedBuiltInFormat) {
                        backendSafe.renameFormatDefinition(rootWindow.formatDesignerSelectedFormatIndex, text)
                    }
                }
                Keys.onEscapePressed: function(event) {
                    focus = false
                    rootWindow.clearFormatCreateEditorFocus()
                    event.accepted = true
                }
            }

            RowLayout {
                Layout.fillWidth: true
                spacing: 6 * scaleFactor

                PixelButton {
            sliceLeft: 4
            sliceRight: 4
            sliceTop: 4
            sliceBottom: 4
                    Layout.preferredWidth: 80 * scaleFactor
                    Layout.preferredHeight: 28 * scaleFactor
                    text: "Add Column"
                    textPixelSize: 10 * scaleFactor
                    enabled: backendSafe.formatModel.length > 0 && !formatCreatePanel.selectedBuiltInFormat
                    fallbackNormal: themeLayer3
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    fallbackDisabled: "#9ca3af"
                    borderColor: themeLayer3
                    onClicked: {
                        var newIndex = backendSafe.addFormatRow(rootWindow.formatDesignerSelectedFormatIndex)
                        if (newIndex >= 0) {
                            rootWindow.formatEditorFocusType = ""
                            rootWindow.formatDesignerSelectedRowIndex = -1
                            rootWindow.markAndJumpFormatRow(newIndex)
                        }
                    }
                }

                PixelButton {
                    sliceLeft: 4
                    sliceRight: 4
                    sliceTop: 4
                    sliceBottom: 4
                    Layout.preferredWidth: 90 * scaleFactor
                    Layout.preferredHeight: 28 * scaleFactor
                    text: "Load Preview"
                    textPixelSize: 10 * scaleFactor
                    fallbackNormal: themeLayer3
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    borderColor: themeLayer3
                    onClicked: {
                        backendSafe.selectPreviewXmlFile()
                        rootWindow.openPreviewDialogReadOnly()
                    }
                }

                PixelButton {
                    sliceLeft: 4
                    sliceRight: 4
                    sliceTop: 4
                    sliceBottom: 4
                    Layout.preferredWidth: 120 * scaleFactor
                    Layout.preferredHeight: 28 * scaleFactor
                    text: "Select Another XML"
                    textPixelSize: 9 * scaleFactor
                    fallbackNormal: themePanel
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    borderColor: themeLayer3
                    onClicked: {
                        backendSafe.selectAnotherPreviewXmlFile()
                        rootWindow.openPreviewDialogReadOnly()
                    }
                }
            }

            Text {
                Layout.fillWidth: true
                text: "Tips: XML Data uses XML column index (0, 1, 2...). Formula should start with '=' and can use {r} and {r-1}."
                color: themeText
                font.pixelSize: 10 * scaleFactor
                wrapMode: Text.Wrap
            }

            Text {
                Layout.fillWidth: true
                text: backendSafe.xmlPreviewStatus
                color: themeTextSecondary
                font.pixelSize: 10 * scaleFactor
                elide: Text.ElideRight
            }

            RowLayout {
                Layout.fillWidth: true
                Layout.leftMargin: 14 * scaleFactor
                Layout.rightMargin: 14 * scaleFactor
                Layout.topMargin: 2 * scaleFactor
                Layout.bottomMargin: 2 * scaleFactor
                spacing: 6 * scaleFactor

                Item {
                    visible: !formatCreatePanel.selectedBuiltInFormat
                    Layout.preferredWidth: 6 * scaleFactor
                    Layout.preferredHeight: 1
                }

                Text {
                    Layout.preferredWidth: 48 * scaleFactor
                    text: "Column"
                    color: themeText
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "column") ? 1.0 : 0.0
                }

                Text {
                    Layout.fillWidth: true
                    text: "Label"
                    color: themeText
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    visible: rootWindow.formatCompactLabelsMode
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "label") ? 1.0 : 0.0
                }

                Text {
                    Layout.preferredWidth: 82 * scaleFactor
                    text: "Source"
                    color: themeText
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    visible: !rootWindow.formatCompactLabelsMode
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "source") ? 1.0 : 0.0
                }

                Text {
                    Layout.fillWidth: true
                    text: "Value"
                    color: themeText
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    visible: !rootWindow.formatCompactLabelsMode
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "value") ? 1.0 : 0.0
                }

                Text {
                    Layout.preferredWidth: 64 * scaleFactor
                    text: "Width"
                    color: themeText
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    visible: !rootWindow.formatCompactLabelsMode
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "width") ? 1.0 : 0.0
                }

                PixelButton {
                    sliceLeft: 4
                    sliceRight: 4
                    sliceTop: 4
                    sliceBottom: 4
                    Layout.preferredWidth: 26 * scaleFactor
                    Layout.preferredHeight: 22 * scaleFactor
                    text: rootWindow.formatCompactLabelsMode ? "<" : ">"
                    textPixelSize: 10 * scaleFactor
                    fallbackNormal: themePanel
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    borderColor: themeLayer3
                    onClicked: rootWindow.formatCompactLabelsMode = !rootWindow.formatCompactLabelsMode
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 6
                clip: true

                ListView {
                    id: formatRowsListView
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 6 * scaleFactor
                    clip: true
                    model: (backendSafe.formatModel.length > 0 && rootWindow.formatDesignerSelectedFormatIndex >= 0 && rootWindow.formatDesignerSelectedFormatIndex < backendSafe.formatModel.length)
                        ? backendSafe.formatModel[rootWindow.formatDesignerSelectedFormatIndex].columns
                        : []
                    onModelChanged: {
                        if (formatRowsJumpToIndexNextChange >= 0 || formatRowsRestorePending) {
                            return
                        }
                        var keepY = formatRowsSavedContentY
                        queueRestoreListContentY(formatRowsListView, keepY)
                    }
                    onContentYChanged: {
                        if (!formatRowsRestorePending) {
                            formatRowsSavedContentY = contentY
                        }
                    }
                    onMovementEnded: {
                        formatRowsSavedContentY = contentY
                    }

                    delegate: Rectangle {
                        property int rowIndex: index
                        property real dragStartGlobalY: 0
                        property real dragStartContentY: 0
                        width: ListView.view.width
                        height: 50 * scaleFactor
                        radius: 5
                        function applyFormulaTemplate(templateText) {
                            valueField.text = templateText
                            rootWindow.prepareFormatRowsModelChange(false)
                            var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "value", templateText)
                            rootWindow.markAndJumpFormatRow(updatedIndex)
                            valueField.forceActiveFocus()
                            valueField.cursorPosition = valueField.text.length
                        }
                        property bool rowExpanded: (
                            colField.activeFocus
                            || valueField.activeFocus
                            || widthField.activeFocus
                            || typeCombo.activeFocus
                            || (typeCombo.popup && typeCombo.popup.visible)
                            || (rootWindow.formatCompactLabelsMode && labelCustomField.activeFocus)
                            || (rootWindow.formatCompactLabelsMode && labelPresetCombo.activeFocus)
                            || (rootWindow.formatCompactLabelsMode && labelPresetCombo.popup && labelPresetCombo.popup.visible)
                            || (formulaTemplateMenu && formulaTemplateMenu.visible)
                        )
                        onRowExpandedChanged: {
                            if (rowExpanded) {
                                rootWindow.formatRowsHighlightIndex = index
                                rootWindow.formatDesignerSelectedRowIndex = index
                            }
                            if (!rowExpanded && rootWindow.formatDesignerSelectedRowIndex === index) {
                                rootWindow.formatDesignerSelectedRowIndex = -1
                            }
                            if (!rowExpanded) {
                                rootWindow.formatEditorFocusType = ""
                            }
                        }
                        color: rootWindow.formatDesignerSelectedRowIndex === index ? themeLayer1 : themePanel
                        border.color: rowExpanded
                                      ? "#dc2626"
                                      : (rootWindow.formatRowsHighlightIndex === index
                                      ? "#dc2626"
                                      : (rootWindow.formatDesignerSelectedRowIndex === index ? themeLayer3 : themeLayer2))
                        border.width: 1

                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 6 * scaleFactor
                            spacing: 6 * scaleFactor

                            Rectangle {
                                visible: !formatCreatePanel.selectedBuiltInFormat
                                Layout.preferredWidth: 6 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                radius: 4
                                color: dragHandleArea.pressed ? themeLayer1 : (dragHandleArea.containsMouse ? themeLayer2 : themePanel)
                                border.width: 1
                                border.color: themeLayer3

                                Text {
                                    anchors.centerIn: parent
                                    text: "||"
                                    color: themeText
                                    font.pixelSize: 7 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    id: dragHandleArea
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    cursorShape: Qt.SizeVerCursor
                                    preventStealing: true
                                    onPressed: function(mouse) {
                                        var p = parent.mapToItem(rootWindow.contentItem, mouse.x, mouse.y)
                                        dragStartGlobalY = p.y
                                        dragStartContentY = formatRowsListView.contentY
                                        rootWindow.formatDragGhostVisible = true
                                        rootWindow.formatDragGhostY = p.y - (18 * scaleFactor)
                                        rootWindow.formatDragGhostText = "Moving " + String(modelData.col || "")
                                    }
                                    onPositionChanged: function(mouse) {
                                        if (!pressed) {
                                            return
                                        }
                                        var p = parent.mapToItem(rootWindow.contentItem, mouse.x, mouse.y)
                                        rootWindow.formatDragGhostY = p.y - (18 * scaleFactor)
                                        var pList = parent.mapToItem(formatRowsListView, mouse.x, mouse.y)
                                        var edge = 18 * scaleFactor
                                        var speed = 4 * scaleFactor
                                        if (pList.y < edge) {
                                            formatRowsListView.contentY = Math.max(0, formatRowsListView.contentY - speed)
                                        } else if (pList.y > (formatRowsListView.height - edge)) {
                                            var maxY = Math.max(0, formatRowsListView.contentHeight - formatRowsListView.height)
                                            formatRowsListView.contentY = Math.min(maxY, formatRowsListView.contentY + speed)
                                        }
                                    }
                                    onReleased: function(mouse) {
                                        rootWindow.formatDragGhostVisible = false
                                        var p = parent.mapToItem(rootWindow.contentItem, mouse.x, mouse.y)
                                        var deltaY = (p.y - dragStartGlobalY) + (formatRowsListView.contentY - dragStartContentY)
                                        var step = (50 * scaleFactor) + formatRowsListView.spacing
                                        var shift = Math.round(deltaY / Math.max(1, step))
                                        if (shift === 0) {
                                            return
                                        }
                                        var targetIndex = Math.max(0, Math.min(formatRowsListView.count - 1, index + shift))
                                        rootWindow.prepareFormatRowsModelChange(false)
                                        var movedIndex = backendSafe.moveFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, targetIndex)
                                        rootWindow.markAndJumpFormatRow(movedIndex)
                                    }
                                    onCanceled: {
                                        rootWindow.formatDragGhostVisible = false
                                    }
                                }
                            }

                            PixelTextField {
                                id: colField
                                Layout.preferredWidth: 48 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                Layout.fillWidth: activeFocus
                                text: modelData.col
                                placeholderText: ""
                                visible: true
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: activeFocus ? 3 : 0
                                font.pixelSize: activeFocus ? 12 * scaleFactor : 11 * scaleFactor
                                property bool submittedByEnter: false
                                property string editStartCol: String(modelData.col || "A")
                                onTextEdited: {
                                    var hadInvalid = /[^A-Za-z]/.test(text)
                                    var cleaned = text.replace(/[^A-Za-z]/g, "").toUpperCase()
                                    if (cleaned !== text) {
                                        text = cleaned
                                        if (hadInvalid) {
                                            rootWindow.showColumnLettersOnlyPopup(colField)
                                        }
                                    }
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "column"
                                        editStartCol = String(modelData.col || "A")
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: {
                                    if (submittedByEnter) {
                                        submittedByEnter = false
                                        return
                                    }
                                    var normalized = text.replace(/[^A-Za-z]/g, "").toUpperCase()
                                    if (normalized.length === 0) {
                                        text = String(modelData.col || "A")
                                        return
                                    }
                                    if (normalized === editStartCol) {
                                        text = normalized
                                        return
                                    }
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "col", text)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    editStartCol = normalized
                                }
                                onAccepted: {
                                    submittedByEnter = true
                                    var normalized = text.replace(/[^A-Za-z]/g, "").toUpperCase()
                                    if (normalized.length === 0) {
                                        text = String(modelData.col || "A")
                                        focus = false
                                        rootWindow.formatDesignerSelectedRowIndex = -1
                                        rootWindow.formatEditorFocusType = ""
                                        return
                                    }
                                    if (normalized === editStartCol) {
                                        focus = false
                                        rootWindow.formatDesignerSelectedRowIndex = -1
                                        return
                                    }
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "col", text)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    editStartCol = normalized
                                    focus = false
                                    rootWindow.formatDesignerSelectedRowIndex = -1
                                }
                                Keys.onEscapePressed: function(event) {
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            PixelComboBox {
                                id: typeCombo
                                Layout.preferredWidth: 82 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                Layout.fillWidth: activeFocus || (popup && popup.visible)
                                textPixelSize: 11 * scaleFactor
                                popupTextPixelSize: 11 * scaleFactor
                                skinYScale: 1.2
                                model: ["XML Data", "Formula", "Empty"]
                                currentIndex: {
                                    var t = (modelData && modelData.type) ? modelData.type : "data"
                                    return t === "formula" ? 1 : (t === "empty" ? 2 : 0)
                                }
                                visible: !rootWindow.formatCompactLabelsMode && (!rowExpanded || activeFocus || (popup && popup.visible))
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: (activeFocus || (popup && popup.visible)) ? 3 : 0
                                property bool sourceEditingActive: activeFocus || (popup && popup.visible)
                                fallbackNormal: themeInset
                                fallbackFocus: themeLayer1
                                fallbackOpen: themeLayer2
                                fallbackDisabled: themeLayer2
                                fallbackBorder: themeLayer2
                                fallbackText: themeText
                                fallbackPlaceholder: "white"
                                fallbackPopup: themePanel
                                onSourceEditingActiveChanged: {
                                    if (sourceEditingActive) {
                                        rootWindow.formatEditorFocusType = "source"
                                    } else if (rootWindow.formatEditorFocusType === "source") {
                                        rootWindow.formatEditorFocusType = ""
                                    }
                                }
                                onActivated: function(comboIndex) {
                                    if (formatCreatePanel.selectedFormatIndex < 0 || rowIndex < 0 || comboIndex < 0) {
                                        return
                                    }
                                    var mappedType = comboIndex === 1 ? "formula" : (comboIndex === 2 ? "empty" : "data")
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "type", mappedType)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    focus = false
                                }
                                Keys.onEscapePressed: function(event) {
                                    if (popup && popup.visible) {
                                        popup.close()
                                    }
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            PixelTextField {
                                id: valueField
                                Layout.fillWidth: true
                                Layout.preferredHeight: 34 * scaleFactor
                                text: modelData.value
                                enabled: !rootWindow.formatCompactLabelsMode && !formatCreatePanel.selectedBuiltInFormat && modelData.type !== "empty"
                                visible: !rootWindow.formatCompactLabelsMode && (!rowExpanded || activeFocus)
                                z: activeFocus ? 3 : 0
                                font.pixelSize: 11 * scaleFactor
                                placeholderText: ""
                                property bool submittedByEnter: false
                                hasError: {
                                    var raw = valueField.text ? valueField.text.trim() : ""
                                    return (modelData.type === "data" && valueField.enabled && raw.length > 0 && isNaN(parseInt(raw)))
                                        || (modelData.type === "formula" && valueField.enabled && raw.length > 0 && raw.charAt(0) !== "=")
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "value"
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: {
                                    if (submittedByEnter) {
                                        submittedByEnter = false
                                        return
                                    }
                                    var raw = text.trim()
                                    if (raw.length === 0) {
                                        text = String(modelData.value || "")
                                        return
                                    }
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "value", text)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                }
                                onAccepted: {
                                    submittedByEnter = true
                                    var raw = text.trim()
                                    if (raw.length === 0) {
                                        text = String(modelData.value || "")
                                        focus = false
                                        rootWindow.formatDesignerSelectedRowIndex = -1
                                        rootWindow.formatEditorFocusType = ""
                                        return
                                    }
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "value", text)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    focus = false
                                    rootWindow.formatDesignerSelectedRowIndex = -1
                                }
                                Keys.onEscapePressed: function(event) {
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            Rectangle {
                                id: sourceIndexButton
                                Layout.preferredWidth: 34 * scaleFactor
                                Layout.preferredHeight: 32 * scaleFactor
                                radius: 4
                                color: themePanel
                                border.width: 1
                                border.color: themeLayer2
                                visible: !formatCreatePanel.selectedBuiltInFormat
                                         && !rootWindow.formatCompactLabelsMode
                                         && modelData.type === "data"
                                         && valueField.activeFocus
                                enabled: backendSafe.xmlPreviewHeaders.length > 0

                                Text {
                                    anchors.centerIn: parent
                                    text: "ix"
                                    color: sourceIndexButton.enabled ? themeText : themeTextSecondary
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    anchors.fill: parent
                                    enabled: sourceIndexButton.enabled
                                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                                    onClicked: {
                                        rootWindow.formatDesignerSelectedRowIndex = rowIndex
                                        rootWindow.formatRowsHighlightIndex = rowIndex
                                        rootWindow.formatEditorFocusType = "value"
                                        rootWindow.openPreviewDialogForIndex(formatCreatePanel.selectedFormatIndex, rowIndex)
                                    }
                                }
                            }

                            Rectangle {
                                id: formulaTemplateButton
                                Layout.preferredWidth: 34 * scaleFactor
                                Layout.preferredHeight: 32 * scaleFactor
                                radius: 4
                                color: themePanel
                                border.width: 1
                                border.color: themeLayer2
                                visible: !formatCreatePanel.selectedBuiltInFormat
                                         && !rootWindow.formatCompactLabelsMode
                                         && modelData.type === "formula"
                                         && valueField.activeFocus

                                Text {
                                    anchors.centerIn: parent
                                    text: "fx"
                                    color: themeText
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: {
                                        formulaTemplateMenu.x = formulaTemplateButton.mapToItem(rootWindow.contentItem, 0, formulaTemplateButton.height).x
                                        formulaTemplateMenu.y = formulaTemplateButton.mapToItem(rootWindow.contentItem, 0, formulaTemplateButton.height).y
                                        formulaTemplateMenu.open()
                                    }
                                }

                                Menu {
                                    id: formulaTemplateMenu

                                    MenuItem {
                                        text: "=C{r}*280"
                                        onTriggered: applyFormulaTemplate("=C{r}*280")
                                    }
                                    MenuItem {
                                        text: "=(E{r}-E{r-1})*280/1000"
                                        onTriggered: applyFormulaTemplate("=(E{r}-E{r-1})*280/1000")
                                    }
                                    MenuItem {
                                        text: "=(E{r}-E{r-1})*1400/1000"
                                        onTriggered: applyFormulaTemplate("=(E{r}-E{r-1})*1400/1000")
                                    }
                                }
                            }

                            PixelTextField {
                                id: labelCustomField
                                Layout.preferredWidth: rootWindow.formatCompactLabelsMode ? 220 * scaleFactor : 200 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                Layout.fillWidth: rootWindow.formatCompactLabelsMode
                                text: rootWindow.customLabelTextFromKey(modelData.labelKey)
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                visible: rootWindow.formatCompactLabelsMode && rootWindow.customLabelIsCustom(modelData.labelKey)
                                z: activeFocus ? 3 : 0
                                font.pixelSize: 10 * scaleFactor
                                placeholderText: "Custom label"
                                onVisibleChanged: {
                                    if (!visible && activeFocus) {
                                        focus = false
                                    }
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "label"
                                    } else if (rootWindow.formatEditorFocusType === "label") {
                                        rootWindow.formatEditorFocusType = ""
                                    }
                                }
                                onEditingFinished: {
                                    var customKey = rootWindow.customLabelKeyFromText(text)
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "labelKey", customKey)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                }
                                onAccepted: {
                                    var customKey = rootWindow.customLabelKeyFromText(text)
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "labelKey", customKey)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    focus = false
                                }
                            }

                            PixelComboBox {
                                id: labelPresetCombo
                                Layout.preferredWidth: rootWindow.formatCompactLabelsMode ? 220 * scaleFactor : 200 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                Layout.fillWidth: rootWindow.formatCompactLabelsMode || activeFocus || (popup && popup.visible)
                                textPixelSize: 10 * scaleFactor
                                popupTextPixelSize: 10 * scaleFactor
                                popupMaxHeight: 220 * scaleFactor
                                skinYScale: 1.2
                                model: rootWindow.customLabelModel(modelData.type)
                                currentIndex: rootWindow.customLabelIndexForKey(modelData.labelKey, modelData.type)
                                visible: rootWindow.formatCompactLabelsMode
                                         && !rootWindow.customLabelIsCustom(modelData.labelKey)
                                         && (!rowExpanded || activeFocus || (popup && popup.visible))
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: (activeFocus || (popup && popup.visible)) ? 3 : 0
                                onVisibleChanged: {
                                    if (!visible) {
                                        if (popup && popup.visible) {
                                            popup.close()
                                        }
                                        if (activeFocus) {
                                            focus = false
                                        }
                                    }
                                }
                                fallbackNormal: themeInset
                                fallbackFocus: themeLayer1
                                fallbackOpen: themeLayer2
                                fallbackDisabled: themeLayer2
                                fallbackBorder: themeLayer2
                                fallbackText: themeText
                                fallbackPlaceholder: "white"
                                fallbackPopup: themePanel
                                onActiveFocusChanged: {
                                    if (activeFocus || (popup && popup.visible)) {
                                        rootWindow.formatEditorFocusType = "label"
                                    } else if (rootWindow.formatEditorFocusType === "label") {
                                        rootWindow.formatEditorFocusType = ""
                                    }
                                }
                                onActivated: function(comboIndex) {
                                    if (formatCreatePanel.selectedFormatIndex < 0 || rowIndex < 0 || comboIndex < 0) {
                                        return
                                    }
                                    var selectedKey = rootWindow.customLabelKeyForIndex(comboIndex, modelData.type)
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "labelKey", selectedKey)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    focus = false
                                }
                                Keys.onEscapePressed: function(event) {
                                    if (popup && popup.visible) {
                                        popup.close()
                                    }
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            PixelButton {
                                sliceLeft: 4
                                sliceRight: 4
                                sliceTop: 4
                                sliceBottom: 4
                                Layout.preferredWidth: 34 * scaleFactor
                                Layout.preferredHeight: 32 * scaleFactor
                                visible: rootWindow.formatCompactLabelsMode
                                text: "C"
                                textPixelSize: 10 * scaleFactor
                                fallbackNormal: themePanel
                                fallbackHover: themeLayer2
                                fallbackPressed: themeLayer1
                                borderColor: themeLayer2
                                onClicked: {
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var currentlyCustom = rootWindow.customLabelIsCustom(modelData.labelKey)
                                    var nextKey = currentlyCustom ? "" : "custom:Custom label"
                                    var updatedIndex = backendSafe.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "labelKey", nextKey)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    if (!currentlyCustom && labelCustomField) {
                                        labelCustomField.forceActiveFocus()
                                    }
                                }
                            }

                            PixelTextField {
                                id: widthField
                                Layout.preferredWidth: 64 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                Layout.fillWidth: activeFocus
                                text: String(modelData.width)
                                placeholderText: ""
                                inputMethodHints: Qt.ImhDigitsOnly
                                validator: IntValidator { bottom: 1; top: 200 }
                                visible: !rootWindow.formatCompactLabelsMode && (!rowExpanded || activeFocus)
                                enabled: !rootWindow.formatCompactLabelsMode && !formatCreatePanel.selectedBuiltInFormat
                                z: activeFocus ? 3 : 0
                                font.pixelSize: 11 * scaleFactor
                                property bool submittedByEnter: false
                                function commitWidthAndClose() {
                                    var raw = text.trim()
                                    if (raw.length === 0) {
                                        text = String(modelData.width || 14)
                                        rootWindow.markAndJumpFormatRow(index)
                                        submittedByEnter = true
                                        focus = false
                                        rootWindow.formatDesignerSelectedRowIndex = -1
                                        rootWindow.formatEditorFocusType = ""
                                        return
                                    }
                                    var widthValue = parseInt(raw)
                                    if (isNaN(widthValue)) {
                                        widthValue = 14
                                    }
                                    submittedByEnter = true
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "width", widthValue)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                    focus = false
                                    rootWindow.formatDesignerSelectedRowIndex = -1
                                }
                                onTextEdited: {
                                    var hadInvalid = /[^0-9]/.test(text)
                                    var cleaned = text.replace(/[^0-9]/g, "")
                                    if (cleaned !== text) {
                                        text = cleaned
                                        if (hadInvalid) {
                                            rootWindow.showWidthNumbersOnlyPopup(widthField)
                                        }
                                    }
                                }
                                Keys.onPressed: function(event) {
                                    if (event.modifiers & Qt.ControlModifier) {
                                        return
                                    }
                                    if (event.key === Qt.Key_Backspace
                                            || event.key === Qt.Key_Delete
                                            || event.key === Qt.Key_Left
                                            || event.key === Qt.Key_Right
                                            || event.key === Qt.Key_Home
                                            || event.key === Qt.Key_End
                                            || event.key === Qt.Key_Tab
                                            || event.key === Qt.Key_Return
                                            || event.key === Qt.Key_Enter
                                            || event.key === Qt.Key_Escape) {
                                        return
                                    }
                                    if (!event.text || event.text.length === 0) {
                                        return
                                    }
                                    if (!/[0-9]/.test(event.text)) {
                                        rootWindow.showWidthNumbersOnlyPopup(widthField)
                                        event.accepted = true
                                    }
                                }
                                Keys.onReturnPressed: function(event) {
                                    commitWidthAndClose()
                                    event.accepted = true
                                }
                                Keys.onEnterPressed: function(event) {
                                    commitWidthAndClose()
                                    event.accepted = true
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "width"
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: {
                                    if (submittedByEnter) {
                                        submittedByEnter = false
                                        return
                                    }
                                    var raw = text.trim()
                                    if (raw.length === 0) {
                                        text = String(modelData.width || 14)
                                        return
                                    }
                                    var widthValue = raw.length > 0 ? parseInt(raw) : 14
                                    if (isNaN(widthValue)) {
                                        widthValue = 14
                                    }
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    var updatedIndex = backendSafe.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "width", widthValue)
                                    rootWindow.markAndJumpFormatRow(updatedIndex)
                                }
                                onAccepted: {
                                    commitWidthAndClose()
                                }
                                Keys.onEscapePressed: function(event) {
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            PixelButton {
                                sliceLeft: 5
                                sliceRight: 5
                                sliceTop: 4
                                sliceBottom: 4
                                visible: !formatCreatePanel.selectedBuiltInFormat && !rowExpanded
                                Layout.preferredWidth: 50 * scaleFactor
                                Layout.preferredHeight: 34 * scaleFactor
                                skinYScaleWide: 1.0
                                skinYScaleCompact: 1.0
                                text: "Delete"
                                textPixelSize: 10 * scaleFactor
                                fallbackNormal: "#dc2626"
                                fallbackHover: "#b91c1c"
                                fallbackPressed: "#991b1b"
                                borderColor: "#dc2626"
                                onClicked: {
                                    rootWindow.prepareFormatRowsModelChange(false)
                                    rootWindow.markFormatRowHighlight(-1)
                                    backendSafe.deleteFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index)
                                }
                            }
                        }

                        TapHandler {
                            onTapped: {
                                if (rowExpanded) {
                                    rootWindow.formatDesignerSelectedRowIndex = index
                                } else {
                                    rootWindow.formatDesignerSelectedRowIndex = -1
                                }
                            }
                        }
                    }
                }
            }

            RowLayout {
                Layout.fillWidth: true
                spacing: 6 * scaleFactor

                PixelButton {
                    sliceLeft: 5
                    sliceRight: 5
                    sliceTop: 4
                    sliceBottom: 4
                    visible: !formatCreatePanel.selectedBuiltInFormat
                    Layout.fillWidth: true
                    Layout.preferredHeight: 38 * scaleFactor
                    text: "Save Formats"
                    textPixelSize: 12 * scaleFactor
                    fallbackNormal: themeLayer3
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    borderColor: themeLayer3
                    onClicked: {
                        backendSafe.renameFormatDefinition(rootWindow.formatDesignerSelectedFormatIndex, formatNameField.text)
                        backendSafe.saveFormatByName(rootWindow.formatDesignerSelectedFormatIndex)
                        backendSafe.commitFormatEdit()
                        processState = "formatDesigner"
                    }
                }

                PixelButton {
                    sliceLeft: 5
                    sliceRight: 5
                    sliceTop: 4
                    sliceBottom: 4
                    Layout.fillWidth: true
                    Layout.preferredHeight: 38 * scaleFactor
                    text: "Back To List"
                    textPixelSize: 12 * scaleFactor
                    fallbackNormal: themePanel
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    textColor: themeText
                    borderColor: themeLayer3
                    onClicked: {
                        if (formatCreatePanel.selectedBuiltInFormat) {
                            backendSafe.cancelFormatEdit()
                            processState = "formatDesigner"
                        } else {
                            rootWindow.openConfirmation("formatCreateDiscard", "Your work will be discarded. Continue?")
                        }
                    }
                }
            }
        }


        ColumnLayout {
            visible: processState === "converting"
            spacing: 20 * scaleFactor
            Layout.alignment: Qt.AlignHCenter
            Layout.fillWidth: true
            Layout.fillHeight: true

            Text {
                text: "CubeFlow"
                color: "white"
                font.family: appFontFamily
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Rectangle {
                visible: selectedFiles.length > 0
                Layout.fillWidth: true
                Layout.preferredHeight: 150 * scaleFactor
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 5
                clip: true

                ScrollView {
                    id: convertingBatchScroll
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    clip: true
                    ScrollBar.vertical.policy: ScrollBar.AsNeeded
                    contentWidth: availableWidth

                    Column {
                        id: convertingBatchContent
                        width: Math.max(0, convertingBatchScroll.availableWidth)
                        spacing: 5 * scaleFactor

                        Repeater {
                            model: selectedFiles.length

                            RowLayout {
                                width: convertingBatchScroll.availableWidth
                                spacing: 6 * scaleFactor
                                clip: true

                                Text {
                                    Layout.fillWidth: true
                                    text: baseName(selectedFiles[index])
                                    font.pixelSize: 11 * scaleFactor
                                    elide: Text.ElideMiddle
                                    wrapMode: Text.NoWrap
                                    clip: true
                                    color: themeText
                                }

                                Text {
                                    Layout.preferredWidth: 88 * scaleFactor
                                    Layout.maximumWidth: 88 * scaleFactor
                                    horizontalAlignment: Text.AlignRight
                                    text: (batchFileStatuses && batchFileStatuses.length > index) ? batchFileStatuses[index] : "Queued"
                                    font.pixelSize: 11 * scaleFactor
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                    clip: true
                                    color: text === "Done" ? "#059669"
                                          : text === "Failed" ? "#dc2626"
                                          : text === "Processing" ? themeLayer3
                                          : "#6b7280"
                                }
                            }
                        }
                    }
                }
            }

            Text {
                text: "Converting File"
                color: "white"
                font.pixelSize: 20 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                color: "white"
                text: "Please wait while we process your XML file..."
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: selectedFiles.length > 0
                    ? ("Processing: " + currentFileName)
                    : ("Using " + selectionType + " conversion")
                font.pixelSize: 12 * scaleFactor
                color: themeTextSecondary
                Layout.alignment: Qt.AlignHCenter
                Layout.fillWidth: true
                horizontalAlignment: Text.AlignHCenter
                elide: Text.ElideRight
                wrapMode: Text.NoWrap
            }

            Text {
                visible: selectedFiles.length > 0 && totalBatchFiles > 0
                text: "(" + (Math.min(totalBatchFiles, currentBatchIndex + 1)) + " of " + totalBatchFiles + ")"
                font.pixelSize: 12 * scaleFactor
                color: themeTextSecondary
                Layout.alignment: Qt.AlignHCenter
                Layout.fillWidth: true
                horizontalAlignment: Text.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64 * scaleFactor
                implicitHeight: 64 * scaleFactor
            }

            Item {
                Layout.fillWidth: true
                Layout.fillHeight: true
                Layout.minimumHeight: 0
            }
        }


        ColumnLayout {
            visible: processState === "creating"
            spacing: 20 * scaleFactor
            Layout.alignment: Qt.AlignHCenter
            Layout.fillWidth: true
            Layout.fillHeight: true

            Text {
                text: "Creating New Excel File"
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
                color: themeText
            }

            Text {
                text: isBatch
                    ? "Please wait, saving batch files..."
                    : "Almost done, generating your output file..."
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
                color: themeText
            }

            Rectangle {
                visible: isBatch && (batchOutputs.length > 0 || selectedFiles.length > 0)
                Layout.fillWidth: true
                Layout.preferredHeight: 150 * scaleFactor
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 5
                clip: true

                ScrollView {
                    id: creatingBatchScroll
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    clip: true
                    ScrollBar.vertical.policy: ScrollBar.AsNeeded
                    contentWidth: availableWidth

                    Column {
                        width: Math.max(0, creatingBatchScroll.availableWidth)
                        spacing: 5 * scaleFactor

                        Repeater {
                            model: batchOutputs.length > 0 ? batchOutputs.length : selectedFiles.length

                            RowLayout {
                                width: creatingBatchScroll.availableWidth
                                spacing: 6 * scaleFactor
                                clip: true

                                Text {
                                    Layout.fillWidth: true
                                    text: (batchOutputs.length > index && batchOutputs[index].sourceFile)
                                          ? String(batchOutputs[index].sourceFile)
                                          : baseName(selectedFiles[index])
                                    font.pixelSize: 11 * scaleFactor
                                    elide: Text.ElideMiddle
                                    wrapMode: Text.NoWrap
                                    clip: true
                                    color: themeText
                                }

                                Text {
                                    Layout.preferredWidth: 88 * scaleFactor
                                    Layout.maximumWidth: 88 * scaleFactor
                                    horizontalAlignment: Text.AlignRight
                                    text: {
                                        var total = batchOutputs.length > 0 ? batchOutputs.length : selectedFiles.length
                                        if (total <= 0) return "Queued"
                                        var saveIndex = Math.max(0, Math.min(total, Math.ceil((progress / 100.0) * total)))
                                        if (index < saveIndex) return "Saved"
                                        if (index === saveIndex && progress < 100) return "Saving"
                                        return "Queued"
                                    }
                                    font.pixelSize: 11 * scaleFactor
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                    clip: true
                                    color: text === "Saved" ? "#059669"
                                          : text === "Saving" ? themeLayer3
                                          : "#6b7280"
                                }
                            }
                        }
                    }
                }
            }

            Text {
                visible: isBatch && totalBatchFiles > 0 && batchOutputs.length > 0
                text: {
                    var nextIndex = Math.max(0, Math.min(totalBatchFiles - 1, currentBatchSaveCount))
                    var fileIdx = Math.max(0, Math.min(batchOutputs.length - 1, nextIndex))
                    var fileName = batchOutputs[fileIdx] && batchOutputs[fileIdx].sourceFile ? batchOutputs[fileIdx].sourceFile : ""
                    return "Saving: " + fileName
                }
                font.pixelSize: 12 * scaleFactor
                color: themeTextSecondary
                Layout.alignment: Qt.AlignHCenter
                Layout.fillWidth: true
                horizontalAlignment: Text.AlignHCenter
                elide: Text.ElideRight
                wrapMode: Text.NoWrap
            }

            Text {
                visible: isBatch && totalBatchFiles > 0
                text: {
                    var savedCount = Math.max(0, Math.min(totalBatchFiles, currentBatchSaveCount))
                    return "(" + savedCount + " of " + totalBatchFiles + ")"
                }
                font.pixelSize: 12 * scaleFactor
                color: themeTextSecondary
                Layout.alignment: Qt.AlignHCenter
                Layout.fillWidth: true
                horizontalAlignment: Text.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64 * scaleFactor
                implicitHeight: 64 * scaleFactor
            }

            Item {
                Layout.fillWidth: true
                Layout.fillHeight: true
                Layout.minimumHeight: 0
            }
        }

        ColumnLayout {
            visible: processState === "batchReview"
            spacing: 12 * scaleFactor
            Layout.fillWidth: true
            Layout.fillHeight: true
            Layout.leftMargin: 8 * scaleFactor
            Layout.rightMargin: 8 * scaleFactor
            Layout.topMargin: 4 * scaleFactor
            Layout.bottomMargin: 6 * scaleFactor

            Text {
                text: "CubeFlow"
                color: "white"
                font.family: appFontFamily
                font.pixelSize: 23 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Review Batch Output"
                color: "white"
                font.pixelSize: 17 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Set one folder for all files, or edit per-file paths."
                color: themeTextSecondary
                font.pixelSize: 11 * scaleFactor
                Layout.fillWidth: true
                wrapMode: Text.Wrap
            }

            Text {
                text: batchOutputs.length + " file(s) ready to save"
                color: themeText
                font.pixelSize: 11 * scaleFactor
                font.bold: true
            }

            Rectangle {
                Layout.fillWidth: true
                visible: hasBatchOutputConflictsInModel()
                color: "#3a1f1f"
                border.color: "#dc2626"
                border.width: 1
                radius: 5
                implicitHeight: conflictInfoColumn.implicitHeight + (12 * scaleFactor)

                ColumnLayout {
                    id: conflictInfoColumn
                    anchors.fill: parent
                    anchors.margins: 6 * scaleFactor
                    spacing: 4 * scaleFactor

                    Text {
                        text: "Conflicting output paths detected"
                        color: "#fecaca"
                        font.pixelSize: 11 * scaleFactor
                        font.bold: true
                        Layout.fillWidth: true
                        wrapMode: Text.Wrap
                    }

                    Text {
                        text: batchOutputConflictSummary(4)
                        color: "#fee2e2"
                        font.pixelSize: 10 * scaleFactor
                        Layout.fillWidth: true
                        wrapMode: Text.Wrap
                    }

                    Text {
                        text: "Change file name or save folder so each output path is unique."
                        color: "#fecaca"
                        font.pixelSize: 10 * scaleFactor
                        Layout.fillWidth: true
                        wrapMode: Text.Wrap
                    }
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: commonBatchDirLayout.implicitHeight + (12 * scaleFactor)
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 5
                clip: false

                ColumnLayout {
                    id: commonBatchDirLayout
                    anchors.fill: parent
                    anchors.margins: 6 * scaleFactor
                    spacing: 6 * scaleFactor

                    Text {
                        text: "Batch Save Folder"
                        color: themeText
                        font.pixelSize: 10 * scaleFactor
                        font.bold: true
                        Layout.fillWidth: true
                    }

                    RowLayout {
                        Layout.fillWidth: true
                        spacing: 6 * scaleFactor

                    ColumnLayout {
                        Layout.fillWidth: true
                        spacing: 2 * scaleFactor

                        Text {
                            text: "Save Folder"
                            color: themeText
                            font.pixelSize: 10 * scaleFactor
                            font.bold: true
                        }

                        PixelTextField {
                            id: allBatchSaveDirField
                            Layout.fillWidth: true
                            Layout.preferredHeight: batchControlHeight
                            font.pixelSize: batchOutputTextSize
                            normalSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                            focusSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                            errorSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                            disabledSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                            text: (batchOutputs && batchOutputs.length > 0) ? batchOutputs[0].saveDir : ""
                            property string dirValidationError: validateBatchSaveDir(text)
                            hasError: dirValidationError.length > 0
                            readOnly: true
                            selectionColor: "#bfdbfe"
                            color: themeText
                            selectedTextColor: themeText
                            fallbackNormal: themeInset
                            fallbackFocus: themeInset
                            fallbackDisabled: themeInset
                            onAccepted: {
                                var path = text.trim()
                                if (path.length > 0) {
                                    backendSafe.applyBatchOutputDirectoryToAll(path)
                                }
                            }
                        }

                        Text {
                            text: allBatchSaveDirField.dirValidationError
                            visible: allBatchSaveDirField.dirValidationError.length > 0
                            color: "#dc2626"
                            font.pixelSize: 10 * scaleFactor
                            font.bold: true
                            Layout.fillWidth: true
                            wrapMode: Text.Wrap
                        }
                    }

                    PixelButton {
                        sliceLeft: 5
                        sliceRight: 5
                        sliceTop: 4
                        sliceBottom: 4
                        Layout.preferredWidth: (compactBatchControls ? 84 : 64) * scaleFactor
                        Layout.preferredHeight: batchControlHeight
                        Layout.alignment: Qt.AlignBottom
                        skinYScaleWide: 1.0
                        skinYScaleCompact: 1.0
                        text: "Browse"
                        textPixelSize: (compactBatchControls ? 9 : 10) * scaleFactor
                        fallbackNormal: themeLayer3
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        borderColor: themeLayer3
                        onClicked: backendSafe.browseBatchOutputDirectoryForAll()
                    }

                    PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                        Layout.preferredWidth: (compactBatchControls ? 96 : 76) * scaleFactor
                        Layout.preferredHeight: batchControlHeight
                        Layout.alignment: Qt.AlignBottom
                        skinYScaleWide: 1.0
                        skinYScaleCompact: 1.0
                        text: "Apply All"
                        textPixelSize: (compactBatchControls ? 9 : 10) * scaleFactor
                        enabled: allBatchSaveDirField.text.trim().length > 0
                        fallbackNormal: themeLayer3
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        fallbackDisabled: "#9ca3af"
                        borderColor: themeLayer3
                        onClicked: backendSafe.applyBatchOutputDirectoryToAll(allBatchSaveDirField.text.trim())
                    }
                    }
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                Layout.minimumHeight: 0
                color: themePanel
                border.color: themeLayer2
                border.width: 1
                radius: 5
                clip: true

                ListView {
                    id: batchOutputsListView
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 8 * scaleFactor
                    clip: true
                    model: batchOutputs
                    onModelChanged: {
                        var keepY = batchOutputsSavedContentY
                        queueRestoreListContentY(batchOutputsListView, keepY)
                    }
                    onContentYChanged: {
                        if (!batchOutputsRestorePending) {
                            batchOutputsSavedContentY = contentY
                        }
                    }
                    onMovementEnded: {
                        batchOutputsSavedContentY = contentY
                    }

                    delegate: Rectangle {
                        width: ListView.view.width
                        color: themePanel
                        radius: 5
                        property bool batchRowEditing: outputFileNameField.activeFocus
                                                      || extCombo.activeFocus
                                                      || (extCombo.popup && extCombo.popup.visible)
                                                      || outputSaveDirField.activeFocus
                                                      || rootWindow.batchRowHighlightIndex === index
                        border.color: batchRowEditing ? "#dc2626" : themeLayer2
                        border.width: 1
                        implicitHeight: outputItemColumn.implicitHeight + (12 * scaleFactor)

                        ColumnLayout {
                            id: outputItemColumn
                            anchors.fill: parent
                            anchors.margins: 6 * scaleFactor
                            spacing: 6 * scaleFactor

                            Text {
                                text: (index + 1) + ". " + modelData.sourceFile
                                font.pixelSize: 11 * scaleFactor
                                font.bold: true
                                color: themeText
                                wrapMode: Text.Wrap
                            }

                            ColumnLayout {
                                Layout.fillWidth: true
                                spacing: 2 * scaleFactor

                                Text {
                                    text: "Output File Name"
                                    color: themeText
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                RowLayout {
                                    Layout.fillWidth: true
                                    spacing: 6 * scaleFactor

                                    PixelTextField {
                                        id: outputFileNameField
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: batchControlHeight
                                        font.pixelSize: batchOutputTextSize
                                        text: stripXlsx(modelData.fileName)
                                        property string validationError: validateBatchBaseName(text)
                                        hasError: validationError.length > 0
                                        selectionColor: "#bfdbfe"
                                        color: themeText
                                        selectedTextColor: themeText
                                        onTextEdited: {
                                            validationError = validateBatchBaseName(text)
                                            batchFileNameDrafts[index] = text
                                            batchFileNameDrafts = Object.assign({}, batchFileNameDrafts)
                                        }
                                        onEditingFinished: {
                                            batchFileNameDrafts[index] = text
                                            batchFileNameDrafts = Object.assign({}, batchFileNameDrafts)
                                            rootWindow.rememberBatchOutputsScroll()
                                            backendSafe.updateBatchOutputFileName(index, text + "." + extCombo.currentText)
                                        }
                                    }

                                    PixelComboBox {
                                        id: extCombo
                                        Layout.preferredWidth: (compactBatchControls ? 104 : 86) * scaleFactor
                                        Layout.preferredHeight: batchControlHeight
                                        textPixelSize: (compactBatchControls ? 9 : 10) * scaleFactor
                                        popupTextPixelSize: (compactBatchControls ? 9 : 10) * scaleFactor
                                        skinYScale: 1.0
                                        model: ["xlsx"]
                                        currentIndex: 0
                                        fallbackNormal: themeInset
                                        fallbackFocus: themeLayer1
                                        fallbackOpen: themeLayer2
                                        fallbackDisabled: themeLayer2
                                        fallbackBorder: themeLayer2
                                        fallbackText: themeText
                                        fallbackPlaceholder: "white"
                                        fallbackPopup: themePanel

                                        onCurrentTextChanged: {
                                            rootWindow.rememberBatchOutputsScroll()
                                            backendSafe.updateBatchOutputFileName(index, stripXlsx(modelData.fileName) + "." + currentText)
                                        }
                                    }
                                }

                                Text {
                                    text: outputFileNameField.validationError
                                    visible: outputFileNameField.validationError.length > 0
                                    color: "#dc2626"
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                    Layout.fillWidth: true
                                    wrapMode: Text.Wrap
                                }
                            }

                            RowLayout {
                                Layout.fillWidth: true
                                spacing: 6 * scaleFactor

                                ColumnLayout {
                                    Layout.fillWidth: true
                                    spacing: 2 * scaleFactor

                                    Text {
                                        text: "Save Folder"
                                        color: themeText
                                        font.pixelSize: 10 * scaleFactor
                                        font.bold: true
                                    }

                                    PixelTextField {
                                        id: outputSaveDirField
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: batchControlHeight
                                        font.pixelSize: batchOutputTextSize
                                        normalSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                                        focusSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                                        errorSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                                        disabledSource: Qt.resolvedUrl("images/ui/textbox_disabled.png")
                                        text: modelData.saveDir
                                        property string dirValidationError: validateBatchSaveDir(text)
                                        hasError: dirValidationError.length > 0
                                        readOnly: true
                                        selectionColor: "#bfdbfe"
                                        color: themeText
                                        selectedTextColor: themeText
                                        fallbackNormal: themeInset
                                        fallbackFocus: themeInset
                                        fallbackDisabled: themeInset
                                        onTextEdited: {
                                            rootWindow.rememberBatchOutputsScroll()
                                            backendSafe.updateBatchOutputDirectory(index, text)
                                        }
                                        onEditingFinished: {
                                            rootWindow.rememberBatchOutputsScroll()
                                            backendSafe.updateBatchOutputDirectory(index, text)
                                        }
                                    }

                                    Text {
                                        text: outputSaveDirField.dirValidationError
                                        visible: outputSaveDirField.dirValidationError.length > 0
                                        color: "#dc2626"
                                        font.pixelSize: 10 * scaleFactor
                                        font.bold: true
                                        Layout.fillWidth: true
                                        wrapMode: Text.Wrap
                                    }
                                }

                                PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                                    Layout.preferredWidth: (compactBatchControls ? 84 : 70) * scaleFactor
                                    Layout.preferredHeight: batchControlHeight
                                    Layout.alignment: Qt.AlignBottom
                                    skinYScaleWide: 1.0
                                    skinYScaleCompact: 1.0
                                    text: "Browse"
                                    textPixelSize: (compactBatchControls ? 9 : 10) * scaleFactor
                                    fallbackNormal: themeLayer3
                                    fallbackHover: themeLayer2
                                    fallbackPressed: themeLayer1
                                    borderColor: themeLayer3
                                    onClicked: {
                                        rootWindow.clearActiveEditorFocus()
                                        rootWindow.batchRowHighlightIndex = index
                                        backendSafe.browseBatchOutputDirectory(index)
                                    }
                                }
                            }
                        }
                    }

                    ScrollBar.vertical: ScrollBar {
                        policy: ScrollBar.AsNeeded
                    }
                }
            }

            Item {
                Layout.fillWidth: true
                Layout.preferredHeight: 40 * scaleFactor

                RowLayout {
                    anchors.fill: parent
                    spacing: 8 * scaleFactor

                    PixelButton {
            sliceLeft: 5
            sliceRight: 5
            sliceTop: 4
            sliceBottom: 4
                        Layout.fillWidth: true
                        Layout.preferredHeight: 40 * scaleFactor
                        text: "Back"
                        textPixelSize: 13 * scaleFactor
                        fallbackNormal: themePanel
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        textColor: themeText
                        borderColor: themeLayer3
                        onClicked: rootWindow.openConfirmation("batchReviewBack", "Go back and discard these batch output edits?")
                    }

                    PixelButton {
                        sliceLeft: 5
                        sliceRight: 5
                        sliceTop: 4
                        sliceBottom: 4
                        Layout.fillWidth: true
                        Layout.preferredHeight: 40 * scaleFactor
                        text: "Confirm"
                        textPixelSize: 13 * scaleFactor
                        enabled: batchOutputs.length > 0
                              && !hasInvalidBatchNamesInModel()
                              && !hasInvalidBatchSaveDirsInModel()
                              && !hasBatchOutputConflictsInModel()
                        fallbackNormal: themeLayer3
                        fallbackHover: themeLayer2
                        fallbackPressed: themeLayer1
                        fallbackDisabled: themeLayer2
                        borderColor: themeLayer3
                        onClicked: backendSafe.saveAllBatchOutputs()
                    }
                }
            }
        }


        CompleteView {
            visible: processState === "complete"
            scaleFactor: rootWindow.scaleFactor
            isBatch: rootWindow.isBatch
            totalBatchFiles: rootWindow.totalBatchFiles
            selectionType: rootWindow.selectionType
            completionDetailMessage: rootWindow.completionDetailMessage
            themeText: rootWindow.themeText
            themeTextSecondary: rootWindow.themeTextSecondary
            themeLayer3: rootWindow.themeLayer3
            themeLayer2: rootWindow.themeLayer2
            themeLayer1: rootWindow.themeLayer1
            backendSafe: rootWindow.backendSafe
        }
    }

    CancelBar {
        visibleWhenRunning: processState === "converting" || processState === "creating"
        scaleFactor: rootWindow.scaleFactor
        themePanel: rootWindow.themePanel
        themeLayer3: rootWindow.themeLayer3
        themeLayer2: rootWindow.themeLayer2
        themeLayer1: rootWindow.themeLayer1
        themeText: rootWindow.themeText
        onCancelRequested: rootWindow.openConfirmation("cancelProcess", "Cancel the current process?")
    }

    CompactHeader {
        visible: compactHeaderMode && processState !== "converting" && processState !== "creating" && processState !== "complete" && processState !== "formatDesigner" && processState !== "formatCreate"
        scaleFactor: rootWindow.scaleFactor
        compactGifSize: rootWindow.compactGifSize
        appFontFamily: appFontFamily
        themeTextSecondary: rootWindow.themeTextSecondary
    }

    BackButton {
        id: backButton
        visible: processState === "batchReview" || processState === "selecting"
        scaleFactor: rootWindow.scaleFactor
        themeLayer3: rootWindow.themeLayer3
        onClicked: rootWindow.openConfirmation("headerBack", "Go back to the previous screen?")
    }

    Item {
        visible: rootWindow.formatDragGhostVisible
        z: 500
        anchors.left: parent.left
        anchors.right: parent.right
        y: rootWindow.formatDragGhostY
        height: 36 * scaleFactor

        Rectangle {
            anchors.horizontalCenter: parent.horizontalCenter
            width: Math.max(140 * scaleFactor, (dragGhostLabel.implicitWidth + (24 * scaleFactor)))
            height: parent.height
            radius: 6
            color: themePanel
            border.color: themeLayer3
            border.width: 1
            opacity: 0.95
            scale: 1.03

            Text {
                id: dragGhostLabel
                anchors.centerIn: parent
                text: rootWindow.formatDragGhostText
                color: themeText
                font.pixelSize: 11 * scaleFactor
                font.bold: true
                elide: Text.ElideRight
            }
        }
    }

    Popup {
        id: infoDialog
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
        width: Math.min(rootWindow.width - (32 * scaleFactor), 340 * scaleFactor)
        height: 150 * scaleFactor
        padding: 0
        z: 351
        background: Item {}

        Rectangle {
            anchors.fill: parent
            color: themePanel
            border.color: themeLayer2
            border.width: 1
            radius: 6

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 12 * scaleFactor
                spacing: 12 * scaleFactor

                Text {
                    text: "Notice"
                    color: themeText
                    font.family: appFontFamily
                    font.pixelSize: 14 * scaleFactor
                    font.bold: true
                    Layout.fillWidth: true
                    horizontalAlignment: Text.AlignHCenter
                }

                Text {
                    text: rootWindow.infoMessage
                    color: themeText
                    font.family: appFontFamily
                    font.pixelSize: 11 * scaleFactor
                    wrapMode: Text.Wrap
                    Layout.fillWidth: true
                    horizontalAlignment: Text.AlignHCenter
                }

                PixelButton {
                    sliceLeft: 5
                    sliceRight: 5
                    sliceTop: 4
                    sliceBottom: 4
                    Layout.fillWidth: true
                    Layout.preferredHeight: 36 * scaleFactor
                    text: "OK"
                    textPixelSize: 11 * scaleFactor
                    fallbackNormal: themeLayer3
                    fallbackHover: themeLayer2
                    fallbackPressed: themeLayer1
                    textColor: themeText
                    borderColor: themeLayer3
                    onClicked: infoDialog.close()
                }
            }
        }
    }

    Connections {
        target: backendSafe
        ignoreUnknownSignals: true
        function onFormatImportNotice(message) {
            rootWindow.openInfoDialog(message)
        }
        function onInAppNotice(message) {
            rootWindow.openInfoDialog(message)
        }
        function onInAppConfirmRequested(token, title, message) {
            rootWindow.openBackendConfirmation(token, title, message)
        }
    }
}

