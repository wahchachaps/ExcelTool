/*
This is a UI file (.ui.qml) that is intended to be edited in Qt Design Studio only.
It is supposed to be strictly declarative and only uses a subset of QML. If you edit
this file manually, you might introduce QML code that is not supported by Qt Design Studio.
Check out https://doc.qt.io/qtcreator/creator-quick-ui-forms.html for details on .ui.qml files.
*/
import QtQuick
import QtQuick.Controls
import QtQuick.Controls.Material
import QtQuick.Layouts
import QtQuick.Window
import QtMultimedia

Window {
    id: rootWindow
    visible: true
    width: 400
    height: 500
    maximumHeight: 600
    minimumWidth: 400
    minimumHeight: 500
    title: "ExcelTool"
    flags: Qt.Window | Qt.WindowTitleHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint
    onProcessStateChanged: {
        requestActivate()
        if (processState !== "batchReview") {
            batchFileNameDrafts = ({})
        }
        if (processState !== "formatDesigner" && processState !== "formatCreate") {
            rootWindow.formatDesignerSelectedFormatIndex = 0
            rootWindow.formatDesignerSelectedRowIndex = -1
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
    property string currentFileName: ""
    property var batchOutputs: []
    property var batchFileNameDrafts: ({})
    property int formatDesignerSelectedFormatIndex: 0
    property int formatDesignerSelectedRowIndex: -1
    property string formatEditorFocusType: ""


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
    property real compactGifSize: Math.max(100, 100 * scaleFactor)
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
        return backend.validateOutputDirectory(String(path || ""))
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

    function batchFilesDescription() {
        if (!selectedFiles || selectedFiles.length === 0) {
            return "No files selected"
        }
        return selectedFiles.map(function(f) {
            return baseName(f) + " (" + backend.getFileSize(f) + ")"
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

                Button {
                    text: "OK"
                    Layout.alignment: Qt.AlignHCenter
                    onClicked: errorDialog.close()
                }
            }
        }
    }


    ColumnLayout {
        anchors.fill: parent
        anchors.leftMargin: 20 * scaleFactor
        anchors.rightMargin: 20 * scaleFactor
        anchors.bottomMargin: (processState === "selecting" ? 18 : 20) * scaleFactor
        anchors.topMargin: ((20 - (12 * headerCompress)) * scaleFactor) + headerReservedTopSpace
        spacing: (20 - (8 * headerCompress)) * scaleFactor


        ColumnLayout {
            Layout.alignment: Qt.AlignHCenter
            spacing: (10 - (4 * headerCompress)) * scaleFactor
            visible: processState !== "batchReview" && processState !== "converting" && processState !== "formatDesigner" && processState !== "formatCreate" && !compactHeaderMode

            AnimatedImage {
                id: copywriting
                Layout.preferredWidth: 200 * scaleFactor * headerScale
                Layout.preferredHeight: 110 * scaleFactor * headerScale
                source: "images/copywriting.gif"
                speed: 0.4724
                fillMode: Image.PreserveAspectFit
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: qsTr("ExcelTool")
                font.family: "Tahoma"
                font.pixelSize: 32 * scaleFactor * headerScale
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                color: "#e6000000"
                text: "Convert your XML files with ease"
                font.pixelSize: 16 * scaleFactor * headerScale
                font.family: "Verdana"
                Layout.alignment: Qt.AlignHCenter
                visible: !compactHeaderMode && headerCompress < 0.45
            }

            Text {
                color: "#bd000000"
                text: "by wahchachaps"
                font.pixelSize: 12 * scaleFactor * headerScale
                font.family: "Verdana"
                Layout.alignment: Qt.AlignHCenter
            }
        }

        Rectangle {
            Layout.fillWidth: true
            Layout.preferredHeight: 165 * scaleFactor
            color: "#f0f0f0"
            border.color: dropArea.containsDrag ? "#2563eb" : "#cccccc"
            border.width: dropArea.containsDrag ? 3 : 2
            radius: 10
            visible: processState === "idle"


            MouseArea {
                anchors.fill: parent
                cursorShape: Qt.PointingHandCursor
                onClicked: {
                    backend.selectFile()
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
                    backend.setDroppedPaths(droppedPaths)
                    selectionType = ""
                    typeComboBox.currentIndex = -1
                }
            }


            Rectangle {
                anchors.fill: parent
                color: dropArea.containsDrag ? "#cce7ff" : (parent.hovered ? "#e0e7ff" : "transparent")
                border.color: dropArea.containsDrag ? "#2563eb" : (parent.hovered ? "#2563eb" : "#cccccc")
                border.width: dropArea.containsDrag ? 3 : (parent.hovered ? 2 : 2)
                radius: 10
                opacity: dropArea.containsDrag ? 0.4 : (parent.hovered ? 0.5 : 0)
                Behavior on opacity { NumberAnimation { duration: 200 } }
                Behavior on color { ColorAnimation { duration: 200 } }
                Behavior on border.width { NumberAnimation { duration: 200 } }
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
                        source: "images/upload.png"
                        fillMode: Image.PreserveAspectFit
                        width: parent.width
                        height: parent.height
                        anchors.centerIn: parent
                    }
                }



                Text {
                    text: dropArea.containsDrag ? "Drop XML file(s) or folder here" : "Click to select XML file(s)"
                    font.pixelSize: 16 * scaleFactor
                    font.bold: true
                    Layout.alignment: Qt.AlignHCenter
                }

                Text {
                    text: "XML files only"
                    font.pixelSize: 12 * scaleFactor
                    color: "#666666"
                    Layout.alignment: Qt.AlignHCenter
                }
            }
        }

        Rectangle {
            Layout.fillWidth: true
            Layout.preferredHeight: 35 * scaleFactor
            color: "#2563eb"
            radius: 5
            border.color: "#2563eb"
            border.width: 1
            visible: processState === "idle"

            Text {
                anchors.centerIn: parent
                text: "Open Format Designer"
                color: "white"
                font.pixelSize: 12 * scaleFactor
                font.bold: true
            }

            MouseArea {
                anchors.fill: parent
                cursorShape: Qt.PointingHandCursor
                onClicked: backend.openFormatDesigner()
            }
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
                color: "#374151"
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
                    color: "#f9f9f9"
                    border.color: "#dddddd"
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
                                    color: "#ffffff"
                                    border.color: "#e5e7eb"
                                    border.width: 1
                                    radius: 4

                                    RowLayout {
                                        anchors.fill: parent
                                        anchors.margins: 4 * scaleFactor
                                        spacing: 6 * scaleFactor

                                        Text {
                                            id: fileNameText
                                            Layout.fillWidth: true
                                            text: baseName(selectedFiles[index]) + " (" + backend.getFileSize(selectedFiles[index]) + ")"
                                            font.pixelSize: 12 * scaleFactor
                                            color: "#333333"
                                            elide: Text.ElideMiddle
                                            verticalAlignment: Text.AlignVCenter
                                        }

                                        Text {
                                            Layout.preferredWidth: 66 * scaleFactor
                                            text: (batchFileStatuses && batchFileStatuses.length > index) ? batchFileStatuses[index] : "Queued"
                                            font.pixelSize: 11 * scaleFactor
                                            color: text === "Done" ? "#059669"
                                                  : text === "Failed" ? "#dc2626"
                                                  : text === "Processing" ? "#2563eb"
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
                                                onClicked: backend.removeSelectedFile(index)
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
                font.pixelSize: 14 * scaleFactor
                font.bold: true
            }


            ComboBox {
                id: typeComboBox
                Layout.fillWidth: true
                Layout.preferredHeight: 35 * scaleFactor
                model: backend.xmlTypeOptions
                currentIndex: -1
                displayText: currentIndex === -1 ? "Select XML type" : currentText
                onCurrentTextChanged: {
                    if (currentIndex >= 0 && currentText !== selectionType) {

                        backend.setSelectionType(currentText);
                    }
                }


                background: Rectangle {
                    color: "#ffffff"
                    border.color: "#cccccc"
                    border.width: 1
                    radius: height / 2
                }


                contentItem: Text {
                    text: typeComboBox.displayText
                    font.pixelSize: 12 * scaleFactor
                    color: typeComboBox.currentIndex === -1 ? "#999999" : "#333333"
                    verticalAlignment: Text.AlignVCenter
                    leftPadding: 15
                    rightPadding: 15
                }


                delegate: ItemDelegate {
                    width: typeComboBox.width
                    contentItem: Text {
                        text: modelData
                        font.pixelSize: 12 * scaleFactor
                        color: "#333333"
                        verticalAlignment: Text.AlignVCenter
                        leftPadding: 15
                    }
                    background: Rectangle {
                        color: highlighted ? "#e0e7ff" : "transparent"
                        border.color: highlighted ? "#2563eb" : "transparent"
                        border.width: 1
                    }
                }


                indicator: Canvas {
                    id: canvas
                    x: typeComboBox.width - width - typeComboBox.rightPadding
                    y: typeComboBox.topPadding + (typeComboBox.availableHeight - height) / 2
                    width: 10 * scaleFactor
                    height: 6 * scaleFactor
                    contextType: "2d"

                    Connections {
                        target: typeComboBox
                        function onPressedChanged() { canvas.requestPaint(); }
                    }

                    onPaint: {
                        context.reset();
                        context.moveTo(0, 0);
                        context.lineTo(width, 0);
                        context.lineTo(width / 2, height);
                        context.closePath();
                        context.fillStyle = "#666666";
                        context.fill();
                    }
                }


                popup: Popup {
                    y: typeComboBox.height - 1
                    width: typeComboBox.width
                    implicitHeight: contentItem.implicitHeight
                    padding: 1

                    contentItem: ListView {
                        clip: true
                        implicitHeight: contentHeight
                        model: typeComboBox.popup.visible ? typeComboBox.delegateModel : null
                        currentIndex: typeComboBox.highlightedIndex
                        ScrollIndicator.vertical: ScrollIndicator { }
                    }

                    background: Rectangle {
                        border.color: "#cccccc"
                        color: "#ffffff"
                        radius: 5
                    }
                }
            }


            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 35 * scaleFactor
                color: selectionType !== "" ? "#2563eb" : "#cccccc"
                radius: 5
                border.color: "#2563eb"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Confirm and Convert"
                    color: "white"
                    font.pixelSize: 12 * scaleFactor
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    enabled: selectionType !== ""
                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                    onClicked: {
                        if (typeComboBox.currentIndex >= 0) {
                            backend.setSelectionType(typeComboBox.currentText)
                        }
                        backend.confirmAndConvert()
                    }
                }
        }


            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 35 * scaleFactor
                color: "#2563eb"
                radius: 5
                border.color: "#2563eb"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Select Different File"
                    color: "white"
                    font.pixelSize: 12 * scaleFactor
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.selectFile()
                }
            }
        }


        ColumnLayout {
            visible: processState === "formatDesigner"
            spacing: 8 * scaleFactor
            Layout.fillWidth: true
            Layout.fillHeight: true

            Text {
                text: "Select Format"
                font.pixelSize: 22 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Click a format to edit it"
                color: "#64748b"
                font.pixelSize: 11 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Create Format"
                color: "#2563eb"
                font.pixelSize: 12 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: {
                        var newIndex = backend.createFormatDraft()
                        rootWindow.formatDesignerSelectedFormatIndex = Math.max(0, newIndex)
                        rootWindow.formatDesignerSelectedRowIndex = -1
                        processState = "formatCreate"
                    }
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                color: "#f8fafc"
                border.color: "#d1d5db"
                border.width: 1
                radius: 6
                clip: true

                ListView {
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 6 * scaleFactor
                    model: backend.formatModel
                    clip: true

                    delegate: Rectangle {
                        width: ListView.view.width
                        height: 44 * scaleFactor
                        radius: 5
                        property bool isBuiltInFormat: (modelData.name === "Den" || modelData.name === "Glacier" || modelData.name === "Globe")
                        color: rootWindow.formatDesignerSelectedFormatIndex === index ? "#eef2ff" : "white"
                        border.color: rootWindow.formatDesignerSelectedFormatIndex === index ? "#2563eb" : "#e5e7eb"
                        border.width: 1

                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 6 * scaleFactor
                            spacing: 6 * scaleFactor

                            Text {
                                Layout.fillWidth: true
                                text: modelData.name
                                color: "#111827"
                                font.pixelSize: 12 * scaleFactor
                                font.bold: true
                                elide: Text.ElideRight
                            }

                            Rectangle {
                                Layout.preferredWidth: 56 * scaleFactor
                                Layout.preferredHeight: 28 * scaleFactor
                                radius: 4
                                color: "#2563eb"
                                border.color: "#2563eb"
                                border.width: 1

                                Text {
                                    anchors.centerIn: parent
                                    text: isBuiltInFormat ? "Open" : "Edit"
                                    color: "white"
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: {
                                        rootWindow.formatDesignerSelectedFormatIndex = index
                                        rootWindow.formatDesignerSelectedRowIndex = -1
                                        backend.beginFormatEdit(index)
                                        processState = "formatCreate"
                                    }
                                }
                            }

                            Rectangle {
                                visible: !isBuiltInFormat
                                Layout.preferredWidth: 56 * scaleFactor
                                Layout.preferredHeight: 28 * scaleFactor
                                radius: 4
                                color: backend.formatModel.length > 1 ? "#dc2626" : "#9ca3af"
                                border.color: color
                                border.width: 1

                                Text {
                                    anchors.centerIn: parent
                                    text: "Delete"
                                    color: "white"
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    anchors.fill: parent
                                    enabled: backend.formatModel.length > 1
                                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                                    onClicked: {
                                        backend.deleteFormatDefinition(index)
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Text {
                Layout.fillWidth: true
                text: backend.formatDesignerStatus
                color: backend.formatDesignerStatus.indexOf("Failed") === 0 ? "#dc2626" : "#2563eb"
                font.pixelSize: 10 * scaleFactor
                wrapMode: Text.Wrap
                visible: backend.formatDesignerStatus.length > 0
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 38 * scaleFactor
                radius: 5
                color: "#2563eb"
                border.color: "#2563eb"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Open File"
                    color: "white"
                    font.pixelSize: 12 * scaleFactor
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.importFormatModelFromFile()
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 36 * scaleFactor
                radius: 5
                color: "#ffffff"
                border.color: "#2563eb"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Back"
                    color: "#1d4ed8"
                    font.pixelSize: 12 * scaleFactor
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.closeFormatDesigner()
                }
            }
        }

        ColumnLayout {
            id: formatCreatePanel
            visible: processState === "formatCreate"
            property int selectedFormatIndex: rootWindow.formatDesignerSelectedFormatIndex
            property bool selectedBuiltInFormat: (
                backend.formatModel.length > 0
                && selectedFormatIndex >= 0
                && selectedFormatIndex < backend.formatModel.length
                && (backend.formatModel[selectedFormatIndex].name === "Den"
                    || backend.formatModel[selectedFormatIndex].name === "Glacier"
                    || backend.formatModel[selectedFormatIndex].name === "Globe")
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
                font.pixelSize: 22 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            TextField {
                id: formatNameField
                Layout.fillWidth: true
                text: (backend.formatModel.length > 0 && rootWindow.formatDesignerSelectedFormatIndex >= 0 && rootWindow.formatDesignerSelectedFormatIndex < backend.formatModel.length)
                    ? backend.formatModel[rootWindow.formatDesignerSelectedFormatIndex].name
                    : ""
                placeholderText: "Format name"
                readOnly: formatCreatePanel.selectedBuiltInFormat
                onEditingFinished: {
                    if (!formatCreatePanel.selectedBuiltInFormat) {
                        backend.renameFormatDefinition(rootWindow.formatDesignerSelectedFormatIndex, text)
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

                Rectangle {
                    Layout.preferredWidth: 100 * scaleFactor
                    Layout.preferredHeight: 32 * scaleFactor
                    radius: 4
                    color: (backend.formatModel.length > 0 && !formatCreatePanel.selectedBuiltInFormat) ? "#2563eb" : "#9ca3af"
                    border.color: color
                    border.width: 1

                    Text {
                        anchors.centerIn: parent
                        text: "Add Column"
                        color: "white"
                        font.pixelSize: 10 * scaleFactor
                        font.bold: true
                    }

                    MouseArea {
                        anchors.fill: parent
                        enabled: backend.formatModel.length > 0 && !formatCreatePanel.selectedBuiltInFormat
                        cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                        onClicked: backend.addFormatRow(rootWindow.formatDesignerSelectedFormatIndex)
                    }
                }
            }

            Text {
                Layout.fillWidth: true
                text: "Tips: XML Data uses XML column index (0, 1, 2...). Formula should start with '=' and can use {r} and {r-1}."
                color: "#475569"
                font.pixelSize: 10 * scaleFactor
                wrapMode: Text.Wrap
            }

            RowLayout {
                Layout.fillWidth: true
                Layout.leftMargin: 14 * scaleFactor
                Layout.rightMargin: 14 * scaleFactor
                Layout.topMargin: 2 * scaleFactor
                Layout.bottomMargin: 2 * scaleFactor
                spacing: 6 * scaleFactor

                Text {
                    Layout.preferredWidth: 48 * scaleFactor
                    text: "Column"
                    color: "#334155"
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "column") ? 1.0 : 0.0
                }

                Text {
                    Layout.preferredWidth: 82 * scaleFactor
                    text: "Source"
                    color: "#334155"
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "source") ? 1.0 : 0.0
                }

                Text {
                    Layout.fillWidth: true
                    text: "Value"
                    color: "#334155"
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "value") ? 1.0 : 0.0
                }

                Text {
                    Layout.preferredWidth: 64 * scaleFactor
                    text: "Width"
                    color: "#334155"
                    font.pixelSize: 9 * scaleFactor
                    font.bold: true
                    horizontalAlignment: Text.AlignHCenter
                    elide: Text.ElideRight
                    opacity: (rootWindow.formatEditorFocusType === "" || rootWindow.formatEditorFocusType === "width") ? 1.0 : 0.0
                }

                Item {
                    Layout.preferredWidth: 50 * scaleFactor
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                color: "#f8fafc"
                border.color: "#d1d5db"
                border.width: 1
                radius: 6
                clip: true

                ListView {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 6 * scaleFactor
                    clip: true
                    model: (backend.formatModel.length > 0 && rootWindow.formatDesignerSelectedFormatIndex >= 0 && rootWindow.formatDesignerSelectedFormatIndex < backend.formatModel.length)
                        ? backend.formatModel[rootWindow.formatDesignerSelectedFormatIndex].columns
                        : []

                    delegate: Rectangle {
                        property int rowIndex: index
                        width: ListView.view.width
                        height: 50 * scaleFactor
                        radius: 5
                        function applyFormulaTemplate(templateText) {
                            valueField.text = templateText
                            backend.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "value", templateText)
                            valueField.forceActiveFocus()
                            valueField.cursorPosition = valueField.text.length
                        }
                        property bool rowExpanded: (
                            colField.activeFocus
                            || valueField.activeFocus
                            || widthField.activeFocus
                            || typeCombo.activeFocus
                            || (typeCombo.popup && typeCombo.popup.visible)
                        )
                        onRowExpandedChanged: {
                            if (!rowExpanded && rootWindow.formatDesignerSelectedRowIndex === index) {
                                rootWindow.formatDesignerSelectedRowIndex = -1
                            }
                            if (!rowExpanded) {
                                rootWindow.formatEditorFocusType = ""
                            }
                        }
                        color: rootWindow.formatDesignerSelectedRowIndex === index ? "#eef2ff" : "white"
                        border.color: rootWindow.formatDesignerSelectedRowIndex === index ? "#2563eb" : "#e5e7eb"
                        border.width: 1

                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 6 * scaleFactor
                            spacing: 6 * scaleFactor

                            TextField {
                                id: colField
                                Layout.preferredWidth: 48 * scaleFactor
                                Layout.preferredHeight: activeFocus ? 40 * scaleFactor : 32 * scaleFactor
                                Layout.fillWidth: activeFocus
                                text: modelData.col
                                placeholderText: ""
                                visible: !rowExpanded || activeFocus
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: activeFocus ? 3 : 0
                                font.pixelSize: activeFocus ? 12 * scaleFactor : 11 * scaleFactor
                                background: Rectangle {
                                    radius: 4
                                    color: "white"
                                    border.width: 1
                                    border.color: colField.activeFocus ? "#e91e63" : "#cbd5e1"
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "column"
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "col", text)
                                onAccepted: {
                                    backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "col", text)
                                    focus = false
                                    rootWindow.formatDesignerSelectedRowIndex = -1
                                }
                                Keys.onEscapePressed: function(event) {
                                    focus = false
                                    rootWindow.clearFormatCreateEditorFocus()
                                    event.accepted = true
                                }
                            }

                            ComboBox {
                                id: typeCombo
                                Layout.preferredWidth: 82 * scaleFactor
                                Layout.preferredHeight: (activeFocus || (popup && popup.visible)) ? 40 * scaleFactor : 32 * scaleFactor
                                Layout.fillWidth: activeFocus || (popup && popup.visible)
                                model: ["XML Data", "Formula", "Empty"]
                                currentIndex: {
                                    var t = (modelData && modelData.type) ? modelData.type : "data"
                                    return t === "formula" ? 1 : (t === "empty" ? 2 : 0)
                                }
                                visible: !rowExpanded || activeFocus || (popup && popup.visible)
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: (activeFocus || (popup && popup.visible)) ? 3 : 0
                                background: Rectangle {
                                    radius: 4
                                    color: "white"
                                    border.width: 1
                                    border.color: (typeCombo.activeFocus || (typeCombo.popup && typeCombo.popup.visible)) ? "#e91e63" : "#cbd5e1"
                                }
                                onActivated: function(comboIndex) {
                                    if (formatCreatePanel.selectedFormatIndex < 0 || rowIndex < 0 || comboIndex < 0) {
                                        return
                                    }
                                    var mappedType = comboIndex === 1 ? "formula" : (comboIndex === 2 ? "empty" : "data")
                                    backend.updateFormatRow(formatCreatePanel.selectedFormatIndex, rowIndex, "type", mappedType)
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

                            TextField {
                                id: valueField
                                Layout.fillWidth: true
                                Layout.preferredHeight: activeFocus ? 40 * scaleFactor : 32 * scaleFactor
                                text: modelData.value
                                enabled: !formatCreatePanel.selectedBuiltInFormat && modelData.type !== "empty"
                                visible: !rowExpanded || activeFocus
                                z: activeFocus ? 3 : 0
                                font.pixelSize: activeFocus ? 12 * scaleFactor : 11 * scaleFactor
                                placeholderText: ""
                                background: Rectangle {
                                    radius: 4
                                    color: "white"
                                    border.width: 1
                                    border.color: {
                                        var raw = valueField.text ? valueField.text.trim() : ""
                                        if (modelData.type === "data" && valueField.enabled && raw.length > 0 && isNaN(parseInt(raw))) {
                                            return "#dc2626"
                                        }
                                        if (modelData.type === "formula" && valueField.enabled && raw.length > 0 && raw.charAt(0) !== "=") {
                                    return "#ea580c"
                                        }
                                        return valueField.activeFocus ? "#e91e63" : "#cbd5e1"
                                    }
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "value"
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "value", text)
                                onAccepted: {
                                    backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "value", text)
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
                                id: formulaTemplateButton
                                Layout.preferredWidth: 34 * scaleFactor
                                Layout.preferredHeight: 32 * scaleFactor
                                radius: 4
                                color: "#f8fafc"
                                border.width: 1
                                border.color: "#cbd5e1"
                                visible: !formatCreatePanel.selectedBuiltInFormat
                                         && modelData.type === "formula"
                                         && valueField.activeFocus

                                Text {
                                    anchors.centerIn: parent
                                    text: "fx"
                                    color: "#334155"
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

                            TextField {
                                id: widthField
                                Layout.preferredWidth: 64 * scaleFactor
                                Layout.preferredHeight: activeFocus ? 40 * scaleFactor : 32 * scaleFactor
                                Layout.fillWidth: activeFocus
                                text: String(modelData.width)
                                placeholderText: ""
                                inputMethodHints: Qt.ImhDigitsOnly
                                validator: IntValidator { bottom: 1; top: 200 }
                                visible: !rowExpanded || activeFocus
                                enabled: !formatCreatePanel.selectedBuiltInFormat
                                z: activeFocus ? 3 : 0
                                font.pixelSize: activeFocus ? 12 * scaleFactor : 11 * scaleFactor
                                background: Rectangle {
                                    radius: 4
                                    color: "white"
                                    border.width: 1
                                    border.color: widthField.activeFocus ? "#e91e63" : "#cbd5e1"
                                }
                                onActiveFocusChanged: {
                                    if (activeFocus) {
                                        rootWindow.formatEditorFocusType = "width"
                                        cursorPosition = text.length
                                    }
                                }
                                onEditingFinished: {
                                    var raw = text.trim()
                                    var widthValue = raw.length > 0 ? parseInt(raw) : 14
                                    if (isNaN(widthValue)) {
                                        widthValue = 14
                                    }
                                    backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "width", widthValue)
                                }
                                onAccepted: {
                                    var raw = text.trim()
                                    var widthValue = raw.length > 0 ? parseInt(raw) : 14
                                    if (isNaN(widthValue)) {
                                        widthValue = 14
                                    }
                                    backend.updateFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index, "width", widthValue)
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
                                visible: !formatCreatePanel.selectedBuiltInFormat && !rowExpanded
                                Layout.preferredWidth: 50 * scaleFactor
                                Layout.preferredHeight: 32 * scaleFactor
                                radius: 4
                                color: "#dc2626"
                                border.color: "#dc2626"
                                border.width: 1

                                Text {
                                    anchors.centerIn: parent
                                    text: "Delete"
                                    color: "white"
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: backend.deleteFormatRow(rootWindow.formatDesignerSelectedFormatIndex, index)
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

                Rectangle {
                    visible: !formatCreatePanel.selectedBuiltInFormat
                    Layout.fillWidth: true
                    Layout.preferredHeight: 38 * scaleFactor
                    radius: 5
                    color: "#2563eb"
                    border.color: "#2563eb"
                    border.width: 1

                    Text {
                        anchors.centerIn: parent
                        text: "Save Formats"
                        color: "white"
                        font.pixelSize: 12 * scaleFactor
                        font.bold: true
                    }

                    MouseArea {
                        anchors.fill: parent
                        cursorShape: Qt.PointingHandCursor
                        onClicked: {
                            backend.renameFormatDefinition(rootWindow.formatDesignerSelectedFormatIndex, formatNameField.text)
                            backend.saveFormatByName(rootWindow.formatDesignerSelectedFormatIndex)
                            backend.commitFormatEdit()
                            processState = "formatDesigner"
                        }
                    }
                }

                Rectangle {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 38 * scaleFactor
                    radius: 5
                    color: "#ffffff"
                    border.color: "#2563eb"
                    border.width: 1

                    Text {
                        anchors.centerIn: parent
                        text: "Back To List"
                        color: "#1d4ed8"
                        font.pixelSize: 12 * scaleFactor
                        font.bold: true
                    }

                    MouseArea {
                        anchors.fill: parent
                        cursorShape: Qt.PointingHandCursor
                        onClicked: {
                            if (formatCreatePanel.selectedBuiltInFormat) {
                                backend.cancelFormatEdit()
                                processState = "formatDesigner"
                            } else if (backend.confirmDiscardFormatEdit()) {
                                backend.cancelFormatEdit()
                                processState = "formatDesigner"
                            }
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

            Text {
                text: "ExcelTool"
                font.family: "Tahoma"
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Rectangle {
                visible: selectedFiles.length > 0
                Layout.fillWidth: true
                Layout.preferredHeight: 150 * scaleFactor
                color: "#f9f9f9"
                border.color: "#dddddd"
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
                        width: Math.max(0, convertingBatchScroll.availableWidth)
                        spacing: 5 * scaleFactor

                        Repeater {
                            model: selectedFiles.length

                            RowLayout {
                                width: convertingBatchScroll.availableWidth
                                spacing: 6 * scaleFactor

                                Text {
                                    Layout.fillWidth: true
                                    text: baseName(selectedFiles[index])
                                    font.pixelSize: 11 * scaleFactor
                                    elide: Text.ElideMiddle
                                    color: "#374151"
                                }

                                Text {
                                    Layout.preferredWidth: 74 * scaleFactor
                                    horizontalAlignment: Text.AlignRight
                                    text: (batchFileStatuses && batchFileStatuses.length > index) ? batchFileStatuses[index] : "Queued"
                                    font.pixelSize: 11 * scaleFactor
                                    color: text === "Done" ? "#059669"
                                          : text === "Failed" ? "#dc2626"
                                          : text === "Processing" ? "#2563eb"
                                          : "#6b7280"
                                }
                            }
                        }
                    }
                }
            }

            Text {
                text: "Converting File"
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Please wait while we process your XML file..."
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: selectedFiles.length > 0
                    ? "Processing: " + currentFileName + " (" + (currentBatchIndex + 1) + " of " + totalBatchFiles + ")"
                    : "Using " + selectionType + " conversion"
                font.pixelSize: 12 * scaleFactor
                color: "#666666"
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64 * scaleFactor
                implicitHeight: 64 * scaleFactor
                Material.accent: "#2563eb"
            }
        }


        ColumnLayout {
            visible: processState === "creating"
            spacing: 20 * scaleFactor
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Creating New Excel File"
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: isBatch
                    ? "Please wait, saving batch files..."
                    : "Almost done, generating your output file..."
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64 * scaleFactor
                implicitHeight: 64 * scaleFactor
                Material.accent: "#2563eb"
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
                text: "ExcelTool"
                font.family: "Tahoma"
                font.pixelSize: 23 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Review Batch Output"
                font.pixelSize: 17 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Set one folder for all files, or edit per-file paths."
                color: "#8a8a8a"
                font.pixelSize: 11 * scaleFactor
                Layout.fillWidth: true
                wrapMode: Text.Wrap
            }

            Text {
                text: batchOutputs.length + " file(s) ready to save"
                color: "#4b5563"
                font.pixelSize: 11 * scaleFactor
                font.bold: true
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: commonBatchDirLayout.implicitHeight + (12 * scaleFactor)
                color: "#eef2ff"
                border.color: "#d1d5db"
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
                        color: "#475569"
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
                            color: "#475569"
                            font.pixelSize: 10 * scaleFactor
                            font.bold: true
                        }

                        TextField {
                            id: allBatchSaveDirField
                            Layout.fillWidth: true
                            text: (batchOutputs && batchOutputs.length > 0) ? batchOutputs[0].saveDir : ""
                            property string dirValidationError: validateBatchSaveDir(text)
                            readOnly: true
                            selectionColor: "#bfdbfe"
                            color: "#111827"
                            selectedTextColor: "#111827"
                            background: Rectangle {
                                radius: 4
                                color: "#f8fafc"
                                border.color: allBatchSaveDirField.dirValidationError.length > 0 ? "#dc2626" : "#d1d5db"
                                border.width: 1
                            }
                            onAccepted: {
                                var path = text.trim()
                                if (path.length > 0) {
                                    backend.applyBatchOutputDirectoryToAll(path)
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

                    Rectangle {
                        Layout.preferredWidth: 64 * scaleFactor
                        Layout.preferredHeight: 30 * scaleFactor
                        Layout.alignment: Qt.AlignBottom
                        radius: 5
                        color: "#2563eb"
                        border.color: "#2563eb"
                        border.width: 1

                        Text {
                            anchors.centerIn: parent
                            text: "Browse"
                            color: "white"
                            font.pixelSize: 11 * scaleFactor
                            font.bold: true
                        }

                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: backend.browseBatchOutputDirectoryForAll()
                        }
                    }

                    Rectangle {
                        Layout.preferredWidth: 76 * scaleFactor
                        Layout.preferredHeight: 30 * scaleFactor
                        Layout.alignment: Qt.AlignBottom
                        radius: 5
                        color: allBatchSaveDirField.text.trim().length > 0 ? "#2563eb" : "#9ca3af"
                        border.color: color
                        border.width: 1

                        Text {
                            anchors.centerIn: parent
                            text: "Apply All"
                            color: "white"
                            font.pixelSize: 11 * scaleFactor
                            font.bold: true
                        }

                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            enabled: allBatchSaveDirField.text.trim().length > 0
                            onClicked: backend.applyBatchOutputDirectoryToAll(allBatchSaveDirField.text.trim())
                        }
                    }
                    }
                }
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.fillHeight: true
                Layout.minimumHeight: 0
                color: "#f8fafc"
                border.color: "#dddddd"
                border.width: 1
                radius: 5
                clip: true

                ListView {
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    spacing: 8 * scaleFactor
                    clip: true
                    model: batchOutputs

                    delegate: Rectangle {
                        width: ListView.view.width
                        color: "#ffffff"
                        radius: 5
                        border.color: "#e5e7eb"
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
                                color: "#334155"
                                wrapMode: Text.Wrap
                            }

                            ColumnLayout {
                                Layout.fillWidth: true
                                spacing: 2 * scaleFactor

                                Text {
                                    text: "Output File Name"
                                    color: "#475569"
                                    font.pixelSize: 10 * scaleFactor
                                    font.bold: true
                                }

                                RowLayout {
                                    Layout.fillWidth: true
                                    spacing: 6 * scaleFactor

                                    TextField {
                                        id: outputFileNameField
                                        Layout.fillWidth: true
                                        text: stripXlsx(modelData.fileName)
                                        property string validationError: validateBatchBaseName(text)
                                        selectionColor: "#bfdbfe"
                                        color: "#111827"
                                        selectedTextColor: "#111827"
                                        background: Rectangle {
                                            radius: 4
                                            color: "white"
                                            border.color: outputFileNameField.validationError.length > 0 ? "#dc2626" : "#d1d5db"
                                            border.width: 1
                                        }
                                        onTextEdited: {
                                            validationError = validateBatchBaseName(text)
                                            batchFileNameDrafts[index] = text
                                            batchFileNameDrafts = Object.assign({}, batchFileNameDrafts)
                                        }
                                        onEditingFinished: {
                                            batchFileNameDrafts[index] = text
                                            batchFileNameDrafts = Object.assign({}, batchFileNameDrafts)
                                            backend.updateBatchOutputFileName(index, text + "." + extCombo.currentText)
                                        }
                                    }

                                    ComboBox {
                                        id: extCombo
                                        Layout.preferredWidth: 86 * scaleFactor
                                        Layout.preferredHeight: 32 * scaleFactor
                                        model: ["xlsx"]
                                        currentIndex: 0

                                        background: Rectangle {
                                            radius: 4
                                            color: "#ffffff"
                                            border.color: "#d1d5db"
                                            border.width: 1
                                        }

                                        contentItem: Text {
                                            text: extCombo.currentText
                                            color: "#111827"
                                            font.pixelSize: 11 * scaleFactor
                                            verticalAlignment: Text.AlignVCenter
                                            leftPadding: 10 * scaleFactor
                                            rightPadding: 6 * scaleFactor
                                        }

                                        onCurrentTextChanged: {
                                            backend.updateBatchOutputFileName(index, stripXlsx(modelData.fileName) + "." + currentText)
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
                                        color: "#475569"
                                        font.pixelSize: 10 * scaleFactor
                                        font.bold: true
                                    }

                                    TextField {
                                        id: outputSaveDirField
                                        Layout.fillWidth: true
                                        text: modelData.saveDir
                                        property string dirValidationError: validateBatchSaveDir(text)
                                        readOnly: true
                                        selectionColor: "#bfdbfe"
                                        color: "#111827"
                                        selectedTextColor: "#111827"
                                        background: Rectangle {
                                            radius: 4
                                            color: "#f8fafc"
                                            border.color: outputSaveDirField.dirValidationError.length > 0 ? "#dc2626" : "#d1d5db"
                                            border.width: 1
                                        }
                                        onTextEdited: backend.updateBatchOutputDirectory(index, text)
                                        onEditingFinished: backend.updateBatchOutputDirectory(index, text)
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

                                Rectangle {
                                    Layout.preferredWidth: 70 * scaleFactor
                                    Layout.preferredHeight: 32 * scaleFactor
                                    Layout.alignment: Qt.AlignBottom
                                    radius: 5
                                    color: "#2563eb"
                                    border.color: "#2563eb"
                                    border.width: 1

                                    Text {
                                        anchors.centerIn: parent
                                        text: "Browse"
                                        color: "white"
                                        font.pixelSize: 11 * scaleFactor
                                        font.bold: true
                                    }

                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: backend.browseBatchOutputDirectory(index)
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

                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 40 * scaleFactor
                        radius: 5
                        color: backActionArea.pressed ? "#eef2ff" : (backActionArea.containsMouse ? "#f8faff" : "white")
                        border.color: "#2563eb"
                        border.width: 1

                        Text {
                            anchors.centerIn: parent
                            text: "Back"
                            color: "#1d4ed8"
                            font.pixelSize: 13 * scaleFactor
                            font.bold: true
                        }

                        MouseArea {
                            id: backActionArea
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: backend.convertAnotherFile()
                        }
                    }

                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 40 * scaleFactor
                        color: (batchOutputs.length > 0 && !hasInvalidBatchNamesInModel() && !hasInvalidBatchSaveDirsInModel())
                            ? (confirmActionArea.pressed ? "#1d4ed8" : (confirmActionArea.containsMouse ? "#3b82f6" : "#2563eb"))
                            : "#cccccc"
                        radius: 5
                        border.color: (batchOutputs.length > 0 && !hasInvalidBatchNamesInModel() && !hasInvalidBatchSaveDirsInModel()) ? "#2563eb" : "#cccccc"
                        border.width: 1

                        Text {
                            anchors.centerIn: parent
                            text: "Confirm"
                            color: "white"
                            font.pixelSize: 13 * scaleFactor
                            font.bold: true
                        }

                        MouseArea {
                            id: confirmActionArea
                            anchors.fill: parent
                            enabled: batchOutputs.length > 0 && !hasInvalidBatchNamesInModel() && !hasInvalidBatchSaveDirsInModel()
                            hoverEnabled: true
                            cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                            onClicked: backend.saveAllBatchOutputs()
                        }
                    }
                }
            }
        }


        ColumnLayout {
            visible: processState === "complete"
            spacing: 20 * scaleFactor
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Conversion Complete!"
                font.pixelSize: 24 * scaleFactor
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }
            Text {
                text: isBatch
                    ? "Batch conversion complete! " + totalBatchFiles + " files processed successfully."
                    : "Your file has been successfully converted using " + selectionType
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 40 * scaleFactor
                color: "#2563eb"
                radius: 5
                border.color: "#2563eb"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Convert Another File"
                    color: "white"
                    font.pixelSize: 14 * scaleFactor
                    font.bold: true
                }
                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: {
                        backend.convertAnotherFile()
                    }
                }
            }
        }
    }

    Item {
        visible: compactHeaderMode && processState !== "converting" && processState !== "formatDesigner" && processState !== "formatCreate"
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.topMargin: 12 * scaleFactor
        anchors.leftMargin: 20 * scaleFactor
        width: (compactGifSize + (10 * scaleFactor) + (220 * scaleFactor))
        height: compactGifSize
        z: 90

        Row {
            anchors.fill: parent
            anchors.verticalCenter: parent.verticalCenter
            spacing: 6 * scaleFactor

            AnimatedImage {
                width: compactGifSize
                height: compactGifSize
                source: "images/copywriting.gif"
                speed: 0.4724
                fillMode: Image.PreserveAspectFit
                anchors.verticalCenter: parent.verticalCenter
            }

            Item {
                width: 220 * scaleFactor
                height: compactGifSize
                anchors.verticalCenter: parent.verticalCenter

                Column {
                    anchors.fill: parent
                    spacing: 0

                    Text {
                        width: parent.width
                        height: parent.height * 0.58
                        text: qsTr("ExcelTool")
                        font.family: "Tahoma"
                        font.pixelSize: Math.max(24, compactGifSize * 0.32)
                        font.bold: true
                        verticalAlignment: Text.AlignBottom
                        horizontalAlignment: Text.AlignLeft
                        elide: Text.ElideRight
                    }

                    Text {
                        width: parent.width
                        height: parent.height * 0.42
                        text: "by wahchachaps"
                        color: "#666666"
                        font.family: "Verdana"
                        font.pixelSize: Math.max(13, compactGifSize * 0.18)
                        verticalAlignment: Text.AlignTop
                        horizontalAlignment: Text.AlignLeft
                        elide: Text.ElideRight
                    }
                }
            }
        }
    }

    Item {
        id: backButton
        visible: processState === "batchReview" || processState === "selecting"
        z: 100
        width: 36 * scaleFactor
        height: 36 * scaleFactor
        anchors.left: parent.left
        anchors.top: parent.top
        anchors.leftMargin: 12 * scaleFactor
        anchors.topMargin: 12 * scaleFactor

        Text {
            anchors.centerIn: parent
            text: "\u2190"
            color: backButtonArea.pressed ? "#1e40af" : (backButtonArea.containsMouse ? "#2563eb" : "#1d4ed8")
            font.pixelSize: 18 * scaleFactor
            font.bold: true
        }

        MouseArea {
            id: backButtonArea
            anchors.fill: parent
            hoverEnabled: true
            cursorShape: Qt.PointingHandCursor
            onClicked: {
                if (processState === "selecting") {
                    backend.selectDifferentFile()
                } else {
                    backend.convertAnotherFile()
                }
            }
        }
    }
}
