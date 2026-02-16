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
    visible: true
    width: 400
    height: 500
    minimumWidth: 400
    minimumHeight: 500
    title: "ExcelTool"
    flags: Qt.Window | Qt.WindowTitleHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint
    onProcessStateChanged: requestActivate()

    property string processState: "idle"
    property string selectedFile: ""
    property string selectionType: ""
    property int progress: 0
    property string fileSize: ""
    property var selectedFiles: []
    property bool isBatch: false
    property int currentBatchIndex: 0
    property int totalBatchFiles: 0
    property string currentFileName: ""


    property real scaleFactor: Math.min(width / 400, height / 500)
    property real fileInfoBaseHeight: 60 * scaleFactor
    property real fileInfoMaxHeight: 220 * scaleFactor
    property real fileInfoDesiredHeight: fileInfoText.implicitHeight + (16 * scaleFactor)
    property real fileInfoCardHeight: Math.min(fileInfoMaxHeight, Math.max(fileInfoBaseHeight, fileInfoDesiredHeight))
    property real headerCompress: processState === "selecting"
                                 ? Math.max(0, Math.min(1, (fileInfoCardHeight - (90 * scaleFactor)) / (130 * scaleFactor)))
                                 : 0
    property real headerScale: 1.0 - (0.38 * headerCompress)

    function baseName(path) {
        var normalized = String(path).replace(/\\/g, "/")
        var parts = normalized.split("/")
        return parts.length ? parts[parts.length - 1] : normalized
    }

    function batchFilesDescription() {
        if (!selectedFiles || selectedFiles.length === 0) {
            return "No files selected"
        }
        return selectedFiles.map(function(f) {
            return baseName(f) + " (" + backend.getFileSize(f) + ")"
        }).join("\n")
    }

    onSelectionTypeChanged: {
        var idx = typeComboBox.model.indexOf(selectionType);
        if (idx >= 0 && typeComboBox.currentIndex !== idx) {
            typeComboBox.currentIndex = idx;
        }

        if (!isBatch) {
            selectedFiles = [];
            currentBatchIndex = 0;
            totalBatchFiles = 0;
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
        anchors.bottomMargin: 20 * scaleFactor
        anchors.topMargin: (20 - (8 * headerCompress)) * scaleFactor
        spacing: (20 - (6 * headerCompress)) * scaleFactor


        ColumnLayout {
            Layout.alignment: Qt.AlignHCenter
            spacing: (10 - (4 * headerCompress)) * scaleFactor

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
            border.color: dropArea.containsDrag ? "#4f46e5" : "#cccccc"
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
                    if (drop.hasUrls) {
                        var fileUrl = drop.urls[0]
                        var filePath = fileUrl.toString().replace("file:///", "")


                        if (filePath.toLowerCase().endsWith(".xml")) {
                            selectedFile = filePath
                            backend.setSelectedFile(filePath)
                            selectionType = ""
                            typeComboBox.currentIndex = -1
                            processState = "selecting"


                            fileSize = backend.getFileSize(filePath)
                        } else {

                            errorDialog.open()
                        }
                    }
                }
            }


            Rectangle {
                anchors.fill: parent
                color: dropArea.containsDrag ? "#cce7ff" : (parent.hovered ? "#e0e7ff" : "transparent")
                border.color: dropArea.containsDrag ? "#4f46e5" : (parent.hovered ? "#4f46e5" : "#cccccc")
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
                    text: dropArea.containsDrag ? "Drop XML file here" : "Click to select XML file(s)"
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


        ColumnLayout {
            visible: processState === "selecting"
            spacing: (8 - (2 * headerCompress)) * scaleFactor

            Rectangle {
                id: fileInfoCard
                Layout.fillWidth: true
                Layout.preferredHeight: fileInfoCardHeight
                color: "#f9f9f9"
                border.color: "#dddddd"
                border.width: 1
                radius: 5
                clip: true

                ScrollView {
                    id: fileInfoScroll
                    anchors.fill: parent
                    anchors.margins: 8 * scaleFactor
                    clip: true

                    Text {
                        id: fileInfoText
                        width: fileInfoScroll.availableWidth
                        text: isBatch
                            ? "Batch Mode: " + selectedFiles.length + " files selected\n" + batchFilesDescription()
                            : "File: " + (selectedFile ? baseName(selectedFile) : "") + "\nSize: " + fileSize
                        font.pixelSize: 12 * scaleFactor
                        wrapMode: Text.Wrap
                        elide: Text.ElideNone
                        lineHeight: 1.15
                        lineHeightMode: Text.ProportionalHeight
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
                model: ["Den", "Glacier", "Globe"]
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
                        border.color: highlighted ? "#4f46e5" : "transparent"
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
                color: selectionType !== "" ? "#4f46e5" : "#cccccc"
                radius: 5
                border.color: "#4f46e5"
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
                    onClicked: backend.confirmAndConvert()
                }
        }


            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 35 * scaleFactor
                Layout.bottomMargin: 12 * scaleFactor
                color: "#4f46e5"
                radius: 5
                border.color: "#4f46e5"
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
            visible: processState === "converting"
            spacing: 20 * scaleFactor
            Layout.alignment: Qt.AlignHCenter

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
                text: isBatch
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
                Material.accent: "#4f46e5"
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
                    ? "Saving: " + currentFileName + " (" + (currentBatchIndex + 1) + " of " + totalBatchFiles + ")"
                    : "Almost done, generating your output file..."
                font.pixelSize: 14 * scaleFactor
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64 * scaleFactor
                implicitHeight: 64 * scaleFactor
                Material.accent: "#4f46e5"
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
                color: "#4f46e5"
                radius: 5
                border.color: "#4f46e5"
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
                        isBatch = false
                        selectedFiles = []
                        currentBatchIndex = 0
                        totalBatchFiles = 0
                        processState = "idle"
                    }
                }
            }
        }
    }

    Rectangle {
        id: backButton
        visible: processState === "selecting" || processState === "complete"
        z: 100
        width: 36 * scaleFactor
        height: 36 * scaleFactor
        radius: width / 2
        anchors.left: parent.left
        anchors.top: parent.top
        anchors.leftMargin: 12 * scaleFactor
        anchors.topMargin: 12 * scaleFactor
        color: backButtonArea.pressed ? "#3f36bd" : (backButtonArea.containsMouse ? "#4f46e5" : "#4338ca")
        border.color: "#c7d2fe"
        border.width: 1

        Text {
            anchors.centerIn: parent
            text: "\u2190"
            color: "white"
            font.pixelSize: 18 * scaleFactor
            font.bold: true
        }

        MouseArea {
            id: backButtonArea
            anchors.fill: parent
            hoverEnabled: true
            cursorShape: Qt.PointingHandCursor
            onClicked: backend.convertAnotherFile()
        }
    }
}
