import QtQuick
import QtQuick.Layouts
import "../controls"

ColumnLayout {
    property real scaleFactor: 1.0
    property bool isBatch: false
    property int totalBatchFiles: 0
    property string selectionType: ""
    property string completionDetailMessage: ""
    property color themeText: "white"
    property color themeTextSecondary: "#b8b8c4"
    property color themeLayer3: "#7d7d8a"
    property color themeLayer2: "#6b6b7a"
    property color themeLayer1: "#52525e"
    property var backendSafe: null

    spacing: 20 * scaleFactor
    Layout.alignment: Qt.AlignHCenter
    Layout.fillWidth: true

    Text {
        text: "Conversion Complete!"
        font.pixelSize: 24 * scaleFactor
        font.bold: true
        Layout.alignment: Qt.AlignHCenter
        Layout.fillWidth: true
        horizontalAlignment: Text.AlignHCenter
        color: themeText
    }

    Text {
        text: isBatch
            ? "Batch conversion complete! " + totalBatchFiles + " files processed successfully."
            : "Your file has been successfully converted using " + selectionType
        font.pixelSize: 14 * scaleFactor
        Layout.alignment: Qt.AlignHCenter
        Layout.fillWidth: true
        horizontalAlignment: Text.AlignHCenter
        wrapMode: Text.Wrap
        color: themeText
    }

    Text {
        visible: String(completionDetailMessage || "").length > 0
        text: completionDetailMessage
        font.pixelSize: 12 * scaleFactor
        Layout.alignment: Qt.AlignHCenter
        Layout.fillWidth: true
        horizontalAlignment: Text.AlignHCenter
        wrapMode: Text.Wrap
        color: themeTextSecondary
    }

    PixelButton {
        sliceLeft: 5
        sliceRight: 5
        sliceTop: 4
        sliceBottom: 4
        Layout.fillWidth: true
        Layout.preferredHeight: 40 * scaleFactor
        text: "Convert Another File"
        textPixelSize: 14 * scaleFactor
        fallbackNormal: themeLayer3
        fallbackHover: themeLayer2
        fallbackPressed: themeLayer1
        textColor: themeText
        borderColor: themeLayer3
        onClicked: {
            if (backendSafe && backendSafe.convertAnotherFile) {
                backendSafe.convertAnotherFile()
            }
        }
    }
}
