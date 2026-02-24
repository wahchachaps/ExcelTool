import QtQuick

Item {
    property real scaleFactor: 1.0
    property color themeLayer3: "#7d7d8a"

    signal clicked()

    visible: false
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
        color: themeLayer3
        font.pixelSize: 18 * scaleFactor
        font.bold: true
    }

    MouseArea {
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        onClicked: parent.clicked()
    }
}
