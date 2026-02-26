import QtQuick
import "../controls"

Item {
    property real scaleFactor: 1.0
    property bool visibleWhenRunning: false
    property color themePanel: "#3d3d4d"
    property color themeLayer3: "#7d7d8a"
    property color themeLayer2: "#6b6b7a"
    property color themeLayer1: "#52525e"
    property color themeText: "white"

    signal cancelRequested()

    visible: visibleWhenRunning
    anchors.left: parent.left
    anchors.right: parent.right
    anchors.bottom: parent.bottom
    anchors.leftMargin: 20 * scaleFactor
    anchors.rightMargin: 20 * scaleFactor
    anchors.bottomMargin: 16 * scaleFactor
    height: 40 * scaleFactor
    z: 120

    PixelButton {
        sliceLeft: 5
        sliceRight: 5
        sliceTop: 4
        sliceBottom: 4
        anchors.fill: parent
        text: "Cancel"
        textPixelSize: 12 * scaleFactor
        fallbackNormal: themePanel
        fallbackHover: themeLayer2
        fallbackPressed: themeLayer1
        textColor: themeText
        borderColor: themeLayer3
        onClicked: cancelRequested()
    }
}
