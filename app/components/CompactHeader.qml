import QtQuick

Item {
    property real scaleFactor: 1.0
    property real compactGifSize: 130
    property string appFontFamily: ""
    property color themeTextSecondary: "#b8b8c4"

    visible: false
    anchors.top: parent.top
    anchors.left: parent.left
    anchors.right: parent.right
    anchors.topMargin: 12 * scaleFactor
    anchors.leftMargin: 20 * scaleFactor
    anchors.rightMargin: 20 * scaleFactor
    height: compactGifSize
    z: 90

    Row {
        anchors.fill: parent
        anchors.verticalCenter: parent.verticalCenter
        spacing: 6 * scaleFactor

        AnimatedImage {
            id: compactCopywriting
            width: compactGifSize
            height: compactGifSize
            source: "../images/copywriting.gif"
            speed: 1.0
            fillMode: Image.PreserveAspectFit
            transformOrigin: Item.Center
            scale: 1.2
            anchors.verticalCenter: parent.verticalCenter
        }

        Item {
            width: Math.max(0, parent.width - compactCopywriting.width - (6 * scaleFactor))
            height: compactGifSize
            anchors.verticalCenter: parent.verticalCenter

            Column {
                anchors.fill: parent
                spacing: 0

                Text {
                    width: parent.width
                    height: parent.height * 0.58
                    text: qsTr("CubeFlow")
                    color: "white"
                    font.family: appFontFamily
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
                    color: themeTextSecondary
                    font.family: appFontFamily
                    font.pixelSize: Math.max(13, compactGifSize * 0.18)
                    verticalAlignment: Text.AlignTop
                    horizontalAlignment: Text.AlignLeft
                    elide: Text.ElideRight
                }
            }
        }
    }
}
