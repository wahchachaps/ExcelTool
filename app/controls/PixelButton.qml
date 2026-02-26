import QtQuick

Item {
    id: root

    property alias text: label.text
    property color textColor: "white"
    property real textPixelSize: 12

    property string normalSource: Qt.resolvedUrl("../images/ui/button_normal.png")
    property string hoverSource: Qt.resolvedUrl("../images/ui/button_hover.png")
    property string pressedSource: Qt.resolvedUrl("../images/ui/button_pressed.png")
    property string disabledSource: Qt.resolvedUrl("../images/ui/button_disabled.png")

    property color fallbackNormal: "#2563eb"
    property color fallbackHover: "#3b82f6"
    property color fallbackPressed: "#1d4ed8"
    property color fallbackDisabled: "#9ca3af"
    property color borderColor: "#2563eb"
    property int sliceLeft: 5
    property int sliceRight: 5
    property int sliceTop: 4
    property int sliceBottom: 4
    property real skinXScale: 1.0
    property real skinYScaleWide: 1.6
    property real skinYScaleCompact: 1.6
    property int compactWidthThreshold: 120

    signal clicked()

    implicitWidth: 220
    implicitHeight: 40

    readonly property bool _hovered: clickArea.containsMouse
    readonly property bool _pressed: clickArea.pressed
    readonly property bool _compact: width < compactWidthThreshold
    readonly property real _skinYScale: _compact ? skinYScaleCompact : skinYScaleWide
    readonly property int _maxHorzSlice: Math.max(0, Math.floor(width / 2) - 1)
    readonly property int _maxVertSlice: Math.max(0, Math.floor(height / 2) - 1)
    readonly property int _leftSlice: Math.min(sliceLeft, _maxHorzSlice)
    readonly property int _rightSlice: Math.min(sliceRight, _maxHorzSlice)
    readonly property int _topSlice: Math.min(sliceTop, _maxVertSlice)
    readonly property int _bottomSlice: Math.min(sliceBottom, _maxVertSlice)
    readonly property string _currentSource: !root.enabled && disabledSource.length > 0
                                           ? disabledSource
                                           : (_pressed && pressedSource.length > 0
                                              ? pressedSource
                                              : (_hovered && hoverSource.length > 0
                                                 ? hoverSource
                                                 : normalSource))

    BorderImage {
        id: sprite
        anchors.fill: parent
        source: root._currentSource
        visible: root._currentSource.length > 0 && status === Image.Ready
        smooth: false
        border.left: root._leftSlice
        border.right: root._rightSlice
        border.top: root._topSlice
        border.bottom: root._bottomSlice
        transform: Scale {
            origin.x: sprite.width / 2
            origin.y: sprite.height / 2
            xScale: root.skinXScale
            yScale: root._skinYScale
        }
    }

    Rectangle {
        anchors.fill: parent
        visible: !sprite.visible
        radius: 5
        border.width: 1
        border.color: root.borderColor
        color: !root.enabled ? root.fallbackDisabled
              : root._pressed ? root.fallbackPressed
              : root._hovered ? root.fallbackHover
              : root.fallbackNormal
        transform: Scale {
            origin.x: parent.width / 2
            origin.y: parent.height / 2
            xScale: root.skinXScale
            yScale: root._skinYScale
        }
    }

    Text {
        id: label
        anchors.centerIn: parent
        color: root.textColor
        font.family: appFontFamily
        font.pixelSize: root.textPixelSize
        font.bold: true
    }

    MouseArea {
        id: clickArea
        anchors.fill: parent
        enabled: root.enabled
        hoverEnabled: true
        cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: root.clicked()
    }
}

