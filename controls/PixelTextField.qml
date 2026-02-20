import QtQuick
import QtQuick.Controls

TextField {
    id: root

    property bool hasError: false
    property color fallbackNormal: "#2b2b36"
    property color fallbackFocus: "#52525e"
    property color fallbackError: "#dc2626"
    property color fallbackDisabled: "#6b6b7a"
    property color fallbackBorder: "#6b6b7a"
    property color fallbackText: "white"
    property color fallbackPlaceholder: "#b8b8c4"
    property int sliceLeft: 8
    property int sliceRight: 8
    property int sliceTop: 8
    property int sliceBottom: 8

    property string normalSource: Qt.resolvedUrl("../images/ui/textbox_normal.png")
    property string focusSource: Qt.resolvedUrl("../images/ui/textbox_focus.png")
    property string errorSource: Qt.resolvedUrl("../images/ui/textbox_error.png")
    property string disabledSource: Qt.resolvedUrl("../images/ui/textbox_disabled.png")

    implicitHeight: 36
    color: fallbackText
    verticalAlignment: TextInput.AlignVCenter
    placeholderTextColor: fallbackPlaceholder
    selectionColor: "#7d7d8a"
    selectedTextColor: fallbackText
    leftPadding: 12
    rightPadding: 12
    topPadding: 6
    bottomPadding: 6

    readonly property string _currentSource: !enabled && disabledSource.length > 0
                                           ? disabledSource
                                           : (hasError && errorSource.length > 0
                                              ? errorSource
                                              : (activeFocus && focusSource.length > 0
                                                 ? focusSource
                                                 : normalSource))
    onActiveFocusChanged: {
        if (activeFocus) {
            cursorPosition = text.length
        }
    }

    background: Item {
        BorderImage {
            id: sprite
            anchors.fill: parent
            source: root._currentSource
            visible: root._currentSource.length > 0 && status === Image.Ready
            smooth: false
            border.left: root.sliceLeft
            border.right: root.sliceRight
            border.top: root.sliceTop
            border.bottom: root.sliceBottom
        }

        Rectangle {
            anchors.fill: parent
            visible: !sprite.visible
            radius: 4
            border.width: 1
            border.color: root.hasError ? root.fallbackError : root.fallbackBorder
            color: !root.enabled ? root.fallbackDisabled
                  : root.hasError ? root.fallbackError
                  : root.activeFocus ? root.fallbackFocus
                  : root.fallbackNormal
        }
    }
}
