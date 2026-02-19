import QtQuick
import QtQuick.Controls

ComboBox {
    id: root

    property color fallbackNormal: "#2b2b36"
    property color fallbackFocus: "#52525e"
    property color fallbackOpen: "#6b6b7a"
    property color fallbackDisabled: "#6b6b7a"
    property color fallbackBorder: "#6b6b7a"
    property color fallbackText: "white"
    property color fallbackPlaceholder: "white"
    property color fallbackPopup: "#3d3d4d"
    property real textPixelSize: 12
    property real popupTextPixelSize: 11
    property int textLeftInset: 20
    property int textRightInset: 12
    property int sliceLeft: 4
    property int sliceRight: 4
    property int sliceTop: 4
    property int sliceBottom: 4
    property real skinXScale: 1.0
    property real skinYScale: 1.6

    property string normalSource: Qt.resolvedUrl("../images/ui/combobox_normal.png")
    property string focusSource: Qt.resolvedUrl("../images/ui/combobox_focus.png")
    property string openSource: Qt.resolvedUrl("../images/ui/combobox_open.png")
    property string disabledSource: Qt.resolvedUrl("../images/ui/combobox_disabled.png")
    property string arrowSource: Qt.resolvedUrl("../images/ui/combobox_arrow.png")

    implicitHeight: 40

    readonly property bool _isOpen: popup && popup.visible
    readonly property string _currentSource: !enabled && disabledSource.length > 0
                                           ? disabledSource
                                           : (_isOpen && openSource.length > 0
                                              ? openSource
                                              : ((activeFocus || hovered) && focusSource.length > 0
                                                 ? focusSource
                                                 : normalSource))

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
            transform: Scale {
                origin.x: sprite.width / 2
                origin.y: sprite.height / 2
                xScale: root.skinXScale
                yScale: root.skinYScale
            }
        }

        Rectangle {
            anchors.fill: parent
            visible: !sprite.visible
            radius: 4
            border.width: 1
            border.color: root.fallbackBorder
            color: !root.enabled ? root.fallbackDisabled
                  : root._isOpen ? root.fallbackOpen
                  : (root.activeFocus || root.hovered) ? root.fallbackFocus
                  : root.fallbackNormal
            transform: Scale {
                origin.x: parent.width / 2
                origin.y: parent.height / 2
                xScale: root.skinXScale
                yScale: root.skinYScale
            }
        }
    }

    contentItem: Item {
        anchors.fill: parent
        clip: true

        Text {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            anchors.bottom: parent.bottom
            anchors.leftMargin: root.sliceLeft + root.textLeftInset
            anchors.rightMargin: root.sliceRight + root.textRightInset
            anchors.topMargin: root.sliceTop + 2
            anchors.bottomMargin: root.sliceBottom + 2
            text: root.displayText
            color: root.currentIndex === -1 ? root.fallbackPlaceholder : root.fallbackText
            font.family: appFontFamily
            font.pixelSize: root.textPixelSize
            horizontalAlignment: Text.AlignLeft
            verticalAlignment: Text.AlignVCenter
            elide: Text.ElideRight
        }
    }

    indicator: Item {
        width: 0
        height: 0
        visible: false
    }

    popup: Popup {
        y: root.height - 1
        width: root.width
        padding: 1
        contentItem: ListView {
            clip: true
            implicitHeight: contentHeight
            model: root.popup.visible ? root.delegateModel : null
            currentIndex: root.highlightedIndex
            ScrollIndicator.vertical: ScrollIndicator {}
        }
        background: Rectangle {
            color: root.fallbackPopup
            border.color: root.fallbackBorder
            border.width: 1
            radius: 4
        }
    }

    delegate: ItemDelegate {
        width: root.width
        contentItem: Text {
            text: modelData
            color: root.fallbackText
            font.family: appFontFamily
            font.pixelSize: root.popupTextPixelSize
            verticalAlignment: Text.AlignVCenter
            leftPadding: 10
            elide: Text.ElideRight
        }
        background: Rectangle {
            color: highlighted ? root.fallbackOpen : "transparent"
            border.color: highlighted ? root.fallbackBorder : "transparent"
            border.width: 1
        }
    }
}
