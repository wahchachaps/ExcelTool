/*
This is a UI file (.ui.qml) that is intended to be edited in Qt Design Studio only.
It is supposed to be strictly declarative and only uses a subset of QML. If you edit
this file manually, you might introduce QML code that is not supported by Qt Design Studio.
Check out https://doc.qt.io/qtcreator/creator-quick-ui-forms.html for details on .ui.qml files.
*/
import QtQuick
import QtQuick.Controls
import QtQuick.Controls.Material  // Added for Material style customizations (e.g., ComboBox)
import QtQuick.Layouts
import QtQuick.Window  // Added for Window
import QtMultimedia

Window {  // Changed from Rectangle to Window for top-level display
    visible: true  // Ensure it's visible
    width: processState === "selecting" ? 600 : 600  // Auto-expand width on file selection
    height: processState === "selecting" ? 700 : 600  // Auto-expand height on file selection
    minimumWidth: 500  // Prevent shrinking below this
    minimumHeight: 600  // Prevent shrinking below this
    title: "ExcelTool"  // Added title for the window
    flags: Qt.Window | Qt.WindowTitleHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint  // Added MaximizeButtonHint for resizing
    onProcessStateChanged: requestActivate()  // Bring window to front when state changes

    property string processState: "idle"
    property string selectedFile: ""
    property string selectionType: ""
    property int progress: 0
    property string fileSize: ""  // Added: For displaying actual file size

    onSelectionTypeChanged: {
        var idx = typeComboBox.model.indexOf(selectionType);
        if (idx >= 0 && typeComboBox.currentIndex !== idx) {
            typeComboBox.currentIndex = idx;
        }
    }

    // Animations for progress
    NumberAnimation {
        id: convertAnimation
        target: this  // Changed from 'root' to 'this'
        property: "progress"
        from: 0
        to: 100
        duration: 2000 // 2 seconds for simulation
        onStopped: {
            processState = "creating"
            progress = 0
            createAnimation.start()
        }
    }

    NumberAnimation {
        id: createAnimation
        target: this  // Changed from 'root' to 'this'
        property: "progress"
        from: 0
        to: 100
        duration: 1500 // 1.5 seconds for simulation
        onStopped: {
            processState = "complete"
        }
    }

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: 20
        spacing: 20

        // Header
        ColumnLayout {
            Layout.alignment: Qt.AlignHCenter
            spacing: 10

            AnimatedImage {
                id: copywriting
                Layout.preferredWidth: 234
                Layout.preferredHeight: 134
                source: "images/copywriting.gif"
                speed: 0.4724
                fillMode: Image.PreserveAspectFit
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: qsTr("ExcelTool")
                font.styleName: "Bold"
                font.family: "Tahoma"
                font.pointSize: 38
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                color: "#e6000000"
                text: "Convert your XML files with ease"
                font.pixelSize: 20
                font.family: "Verdana"
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                color: "#bd000000"
                text: "by wahchachaps"
                font.pixelSize: 15
                font.family: "Verdana"
                Layout.alignment: Qt.AlignHCenter
            }
        }

        // Idle State: Styled File Selection (Inspired by React Code)
        Rectangle {
            Layout.fillWidth: true
            Layout.preferredHeight: 200
            color: "#f0f0f0"
            border.color: "#cccccc"
            border.width: 2
            radius: 10
            visible: processState === "idle"

            // MouseArea for click interaction (simulates React's input click)
            MouseArea {
                anchors.fill: parent
                cursorShape: Qt.PointingHandCursor
                onClicked: backend.selectFile()  // Triggers Python file dialog
            }

            // Hover effect (changes border and background on hover, like React)
            Rectangle {
                anchors.fill: parent
                color: parent.hovered ? "#e0e7ff" : "transparent"  // Light blue on hover
                border.color: parent.hovered ? "#4f46e5" : "#cccccc"  // Indigo border on hover
                border.width: 2
                radius: 10
                opacity: 0.5  // Subtle overlay
            }

            ColumnLayout {
                anchors.centerIn: parent
                spacing: 10

                // Upload Icon (simulates React's Upload icon from lucide-react)
                Rectangle {
                    Layout.preferredWidth: 48
                    Layout.preferredHeight: 48
                    color: "transparent"
                    Layout.alignment: Qt.AlignHCenter

                    Text {
                        text: "⬆"  // Simple upload arrow icon (replace with Image if you have an icon file)
                        font.pixelSize: 48
                        color: "#9ca3af"  // Gray color
                        anchors.centerIn: parent
                    }
                }

                Text {
                    text: "Click to select XML file"
                    font.pixelSize: 16
                    font.bold: true
                    Layout.alignment: Qt.AlignHCenter
                }

                Text {
                    text: "XML files only"
                    font.pixelSize: 12
                    color: "#666666"
                    Layout.alignment: Qt.AlignHCenter
                }
            }
        }

        // Selecting State: File Info and Type Selection
        ColumnLayout {
            visible: processState === "selecting"
            spacing: 20

            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 80
                color: "#f9f9f9"
                border.color: "#dddddd"
                border.width: 1
                radius: 5

                Text {
                    anchors.centerIn: parent
                    text: "File: " + (selectedFile ? selectedFile.split('/').pop() : "") + "\nSize: " + fileSize  // Updated: Use actual fileSize
                    font.pixelSize: 14
                }
            }

            Text {
                text: "Select Conversion Type"
                font.pixelSize: 16
                font.bold: true
            }

            // Replaced RadioButtons with a styled ComboBox
            ComboBox {
                id: typeComboBox
                Layout.fillWidth: true
                Layout.preferredHeight: 40
                model: ["Den", "Glacier", "Globe"]  // Only actual options in model
                currentIndex: -1  // No selection initially
                displayText: currentIndex === -1 ? "Select XML type" : currentText  // Show placeholder when no selection
                onCurrentTextChanged: {
                    if (currentIndex >= 0 && currentText !== selectionType) {  // Only update if a valid option is selected
                        backend.setSelectionType(currentText);
                    }
                }

                // Oval-shaped background
                background: Rectangle {
                    color: "#ffffff"
                    border.color: "#cccccc"
                    border.width: 1
                    radius: height / 2  // Makes it oval
                }

                // Content item styling
                contentItem: Text {
                    text: typeComboBox.displayText  // Use displayText to show placeholder or selected text
                    font.pixelSize: 14
                    color: typeComboBox.currentIndex === -1 ? "#999999" : "#333333"  // Gray for placeholder, black for selected
                    verticalAlignment: Text.AlignVCenter
                    leftPadding: 15
                    rightPadding: 15
                }

                // Delegate for dropdown items
                delegate: ItemDelegate {
                    width: typeComboBox.width
                    contentItem: Text {
                        text: modelData
                        font.pixelSize: 14
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

                // Indicator (dropdown arrow)
                indicator: Canvas {
                    id: canvas
                    x: typeComboBox.width - width - typeComboBox.rightPadding
                    y: typeComboBox.topPadding + (typeComboBox.availableHeight - height) / 2
                    width: 12
                    height: 8
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

                // Popup styling
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

            // Custom "Confirm and Convert" Button (Rectangle + Text + MouseArea)
            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 40
                color: selectionType !== "" ? "#4f46e5" : "#cccccc"  // Indigo when enabled, gray when disabled
                radius: 5
                border.color: "#4f46e5"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Confirm and Convert"
                    color: "white"
                    font.pixelSize: 14
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    enabled: selectionType !== ""
                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                    onClicked: backend.confirmAndConvert()  // Connect to Python
                }
            }

            // Custom "Select Different File" Button (Rectangle + Text + MouseArea)
            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 40
                color: "#f3f4f6"  // Light gray background
                radius: 5
                border.color: "#d1d5db"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Select Different File"
                    color: "#374151"  // Dark gray text
                    font.pixelSize: 14
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.selectDifferentFile()  // Connect to Python
                }
            }
        }

                // Converting State
        ColumnLayout {
            visible: processState === "converting"
            spacing: 20
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Converting File"
                font.pixelSize: 24
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Please wait while we process your XML file..."
                font.pixelSize: 14
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Using " + selectionType + " conversion"
                font.pixelSize: 12
                color: "#666666"
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64
                implicitHeight: 64
                Material.accent: "#4f46e5"  // Added: Matches button color
            }
        }

        // Creating State
        ColumnLayout {
            visible: processState === "creating"
            spacing: 20
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Creating New Excel File"
                font.pixelSize: 24
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Almost done, generating your output file..."
                font.pixelSize: 14
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64
                implicitHeight: 64
                Material.accent: "#4f46e5"  // Added: Matches button color
            }
        }

        // Creating State
        ColumnLayout {
            visible: processState === "creating"
            spacing: 20
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Creating New Excel File"
                font.pixelSize: 24
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Almost done, generating your output file..."
                font.pixelSize: 14
                Layout.alignment: Qt.AlignHCenter
            }

            BusyIndicator {
                Layout.alignment: Qt.AlignHCenter
                running: true
                implicitWidth: 64
                implicitHeight: 64
            }
        }

        // Complete State
        ColumnLayout {
            visible: processState === "complete"
            spacing: 20
            Layout.alignment: Qt.AlignHCenter

            Text {
                text: "Conversion Complete!"
                font.pixelSize: 24
                font.bold: true
                Layout.alignment: Qt.AlignHCenter
            }

            Text {
                text: "Your file has been successfully converted using " + selectionType
                font.pixelSize: 14
                Layout.alignment: Qt.AlignHCenter
            }

            // Custom "Convert Another File" Button (Rectangle + Text + MouseArea)
            Rectangle {
                Layout.fillWidth: true
                Layout.preferredHeight: 40
                color: "#4f46e5"  // Indigo background
                radius: 5
                border.color: "#4f46e5"
                border.width: 1

                Text {
                    anchors.centerIn: parent
                    text: "Convert Another File"
                    color: "white"
                    font.pixelSize: 14
                    font.bold: true
                }

                MouseArea {
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.convertAnotherFile()  // Connect to Python
                }
            }
        }
    }
}