#include <QApplication>
#include <QMainWindow>
#include <QPushButton>
#include <QTimer>
#include <QTranslator>

#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"

using namespace cc::neolux::fem::xlsxeditor;

int main(int argc, char* argv[]) {
    if (argc < 2) {
        qWarning("Usage: %s <xlsx-file> [--dry-run]", argv[0]);
        qWarning("  --real-delete: Enable real deletion mode (default is dry-run mode)");
        return -1;
    }
    std::string xlsxFile = argv[1];

    // Parse command line arguments for dry_run mode
    bool dryRun = false;
    bool autoTest = false;
    for (int i = 2; i < argc; ++i) {
        if (std::string(argv[i]) == "--dry-run") {
            dryRun = true;
        } else if (std::string(argv[i]) == "--auto-test") {
            autoTest = true;
        }
    }

    QApplication app(argc, argv);

    // Load translations
    QTranslator translator;
    if (translator.load("translations/XLSXEditor_zh_CN.qm")) {
        app.installTranslator(&translator);
    }

    // For testing, create a main window and embed XLSXEditor
    QMainWindow window;
    XLSXEditor* editor = new XLSXEditor(&window, dryRun);
    window.setCentralWidget(editor);
    window.resize(1024, 768);
    window.show();

    // Log mode
    if (dryRun) {
        qInfo() << "Running in DRY-RUN mode (fake delete with red marker)";
    } else {
        qInfo() << "Running in REAL-DELETE mode (will delete pictures and definitions)";
    }

    // 延迟加载，确保进度条可见
    QTimer::singleShot(0, editor, [editor, xlsxFile, autoTest]() {
        editor->loadXLSX(QString::fromStdString(xlsxFile), "SO13(DNo.3)MS", "B:K,7:34");

        if (autoTest) {
            qInfo() << "Auto-test mode: marking all items for deletion and saving...";
            QTimer::singleShot(500, editor, [editor]() {
                // Mark all items for deletion by toggling each one
                QList<QWidget*> items = editor->findChildren<QWidget*>("dataItem");
                qInfo() << "Found" << items.size() << "data items";
                for (auto* item : items) {
                    // Simulate clicking the delete button on each item
                    QList<QPushButton*> btns = item->findChildren<QPushButton*>();
                    if (!btns.isEmpty()) {
                        btns[0]->click();
                    }
                }

                // Save after marking
                QTimer::singleShot(500, editor, [editor]() {
                    qInfo() << "Triggering save...";
                    // Find and click the save button
                    QList<QPushButton*> saveBtns = editor->findChildren<QPushButton*>("btnSave");
                    if (!saveBtns.isEmpty()) {
                        saveBtns[0]->click();
                        qInfo() << "Save triggered";
                    }

                    // Exit after save
                    QTimer::singleShot(1000, []() {
                        qInfo() << "Auto-test complete, exiting...";
                        QApplication::quit();
                    });
                });
            });
        }
    });

    return app.exec();
}
